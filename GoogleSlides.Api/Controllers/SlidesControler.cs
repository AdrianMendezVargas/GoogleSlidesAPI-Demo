using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Slides.v1;
using Google.Apis.Slides.v1.Data;
using Google.Apis.Util.Store;
using GoogleSlides.Api.Data;
using GoogleSlides.Api.Models;
using GoogleSlides.Api.Models.Domain;
using GoogleSlides.Api.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore.Sqlite.Query.Internal;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using Newtonsoft.Json;
using System.Net;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using static Google.Apis.Requests.BatchRequest;
using Request = Google.Apis.Slides.v1.Data.Request;

namespace GoogleSlides.Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SlidesControler : ControllerBase
    {
        private readonly DriveService _DriveService = null!;
        private readonly SlidesService _SlidesService = null!;
        private readonly SheetsService _SheetsService = null!;
        private readonly TemplateService _templateService = null!;
        private readonly IWebHostEnvironment _hostingEnvironment;

        public SlidesControler(
            DriveService driveService, 
            SlidesService slidesService, 
            SheetsService sheetsService, 
            TemplateService templateService, 
            IWebHostEnvironment hostingEnvironment)
        {

            _DriveService = driveService;
            _SlidesService = slidesService;
            _SheetsService = sheetsService;
            _templateService = templateService;
            _hostingEnvironment = hostingEnvironment;

        }

        [HttpPost]
        [ProducesResponseType(200, Type = typeof(string))]
        public async Task<IActionResult> CreateFromTemplate(CreateSlidesDeckFromTemplate req)
        {

            Console.WriteLine("Objeto Peticion: " + JsonConvert.SerializeObject(req));

            var originalPresentation = _DriveService.Files.Get(req.TemplateId).Execute();

            var copyPresentation = new Google.Apis.Drive.v3.Data.File
            {
                Name = req.PresentationName,
                //Parents =   //new string[] { "1HtwFuFKFe4XJH6IIc8bHKbJZoJ4Tmr8s" }
            };

            var copyRequest = _DriveService.Files.Copy(copyPresentation, req.TemplateId);

            string? copyPresentationId = (await copyRequest.ExecuteAsync())?.Id;

            try
            {

                Permission permission = new Permission();
                permission.Type = "user";
                permission.Role = "writer";
                permission.EmailAddress = req.ReciverEmail;

                // Create the permission request and execute it to share the file
                var request = _DriveService.Permissions.Create(permission, copyPresentationId);
                request.SendNotificationEmail = true;
                request.Execute();

                #region Comentario Transferir 
                //PermissionsResource.CreateRequest permissionRequest = _DriveService.Permissions.Create(permission, copyPresentationId);

                //// Set the transferOwnership parameter to true
                //permissionRequest.TransferOwnership = true;

                //// Execute the permission request and transfer ownership of the file
                //permissionRequest.Execute();

                //// Update the file object to include the new owner
                //Google.Apis.Drive.v3.Data.File file = new Google.Apis.Drive.v3.Data.File();
                //file.Owners = new List<User>() { new User() { EmailAddress = req.ReciverEmail } };

                //// Update the file with the new owner
                //FilesResource.UpdateRequest updateRequest = _DriveService.Files.Update(file, copyPresentationId);
                //updateRequest.Execute();

                //// Create the permission request and execute it to share the file
                //var request = _DriveService.Permissions.Create(permission, copyPresentationId);
                //request.SendNotificationEmail = true;
                //request.Execute();
                #endregion

                var slidesPresentation = _SlidesService.Presentations.Get(copyPresentationId).Execute();

                BatchUpdatePresentationRequest batchRequest = new BatchUpdatePresentationRequest();
                var requests = new List<Request>();

                if (req.TextPlaceholders?.Count > 0)
                    requests.AddRange(GetTextPlaceholdersRequests(req.TextPlaceholders));

                if (req.ImagePlaceholders?.Count > 0)
                    requests.AddRange(GetImagePlaceholdersRequests(req.ImagePlaceholders));

                if (req.ChartPlaceholders?.Count > 0)
                    requests.AddRange(GetChartPlaceholdersRequest(req.ChartPlaceholders));

                requests.AddRange(GetRemoveSlidesRequest(req.SlidesToRemove, slidesPresentation));

                if (req.MarketingSlidesIdsToAdd?.Count > 0)   //El problema esta en que es un dictionary y la claves se sobre escriben
                {
                    await AddMarketingSlides(req.MarketingSlidesIdsToAdd , copyPresentationId, req.SlidesOrder);
                }

                if (req.SlidesOrder.Count > 0)
                {
                    await ReorderSlides(req.SlidesOrder, copyPresentationId);
                }

                batchRequest.Requests = requests;

                var response = await _SlidesService.Presentations.BatchUpdate(batchRequest, copyPresentationId).ExecuteAsync();



            }
            catch (Exception e)
            {
                await _DriveService.Files.Delete(copyPresentationId).ExecuteAsync();
                return BadRequest("Error: " + e.Message);
            }

            return Ok($@"https://docs.google.com/presentation/d/{copyPresentationId}/edit");
        }

        private async Task ReorderSlides(IDictionary<string, int> slidesOrder, string? copyPresentationId)
        {
            var requests = new List<Request>();

            foreach (var sOrder in slidesOrder)
            {
                var request = new UpdateSlidesPositionRequest
                {
                    SlideObjectIds = new List<string> { sOrder.Key },
                    InsertionIndex = sOrder.Value
                };

                requests.Add(new Request() { UpdateSlidesPosition = request});
            }

            var batchUpdateRequest = new BatchUpdatePresentationRequest
            {
                Requests = requests
            };

            var updateRequest = _SlidesService.Presentations.BatchUpdate(batchUpdateRequest, copyPresentationId);
            await updateRequest.ExecuteAsync();
        }

        private async Task AddMarketingSlides(IDictionary<string, string> marketingPresentationIdsToAdd, string copyPresentationId)
        {

            foreach (var marketingPresentationKeyValue in marketingPresentationIdsToAdd)
            {

                await AddImageSlide(copyPresentationId, 
                    $"https://googleslidesapi.azurewebsites.net/thumbnails/{marketingPresentationKeyValue.Value}/LARGE/{marketingPresentationKeyValue.Key}.png",
                    newSlideId: marketingPresentationKeyValue.Key);
            }
            
        }

        private async Task AddMarketingSlides(IDictionary<string, string> marketingPresentationIdsToAdd, string targerPresentationId, IDictionary<string, int> SlidesOrder) {

            using (var client = new HttpClient())
            {
                foreach (var marketingPresentationKeyValue in marketingPresentationIdsToAdd)
                {
                    var url = $"https://script.google.com/macros/s/AKfycbzCqVDHp9rXYKHemTzdut0_FuhQP6T4q8SQeI-_b3WO8zXzdXj6mNQemwXpqPpXt_eW/exec?srcId={marketingPresentationKeyValue.Value}&dstId={targerPresentationId}&srcPage={marketingPresentationKeyValue.Key}";
                    
                    HttpResponseMessage respuesta = await client.GetAsync(url);
                    
                    if (respuesta.IsSuccessStatusCode)
                    {

                        string newSlideId = await respuesta.Content.ReadAsStringAsync();

                        if (SlidesOrder.ContainsKey(marketingPresentationKeyValue.Key))
                        {
                            int slideIndex = SlidesOrder[marketingPresentationKeyValue.Key];
                            SlidesOrder.Remove(marketingPresentationKeyValue.Key);

                            SlidesOrder.Add(newSlideId, slideIndex);
                        }

                    }
                }
            }

        }

        private async Task AddMarketingSlide(int marketinSlideIndex, string targetPresentationId)
        {

            string sourcePresentationId = "1aWCcKOeWR6EvTB2ldZThn4fZVObbuoHoWfteISWPqxc";
            int sourceSlideIndex = marketinSlideIndex; // Index of the source slide (zero-based)

            var sourcePresentation = _SlidesService.Presentations.Get(sourcePresentationId).Execute();
            var sourceSlide = sourcePresentation.Slides[sourceSlideIndex];

            var elementsToCopy = new List<PageElement>();

            foreach (var element in sourceSlide.PageElements)
            {
                elementsToCopy.Add(element);
            }

            var copySlideRequest = new CreateSlideRequest
            {
                ObjectId = Guid.NewGuid().ToString()
            };

            var batchUpdateRequest = new BatchUpdatePresentationRequest
            {
                Requests = new List<Request>
                {
                    new Request
                    {
                        CreateSlide = copySlideRequest
                    }
                }
            };

            var batchUpdateResponse = _SlidesService.Presentations.BatchUpdate(batchUpdateRequest, targetPresentationId).Execute();

            string copiedSlideId = batchUpdateResponse.Replies[0].CreateSlide.ObjectId;

            foreach (var element in elementsToCopy)
            {

                if (element.Image != null)
                {
                    var createRequest = new CreateImageRequest
                    {
                        ObjectId = Guid.NewGuid().ToString(),
                        ElementProperties = new PageElementProperties { Transform = element.Transform, Size = element.Size, PageObjectId = copiedSlideId },
                        Url = element.Image.SourceUrl

                    };

                    var batchUpdateRequest2 = new BatchUpdatePresentationRequest
                    {
                        Requests = new List<Request>
                        {
                            new Request
                            {
                                CreateImage = createRequest
                            }
                        }
                    };

                    _SlidesService.Presentations.BatchUpdate(batchUpdateRequest2, targetPresentationId).Execute();
                }
                else
                {

                    var createRequest = new CreateShapeRequest
                    {
                        ObjectId = Guid.NewGuid().ToString(),
                        ShapeType = element.Shape.ShapeType,
                        ElementProperties = new PageElementProperties { Size = element.Size, Transform = element.Transform, PageObjectId = copiedSlideId },

                    };

                    var batchUpdateRequest2 = new BatchUpdatePresentationRequest
                    {
                        Requests = new List<Request>
                    {
                        new Request
                        {
                            CreateShape = createRequest
                        }
                    }
                    };

                    _SlidesService.Presentations.BatchUpdate(batchUpdateRequest2, targetPresentationId).Execute();

                }


            }

            var targetPresentation = _SlidesService.Presentations.Get(targetPresentationId).Execute();

            IList<PageElement>? pageElements = targetPresentation.Slides.FirstOrDefault(s => s.ObjectId == copiedSlideId)?.PageElements;
            for (int i = 0; i < pageElements.Count; i++)
            {
                PageElement? pageElement = pageElements[i];

                if (pageElement.Shape != null)
                {
                    var content = string.Join("", elementsToCopy[i].Shape.Text.TextElements
                                                    .Where(t => t.TextRun != null)
                                                    .Select(t => t.TextRun.Content)
                                             )
                                        .Replace("\n", "");


                    var batchUpdateRequest3 = new BatchUpdatePresentationRequest
                    {
                        Requests = new List<Request>
                        {
                            new Request
                            {
                                InsertText = new InsertTextRequest() { ObjectId = pageElement.ObjectId, Text = content },


                            }
                        }
                    };

                    _SlidesService.Presentations.BatchUpdate(batchUpdateRequest3, targetPresentationId).Execute();

                    var batchUpdateRequest4 = new BatchUpdatePresentationRequest
                    {
                        Requests = new List<Request>
                        {
                            new Request
                            {

                                 UpdateShapeProperties = new UpdateShapePropertiesRequest
                                {
                                    ObjectId  = pageElement.ObjectId,
                                    Fields = "shapeBackgroundFill,outline,shadow,link,contentAlignment",
                                    ShapeProperties = pageElement.Shape.ShapeProperties
                                }
                            }
                        }
                    };

                    _SlidesService.Presentations.BatchUpdate(batchUpdateRequest4, targetPresentationId).Execute();
                }


            }


        }


        [HttpGet("TemplateList")]
        [ProducesResponseType(200, Type = typeof(TemplateInfoResponse[]))]
        public async Task<IActionResult> GetTemplates()
        {

            // Define parameters of request.
            FilesResource.ListRequest listRequest = _DriveService.Files.List();
            listRequest.Q = "mimeType='application/vnd.google-apps.presentation' and name starts with 'STemplate' and trashed = false";
            listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;

            var filesInfo = new List<TemplateInfoResponse>();
            foreach (var file in files)
            {
                filesInfo.Add(new TemplateInfoResponse()
                {
                    Id = file.Id,
                    Name = file.Name
                });
            }

            return Ok(filesInfo);

        }

        [HttpGet("PresentationsList")]
        [ProducesResponseType(200, Type = typeof(TemplateInfoResponse[]))]
        public async Task<IActionResult> GetPresentations()
        {

            // Define parameters of request.
            FilesResource.ListRequest listRequest = _DriveService.Files.List();
            listRequest.Q = "mimeType='application/vnd.google-apps.presentation' and not name starts with 'STemplate' and trashed = false";
            listRequest.Fields = "nextPageToken, files(id, name, createdTime)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;

            var filesInfo = new List<object>();
            foreach (var file in files)
            {
                filesInfo.Add(new
                {
                    Id = file.Id,
                    Name = file.Name,
                    CreatedOn = file.CreatedTime
                });
            }

            return Ok(filesInfo);

        }

        [HttpGet("FilesList")]
        [ProducesResponseType(200, Type = typeof(TemplateInfoResponse[]))]
        public async Task<IActionResult> GetAllItems()
        {

            // Define parameters of request.
            FilesResource.ListRequest listRequest = _DriveService.Files.List();
            listRequest.Q = "trashed = false";
            listRequest.Fields = "nextPageToken, files(id, name, createdTime, mimeType)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;

            var filesInfo = new List<object>();
            foreach (var file in files)
            {
                filesInfo.Add(new
                {
                    Id = file.Id,
                    Name = file.Name,
                    CreatedOn = file.CreatedTime,
                    Type = file.MimeType
                });
            }

            return Ok(filesInfo);

        }

        [HttpGet("About")]
        [ProducesResponseType(200, Type = typeof(object))]
        public async Task<IActionResult> GetAbout()
        {

            var aboutReq = _DriveService.About.Get();
            aboutReq.Fields = "*";

            var about = aboutReq.Execute();
            return Ok(about);

        }

        [HttpDelete("RemoveAllPresentations")]
        [ProducesResponseType(200, Type = typeof(TemplateInfoResponse[]))]
        public async Task<IActionResult> RemoveAllPresentations()
        {

            // Define parameters of request.
            FilesResource.ListRequest listRequest = _DriveService.Files.List();
            listRequest.Q = "mimeType='application/vnd.google-apps.presentation' and not name starts with 'STemplate' and trashed = false";
            listRequest.Fields = "nextPageToken, files(id, name, createdTime)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;

            foreach (var file in files)
            {
                // Create a delete request for the file
                FilesResource.DeleteRequest deleteRequest = _DriveService.Files.Delete(file.Id);

                // Execute the request and delete the file
                deleteRequest.Execute();
            }


            return Ok("All presentations removed");

        }

        [HttpDelete("Remove/{fileId}")]
        [ProducesResponseType(200, Type = typeof(TemplateInfoResponse[]))]
        public async Task<IActionResult> RemovePresentationById(string fileId)
        {
            _DriveService.Files.Delete(fileId).Execute();

            return Ok("File Deleted");
        }


        [HttpGet("PlaceholdesList/{templateId}")]
        [ProducesResponseType(200, Type = typeof(TemplateInfoResponse[]))]
        public async Task<IActionResult> GetPlaceholdersList(string templateId)
        {

            var presentation = _SlidesService.Presentations.Get(templateId).Execute();
            if (presentation == null)
            {
                return BadRequest("Template not found");
            }

            var placeholderPattern = @"{{.[^\{\}]{0,}}}";

            var matchingTextElements = new List<string>();

            foreach (var slide in presentation.Slides)
            {
                if (slide.PageElements == null)
                {
                    continue;
                }

                foreach (var element in slide.PageElements)
                {
                    if (element?.Shape?.Text?.TextElements != null)
                    {

                        var content = string.Join("", element.Shape.Text.TextElements
                                                        .Where(t => t.TextRun != null)
                                                        .Select(t => t.TextRun.Content)
                                                 )
                                            .Replace("\n", "");

                        var matchResult = Regex.Match(content, placeholderPattern);

                        if (matchResult.Success)
                        {
                            matchingTextElements.Add(content);
                        }

                    }
                }
            }

            return Ok(matchingTextElements);

        }

        [HttpGet("Metadata/{templateId}")]
        [ProducesResponseType(200, Type = typeof(TemplateMetadata))]
        public async Task<IActionResult> GetMetadata(string templateId)
        {

            var presentation = _SlidesService.Presentations.Get(templateId).Execute();
            if (presentation == null)
            {
                return BadRequest("Template not found");
            }

            var savedTemplateMetadata = _templateService.GetTemplateById(templateId);
            if (savedTemplateMetadata != null)
            {
                return Ok(savedTemplateMetadata);
            }

            var placeholderPattern = @"{{.[^\{\}]{0,}}}";

            var templateMetadata = new TemplateMetadata();
            templateMetadata.Id = presentation.PresentationId;
            templateMetadata.Name = presentation.Title;

            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                Page? slide = presentation.Slides[i];

                var slideMetadata = new SlideMetadata();
                slideMetadata.Id = slide.ObjectId;
                slideMetadata.Index = i;
                slideMetadata.Removable = true;

                if (slide.PageElements != null)
                {

                    foreach (var element in slide.PageElements)
                    {
                        var placeholderMetadata = new PlaceholderMetadata();

                        if (element?.Shape?.Text?.TextElements != null)
                        {

                            var content = string.Join("", element.Shape.Text.TextElements
                                                            .Where(t => t.TextRun != null)
                                                            .Select(t => t.TextRun.Content)
                                                     )
                                                .Replace("\n", "");

                            var matchResult = Regex.Match(content, placeholderPattern);

                            if (matchResult.Success)
                            {

                                placeholderMetadata.Name = RemoveMetadataText(content);

                                if (content.Contains("IMAGE"))
                                    placeholderMetadata.Type = "IMAGE";

                                else if (content.Contains("CHART"))
                                    placeholderMetadata.Type = "CHART";

                                else
                                    placeholderMetadata.Type = "TEXT";

                                placeholderMetadata.MaxLength = ExtractNumberFromMax(content);
                                placeholderMetadata.Editable = content.Contains("EDITABLE");
                                placeholderMetadata.Removable = content.Contains("REMOVABLE");

                                slideMetadata.Placeholders.Add(placeholderMetadata);
                            }

                        }
                    }

                }

                templateMetadata.Slides.Add(slideMetadata);

            }

            //_templateService.SaveTemplate(templateMetadata);

            return Ok(templateMetadata);

        }

        [HttpGet("Thumbnails/Generate/{presentationId}/{size?}")]
        public async Task<ActionResult<IEnumerable<string>>> GenerateSlideThumbnails(string presentationId, string size = "MEDIUM")
        {
            if (!new string[] {"SMALL", "MEDIUM", "LARGE"}.Contains(size))
            {
                return BadRequest("Invalid size. Use SMALL, MEDIUM or LARGE");
            }

            // Get the presentation
            var presentation = _SlidesService.Presentations.Get(presentationId).Execute();

            var thumbnailUrls = new List<string>();

            // Create a directory path for the thumbnails

            var thumbnailsFolderPath = Path.Combine(_hostingEnvironment.WebRootPath, "thumbnails", presentationId);

            var thumbnailSize = PresentationsResource.PagesResource.GetThumbnailRequest.ThumbnailPropertiesThumbnailSizeEnum.MEDIUM;

            if (!string.IsNullOrWhiteSpace(size))
                thumbnailsFolderPath = Path.Combine(thumbnailsFolderPath, size);


            switch (size)
            {
                case "SMALL":
                    thumbnailSize = PresentationsResource.PagesResource.GetThumbnailRequest.ThumbnailPropertiesThumbnailSizeEnum.SMALL;
                    break;

                case "MEDIUM":
                    thumbnailSize = PresentationsResource.PagesResource.GetThumbnailRequest.ThumbnailPropertiesThumbnailSizeEnum.MEDIUM;
                    break;

                case "LARGE":
                    thumbnailSize = PresentationsResource.PagesResource.GetThumbnailRequest.ThumbnailPropertiesThumbnailSizeEnum.LARGE;
                    break;
            }


            // Create the directory if it doesn't exist
            Directory.CreateDirectory(thumbnailsFolderPath);

            // Create HttpClient instance
            using (var httpClient = new HttpClient())
            {
                // Generate thumbnails for each slide
                for (int i = 0; i < presentation.Slides.Count; i++)
                {
                    Page? slide = presentation.Slides[i];
                    // Create a thumbnail request
                    var thumbnailRequest = _SlidesService.Presentations.Pages.GetThumbnail(presentationId, slide.ObjectId);
                    thumbnailRequest.ThumbnailPropertiesMimeType = PresentationsResource.PagesResource.GetThumbnailRequest.ThumbnailPropertiesMimeTypeEnum.PNG;
                    thumbnailRequest.ThumbnailPropertiesThumbnailSize = thumbnailSize;

                    //TODO: Try to create a batch of request for getting the thumbnails

                    // Execute the request and get the thumbnail image URL
                    var thumbnailResponse = await thumbnailRequest.ExecuteAsync();
                    var thumbnailUrl = thumbnailResponse.ContentUrl;

                    // Generate a unique filename for the thumbnail
                    var thumbnailFilename = $"{slide.ObjectId}.png";
                    var thumbnailFilePath = Path.Combine(thumbnailsFolderPath, thumbnailFilename);

                    // Download the thumbnail image using HttpClient
                    var response = await httpClient.GetAsync(thumbnailUrl);
                    if (response.IsSuccessStatusCode)
                    {
                        using (var stream = await response.Content.ReadAsStreamAsync())
                        {
                            using (var fileStream = new FileStream(thumbnailFilePath, FileMode.Create))
                            {
                                await stream.CopyToAsync(fileStream);
                            }
                        }
                    }

                    // Add the thumbnail file path to the list
                    var thumbnailUrlPath = $"/thumbnails/{presentationId}/{size}/{thumbnailFilename}";
                    thumbnailUrls.Add(thumbnailUrlPath);
                }
            }

            return Ok(thumbnailUrls);
        }

        [HttpGet("Thumbnails/Get/{presentationId}")]
        public async Task<ActionResult<IEnumerable<string>>> GetSlideThumbnails(string presentationId)
        {
            var thumbnailsFolderPath = Path.Combine(_hostingEnvironment.WebRootPath, "thumbnails", presentationId);

            if (!Directory.Exists(thumbnailsFolderPath))
            {
                return NotFound();
            }

            var thumbnailUrls = new List<string>();
            var thumbnailFiles = Directory.GetFiles(thumbnailsFolderPath);

            foreach (var thumbnailFile in thumbnailFiles)
            {
                var thumbnailUrl = Path.GetFileName(thumbnailFile);
                thumbnailUrls.Add(thumbnailUrl);
            }

            return Ok(thumbnailUrls);
        }

        [HttpGet("folder/{folderId}")]
        public async Task<IActionResult> GetMarketingFolders(string folderId)
        {
            return Ok(await GetFolderWithSubfolders(folderId));
        }

        private async Task<FolderMetadata> GetFolderWithSubfolders(string folderId)
        {
            FolderMetadata folderMetadata = new FolderMetadata();

            // Create the request to retrieve the files and folders within the specified folder
            FilesResource.ListRequest request = _DriveService.Files.List();
            request.Q = $"'{folderId}' in parents and mimeType = 'application/vnd.google-apps.folder'";
            request.Fields = "files(id, name)";
            request.PageSize = 1000; // Adjust the page size as needed

            do
            {
                // Retrieve the files and folders
                FileList fileList = await request.ExecuteAsync();

                foreach (Google.Apis.Drive.v3.Data.File folder in fileList.Files)
                {
                    FolderMetadata subfolder = new FolderMetadata
                    {
                        Id = folder.Id,
                        Name = folder.Name
                    };

                    await GetAllSlidesInFolder(subfolder);
                    // Recursively get subfolders
                    subfolder.Subfolders = await GetSubfolders(folder.Id);

                    folderMetadata.Subfolders.Add(subfolder);
                }

                request.PageToken = fileList.NextPageToken;
            }
            while (!string.IsNullOrEmpty(request.PageToken));

            

            return folderMetadata;
        }

        private async Task<List<FolderMetadata>> GetSubfolders(string parentFolderId)
        {
            List<FolderMetadata> subfolders = new List<FolderMetadata>();

            FilesResource.ListRequest request = _DriveService.Files.List();
            request.Q = $"'{parentFolderId}' in parents and mimeType = 'application/vnd.google-apps.folder'";
            request.Fields = "files(id, name)";
            request.PageSize = 1000; // Adjust the page size as needed

            do
            {
                FileList fileList = request.Execute();

                foreach (Google.Apis.Drive.v3.Data.File folder in fileList.Files)
                {
                    FolderMetadata subfolder = new FolderMetadata
                    {
                        Id = folder.Id,
                        Name = folder.Name
                    };

                    await GetGooglePresentationsInFolder(subfolder);
                    // Recursively get subfolders
                    subfolder.Subfolders = await GetSubfolders(folder.Id);

                    subfolders.Add(subfolder);
                }

                request.PageToken = fileList.NextPageToken;
            }
            while (!string.IsNullOrEmpty(request.PageToken));

            return subfolders;
        }


        private async Task GetGooglePresentationsInFolder(FolderMetadata folder)
        {
            var request = _DriveService.Files.List();
            request.Q = $"'{folder.Id}' in parents and mimeType = 'application/vnd.google-apps.presentation'";
            request.Fields = "files(*)";
            var response = request.Execute();
            var presentations = response.Files;

            foreach (var presentation in presentations)
            {
                var slidePresentation = await _SlidesService.Presentations.Get(presentation.Id).ExecuteAsync();

                for (int i = 0; i < slidePresentation.Slides.Count; i++)
                {
                    Page? slide = slidePresentation.Slides[i];
                    var slideItem = new FolderMetadata.SlideItem
                    {
                        SlideId = slide.ObjectId,
                        PresentationId = presentation.Id,
                        PresentationName = presentation.Name,
                        Name = GetSlideTitle(slidePresentation, slide.ObjectId) ?? "No title",
                        ThumbnailUrl = $"/thumbnails/{presentation.Id}/MEDIUM/{slide.ObjectId}.png"
                    };
                    folder.SlideItems.Add(slideItem);
                }
            }
        }

        private async Task GetAllSlidesInFolder(FolderMetadata folder)
        {
            var request = _DriveService.Files.List();
            request.Q = $"'{folder.Id}' in parents and mimeType = 'application/vnd.google-apps.presentation'";
            request.Fields = "files(*)";
            var response = request.Execute();
            var presentations = response.Files;

            foreach (var presentation in presentations)
            {

                var slidePresentation = await _SlidesService.Presentations.Get(presentation.Id).ExecuteAsync();

                for (int i = 0; i < slidePresentation.Slides.Count; i++)
                {
                    Page? slide = slidePresentation.Slides[i];
                    var slideItem = new FolderMetadata.SlideItem
                    {
                        SlideId = slide.ObjectId,
                        PresentationId = presentation.Id,
                        PresentationName = presentation.Name,
                        Name = GetSlideTitle(slidePresentation, slide.ObjectId) ?? "No title",
                        ThumbnailUrl = $"/thumbnails/{presentation.Id}/MEDIUM/{slide.ObjectId}.png"
                    };
                    folder.SlideItems.Add(slideItem);
                }
                
            }
        }


        private string? GetSlideTitle(Presentation presentation, string slideId)
        {
            // Find the slide with the specified slideId
            var slide = presentation.Slides.FirstOrDefault(s => s.ObjectId == slideId);

            if (slide != null)
            {
                // Find the page element of type 'TITLE'
                var titleElement = slide.PageElements.FirstOrDefault(element => element.Shape?.Text != null && element.Shape.Text.TextElements?.FirstOrDefault(x=>x.TextRun?.Content != null)?.TextRun?.Content != null);

                if (titleElement != null)
                {
                    // Return the title of the slide
                    return titleElement.Shape.Text.TextElements.FirstOrDefault(x => x.TextRun?.Content != null)?.TextRun?.Content.Replace("\n", ""); ;
                }
            }

            // Slide or title not found
            return null;
        }




        private string RemoveMetadataText(string input)
        {
            int colonIndex = input.IndexOf(':');
            if (colonIndex != -1)
            {
                string result = input.Substring(0, colonIndex);
                return result + "}}"; // Add curly brackets at the end
            }

            return input; // If colon not found, return the original input
        }

        private int ExtractNumberFromMax(string input)
        {
            string pattern = @"\BMAX([0-9]{0,})";
            Match match = Regex.Match(input, pattern);

            if (match.Success)
            {
                string numberString = match.Groups[1].Value;
                int number;
                if (int.TryParse(numberString, out number))
                {
                    return number;
                }
            }

            return -1; // Default value if no number found or parsing fails
        }


        private IEnumerable<Request> GetChartPlaceholdersRequest(IDictionary<string, ChartInfo> chartPlaceholders)
        {
            var requests = new List<Request>();

            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Google.Apis.Sheets.v4.Data.Request>();

            foreach (var chartPlaceholder in chartPlaceholders)
            {

                var batchUpdateValuesRequest = new BatchUpdateValuesRequest();
                var valueRanges = new List<ValueRange>();

                var chartInfo = chartPlaceholder.Value;

                var chartSpec = new ChartSpec();
                chartSpec.Title = chartInfo.Title;
                chartSpec.Subtitle = chartInfo.Subtitle;

                if (chartInfo.Type == "PIE" || chartInfo.Type == "DOUGHNUT")
                {
                    chartSpec.PieChart = new PieChartSpec();
                    chartSpec.PieChart.ThreeDimensional = false;
                    chartSpec.PieChart.PieHole = chartInfo.Type == "PIE" ? 0 : 0.5;

                    chartSpec.PieChart.Domain = new ChartData()
                    {
                        SourceRange = new ChartSourceRange()
                        {
                            Sources = new List<GridRange>()
                            {
                                new GridRange()
                                {
                                    EndColumnIndex = 1 ,
                                    EndRowIndex = int.MaxValue,
                                    StartColumnIndex = 0,
                                    StartRowIndex = 0
                                }
                            }
                        },

                    };

                    chartSpec.PieChart.Series = new ChartData()
                    {
                        SourceRange = new ChartSourceRange()
                        {
                            Sources = new List<GridRange>()
                            {
                                new GridRange()
                                {
                                    EndColumnIndex = 2,
                                    EndRowIndex = int.MaxValue,

                                    StartColumnIndex = 1,
                                    StartRowIndex = 0
                                }
                            }
                        }
                    };

                    //Column A
                    var columnData = new List<IList<object>>();

                    var domain = chartInfo.Domains.ElementAt(0);

                    columnData.Add(new List<object>() { domain.Key });

                    foreach (var dataCell in domain.Value)
                    {
                        columnData.Add(new List<object>() { dataCell });
                    }

                    var dataRange = new ValueRange()
                    {
                        Range = $"A:A",
                        Values = columnData
                    };

                    valueRanges.Add(dataRange);

                    //Column B
                    var seriesColumnData = new List<IList<object>>();

                    var serie = chartInfo.Series.ElementAt(0);

                    seriesColumnData.Add(new List<object>() { serie.Key });

                    foreach (var dataCell in serie.Value)
                    {
                        seriesColumnData.Add(new List<object>() { dataCell });
                    }

                    var seriesDataRange = new ValueRange()
                    {
                        Range = $"B:B",
                        Values = seriesColumnData
                    };

                    valueRanges.Add(seriesDataRange);

                }
                else
                {



                    chartSpec.BasicChart = new BasicChartSpec();
                    chartSpec.BasicChart.ChartType = chartInfo.Type;
                    chartSpec.BasicChart.StackedType = chartInfo.StackedType;
                    chartSpec.BasicChart.HeaderCount = 1;
                    chartSpec.BasicChart.LegendPosition = chartInfo.LegendPosition;

                    //Adding axis
                    chartSpec.BasicChart.Axis = new List<BasicChartAxis>()
                {
                    new BasicChartAxis()
                    {
                        Position = "BOTTOM_AXIS",
                        Title = chartInfo.BottomAxisName
                    },
                    new BasicChartAxis()
                    {
                        Position = "LEFT_AXIS",
                        Title = chartInfo.LeftAxisName
                    }
                };


                    //Adding domains
                    chartSpec.BasicChart.Domains = new List<BasicChartDomain>();
                    for (int i = 0; i < chartInfo.Domains.Count; i++)
                    {
                        var domain = chartInfo.Domains.ElementAt(i);

                        chartSpec.BasicChart.Domains.Add(new BasicChartDomain()
                        {
                            Domain = new ChartData()
                            {
                                SourceRange = new ChartSourceRange()
                                {
                                    Sources = new List<GridRange>()
                                {
                                    new GridRange()
                                    {
                                        EndColumnIndex = i + 1 ,
                                        EndRowIndex = int.MaxValue,
                                        StartColumnIndex = 0,
                                        StartRowIndex = 0
                                    }
                                }
                                }
                            }
                        });


                        var columnData = new List<IList<object>>();

                        //add header of column
                        columnData.Add(new List<object>() { domain.Key });

                        foreach (var dataCell in domain.Value)
                        {
                            columnData.Add(new List<object>() { dataCell });
                        }

                        var columnLetter = GetColumnLetter(i);

                        // Add the new data to the spreadsheet
                        var dataRange = new ValueRange()
                        {
                            Range = $"{columnLetter}:{columnLetter}",
                            Values = columnData
                        };

                        valueRanges.Add(dataRange);

                    }

                    //Adding series
                    chartSpec.BasicChart.Series = new List<BasicChartSeries>();
                    for (int i = 0; i < chartInfo.Series.Count; i++)
                    {
                        var serie = chartInfo.Series.ElementAt(i);

                        chartSpec.BasicChart.Series.Add(new BasicChartSeries()
                        {
                            Series = new ChartData()
                            {
                                SourceRange = new ChartSourceRange()
                                {
                                    Sources = new List<GridRange>()
                                {
                                    new GridRange()
                                    {
                                        EndColumnIndex = i + (chartInfo.Domains.Count) + 1,
                                        EndRowIndex = int.MaxValue,

                                        StartColumnIndex = (i + (chartInfo.Domains.Count) + 1) - 1,
                                        StartRowIndex = 0
                                    }
                                }
                                }
                            }
                        });

                        var columnData = new List<IList<object>>();

                        //add header of column
                        columnData.Add(new List<object>() { serie.Key });

                        foreach (var dataCell in serie.Value)
                        {
                            columnData.Add(new List<object>() { dataCell });
                        }

                        var columnLetter = GetColumnLetter(i + chartInfo.Domains.Count);

                        // Add the new data to the spreadsheet
                        var dataRange = new ValueRange()
                        {
                            Range = $"{columnLetter}:{columnLetter}",
                            Values = columnData
                        };

                        valueRanges.Add(dataRange);

                    }
                }


                //Update the values
                _SheetsService.Spreadsheets.Values.BatchUpdate(new BatchUpdateValuesRequest()
                {
                    Data = valueRanges,
                    ValueInputOption = "USER_ENTERED"
                }, "1J7sP682rkpLtGiRXcxJVkeCGKhCd8iUNktRomp2iEM0").Execute();


                var addChartRequest = new AddChartRequest()
                {
                    Chart = new EmbeddedChart()
                    {
                        Spec = chartSpec,
                        Position = new EmbeddedObjectPosition()
                        {
                            OverlayPosition = new OverlayPosition()
                            {
                                OffsetXPixels = 50,
                                OffsetYPixels = 50,
                                WidthPixels = 500,
                                HeightPixels = 350,
                                //AnchorCell = new GridCoordinate() { SheetId = 843514080 }
                            }
                        }

                    }
                };

                //Add it to the batchUpdateSpreadheet
                batchUpdateSpreadsheetRequest.Requests.Add(new Google.Apis.Sheets.v4.Data.Request()
                {
                    AddChart = addChartRequest
                });




            }

            var response = _SheetsService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, "1J7sP682rkpLtGiRXcxJVkeCGKhCd8iUNktRomp2iEM0").Execute();
            for (int i = 0; i < response.Replies.Count; i++)
            {
                Google.Apis.Sheets.v4.Data.Response? res = response.Replies[i];
                if (res.AddChart != null)
                {
                    requests.Add(new Request()
                    {
                        ReplaceAllShapesWithSheetsChart = new ReplaceAllShapesWithSheetsChartRequest()
                        {
                            SpreadsheetId = "1J7sP682rkpLtGiRXcxJVkeCGKhCd8iUNktRomp2iEM0",
                            ChartId = res.AddChart.Chart.ChartId,
                            LinkingMode = "NOT_LINKED_IMAGE",     // LINKED or NOT_LINKED_IMAGE
                            ContainsText = new SubstringMatchCriteria() { MatchCase = true, Text = chartPlaceholders.ElementAt(i).Key }
                        }
                    });
                }
            }

            return requests;
        }

        private string GetColumnLetter(int columnIndex)
        {
            string columnLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string columnLetter = "";
            while (columnIndex >= 0)
            {
                int letterIndex = (columnIndex % 26);
                columnLetter = columnLetters[letterIndex] + columnLetter;
                columnIndex = (columnIndex / 26) - 1;
            }
            return columnLetter;
        }

        private List<Request> GetRemoveSlidesRequest(string[] slidesToRemoveIds, Presentation slidesPresentation)
        {
            var request = new List<Request>();

            foreach (var slideToRemoveId in slidesToRemoveIds)
            {
                request.Add(new Request()
                {
                    DeleteObject = new DeleteObjectRequest()
                    {
                        ObjectId = slideToRemoveId
                    }
                });
            }

            return request;
        }

        private async Task AddImageSlide(string presentationId, string imageUrl, string newSlideId)
        {
            //string presentationId = "1aWCcKOeWR6EvTB2ldZThn4fZVObbuoHoWfteISWPqxc";
            //string imageUrl = "https://googleslidesapi.azurewebsites.net/thumbnails/1yJzb5prVVdWZ4WAfKK6y0tOZ_pn0U8rR8KZ3UGvonfg/LARGE/slide0.png";


            var getRequest = _SlidesService.Presentations.Get(presentationId);
            var presentation = getRequest.Execute();

            var createSlideRequest = new CreateSlideRequest
            {
                ObjectId = newSlideId
                //InsertionIndex = presentation.Slides.Count,
                //SlideLayoutReference = new LayoutReference { PredefinedLayout = "BLANK" }
            };
            var createSlideResponse = await _SlidesService.Presentations.BatchUpdate(new BatchUpdatePresentationRequest
            {
                Requests = new List<Request> { new Request { CreateSlide = createSlideRequest } }
            }, presentationId).ExecuteAsync();

            var pageId = createSlideResponse.Replies[0].CreateSlide.ObjectId;

            await _SlidesService.Presentations.BatchUpdate(new BatchUpdatePresentationRequest
            {
                Requests = new List<Request> {
                    new Request {
                        CreateImage = new CreateImageRequest {
                            Url = imageUrl,
                            ElementProperties = new PageElementProperties()
                            {
                                PageObjectId = pageId,
                                Transform = new AffineTransform { ScaleX = 1, ScaleY = 1, TranslateX = 0, TranslateY = 0, Unit = "PT" },
                                Size = new Size
                                {
                                    Height = new Dimension { Magnitude = 405.04, Unit = "PT" },
                                    Width = new Dimension { Magnitude = 719.42, Unit = "PT" }
                                },
                            }
                        }
                    }
                }
            }, presentationId).ExecuteAsync();

        }


        private List<Request> GetTextPlaceholdersRequests(IDictionary<string, string> textPlaceholders)
        {
            List<Request> requests = new List<Request>();

            foreach (KeyValuePair<string, string> placeholder in textPlaceholders)
            {
                requests.Add(new Request
                {
                    ReplaceAllText = new ReplaceAllTextRequest
                    {
                        ContainsText = new SubstringMatchCriteria() { MatchCase = true, Text = placeholder.Key },
                        ReplaceText = placeholder.Value
                    }
                });
            }

            return requests;
        }

        private List<Request> GetImagePlaceholdersRequests(IDictionary<string, string> imagePlaceholders)
        {
            List<Request> requests = new List<Request>();

            foreach (KeyValuePair<string, string> placeholder in imagePlaceholders)
            {
                requests.Add(new Request
                {
                    ReplaceAllShapesWithImage = new ReplaceAllShapesWithImageRequest()
                    {
                        ContainsText = new SubstringMatchCriteria() { MatchCase = true, Text = placeholder.Key },
                        ImageUrl = placeholder.Value,
                        ImageReplaceMethod = "CENTER_INSIDE",
                    }
                });
            }

            return requests;
        }
    }
}
