using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
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
using Microsoft.Extensions.FileSystemGlobbing.Internal;
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

        public SlidesControler(DriveService driveService, SlidesService slidesService, SheetsService sheetsService)
        {

            _DriveService = driveService;
            _SlidesService = slidesService;
            _SheetsService = sheetsService;
        }

        [HttpPost]
        [ProducesResponseType(200, Type = typeof(string))]
        public async Task<IActionResult> CreateFromTemplate(CreateSlidesDeckFromTemplate req)
        {

            var originalPresentation = _DriveService.Files.Get(req.TemplateId).Execute();

            var copyPresentation = new Google.Apis.Drive.v3.Data.File
            {
                Name = req.PresentationName,
                Parents = originalPresentation.Parents
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


                var slidesPresentation = _SlidesService.Presentations.Get(copyPresentationId).Execute();

                BatchUpdatePresentationRequest batchRequest = new BatchUpdatePresentationRequest();
                var requests = new List<Request>();

                requests.AddRange(GetTextPlaceholdersRequests(req.TextPlaceholders));

                requests.AddRange(GetImagePlaceholdersRequests(req.ImagePlaceholders));

                requests.AddRange(GetChartPlaceholdersRequest(req.ChartPlaceholders));

                requests.AddRange(GetRemoveSlidesRequest(req.SlidesToRemove, slidesPresentation));

                //requests.Add(new Request()
                //{
                //    ReplaceAllShapesWithSheetsChart = new ReplaceAllShapesWithSheetsChartRequest()
                //    {
                //        SpreadsheetId = "1J7sP682rkpLtGiRXcxJVkeCGKhCd8iUNktRomp2iEM0",
                //        ChartId = 73277889,
                //        LinkingMode = "LINKED",     // LINKED or NOT_LINKED_IMAGE
                //        ContainsText = new SubstringMatchCriteria() { MatchCase = true, Text = "{{SALES_CHART}}" }
                //    }
                //});

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


        [HttpGet("TemplateList")]
        [ProducesResponseType(200, Type = typeof(TemplateInfoResponse[]))]
        public async Task<IActionResult> GetTemplateList()
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

                chartSpec.BasicChart = new BasicChartSpec();
                chartSpec.BasicChart.ChartType = chartInfo.Type;
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
                                HeightPixels = 300,
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

        private List<Request> GetRemoveSlidesRequest(int[] slidesToRemove, Presentation slidesPresentation)
        {
            var request = new List<Request>();

            foreach (var slideToRemoveindex in slidesToRemove)
            {
                request.Add(new Request()
                {
                    DeleteObject = new DeleteObjectRequest()
                    {
                        ObjectId = slidesPresentation.Slides[slideToRemoveindex].ObjectId
                    }
                });
            }

            return request;
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
