using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Slides.v1;
using Google.Apis.Slides.v1.Data;
using Google.Apis.Util.Store;
using Org.BouncyCastle.Asn1.Crmf;
using System;
using System.Net;
using System.Net.Http.Headers;
using System.Security;
using System.Text;
using System.Xml.Linq;
using Request = Google.Apis.Slides.v1.Data.Request;

namespace GoogleSlide.Test
{
    [TestClass]
    public class GoogleSlidesTest
    {
        string[] Scopes = { SlidesService.Scope.Presentations, DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets, GmailService.Scope.GmailSend };
        UserCredential? Credential;
        DriveService DriveService = null!;
        SlidesService SlidesService = null!;
        SheetsService SheetsService = null!;
        GmailService GmailService = null!;

        [TestInitialize]
        public void InitTest()
        {
            Credential = GetCredentials();

            DriveService = new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = Credential,
                ApplicationName = "Google Drive Copy",
            });

            SlidesService = new SlidesService(new BaseClientService.Initializer
            {
                HttpClientInitializer = Credential,
                ApplicationName = "Google Slides Example"
            });

            SheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = Credential,
                ApplicationName = "Google Sheets Example"
            });

            GmailService = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = Credential,
                ApplicationName = "Gmail API Send Email",
            });
        }

        [TestMethod]
        public void CreateGoogleSlideAndAddText()
        {
            //var salesforcesData = GetData()

            // Create a new slide deck
            var presentation = new Presentation
            {
                Title = "My New Slide Deck from C#"
            };


            var request = SlidesService.Presentations.Create(presentation);
            var createdPresentation = request.Execute();

            Console.WriteLine("Created slide deck with ID: " + createdPresentation.PresentationId);

            var requests = new List<Request>();

            //Insert the text to the textbox
            requests.Add(new Request()
            {
                InsertText = new InsertTextRequest()
                {
                    ObjectId = "i0",
                    InsertionIndex = 0,
                    Text = "Hello, I was generated using Google slides API"
                }
            });

            try
            {
                BatchUpdatePresentationRequest body = new BatchUpdatePresentationRequest()
                {
                    Requests = requests
                };

                var response = SlidesService.Presentations
                   .BatchUpdate(body, createdPresentation.PresentationId)
                   .Execute();
            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }
        }

        [TestMethod]
        public void ClonePresentation()
        {

            string originalPresentationId = "19aWti0YTDmRRHbYh_j0lHu3T00xyIBwykjqGoPLqEZM";

            //Get the file
            var originalPresentation = DriveService.Files.Get(originalPresentationId).Execute();

            // Create a copy of the presentation
            var copyPresentation = new Google.Apis.Drive.v3.Data.File
            {
                Name = "Copy of " + originalPresentation.Name,
                Parents = originalPresentation.Parents
            };

            var copyRequest = DriveService.Files.Copy(copyPresentation, originalPresentationId);
            var copyResponse = copyRequest.Execute();

            // Get the ID of the newly created presentation
            string copyPresentationId = copyResponse.Id;

            Console.WriteLine("The ID of the newly created presentation is: " + copyPresentationId);
        }

        [TestMethod]
        public void ReplaceText()
        {

            string presentationId = "1ySkUyS9XNUJ-XugklHESXZSSM3N6xaeuWStk3IeBqdI";
            string oldText = "{{TITLE}}";
            string newText = "Titulo reemplazado";

            // Replace all instances of old text with new text in the presentation.
            BatchUpdatePresentationRequest request = new BatchUpdatePresentationRequest();
            request.Requests = new List<Request>();
            request.Requests.Add(new Request
            {
                ReplaceAllText = new ReplaceAllTextRequest
                {
                    ContainsText = new SubstringMatchCriteria() { MatchCase = true, Text = oldText },
                    ReplaceText = newText
                }
            });

            SlidesService.Presentations.BatchUpdate(request, presentationId).Execute();
        }


        [TestMethod]
        public void SwapShapeWithImage()
        {

            string presentationId = "16kkjG_XN190JJ8_Ob2yT--q6IM5DpsG_K5MSek4KUPw";
            string imageUrl = @"https://docpromo.in/wp-content/uploads/2022/02/Healthcare-Market-Segment-1-750x400.png";

            // Create a request to insert the image shape
            var batchRequest = new BatchUpdatePresentationRequest();

            batchRequest.Requests = new List<Request>();
            batchRequest.Requests.Add(new Request()
            {
                ReplaceAllShapesWithImage = new ReplaceAllShapesWithImageRequest()
                {
                    ContainsText = new SubstringMatchCriteria() { MatchCase = true, Text = "{{4_IMAGE}}" },
                    ImageUrl = imageUrl,
                    ImageReplaceMethod = "CENTER_INSIDE"
                }
            });

            // Execute the request
            var response = SlidesService.Presentations.BatchUpdate(batchRequest, presentationId).Execute();

        }

        [TestMethod]
        public void CloneTemplateAndReplaceText()
        {
            string templatePresentationId = "16kkjG_XN190JJ8_Ob2yT--q6IM5DpsG_K5MSek4KUPw";

            //Get the file
            var originalPresentation = DriveService.Files.Get(templatePresentationId).Execute();

            // Create a copy of the presentation
            var copyPresentation = new Google.Apis.Drive.v3.Data.File
            {
                Name = "Copy of " + originalPresentation.Name,
                Parents = originalPresentation.Parents
            };

            var copyRequest = DriveService.Files.Copy(copyPresentation, templatePresentationId);
            var copyResponse = copyRequest.Execute();

            // Get the ID of the newly created presentation
            string copyPresentationId = copyResponse.Id;

            Dictionary<string, string> textPlaceholders = new Dictionary<string, string>
            {
                { "{{TITLE}}", "Creating a Digital \nEcosystem for \nHealthcare and \nLife Sciences" },
                { "{{2_TITLE}}", "Healthcare market opportunities and challenges" },
                { "{{2_TEXT_LEFT}}", "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus id convallis justo, at fringilla ex. Mauris non lacus vestibulum magna interdum rhoncus. Nulla elementum justo non malesuada fringilla. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Vivamus turpis magna, interdum at tellus eu, consequat tempus felis. Suspendisse a justo odio. Donec aliquam, nunc a sodales pellentesque, ipsum risus vestibulum tellus, sed tristique enim nunc non nibh." },
                { "{{2_TEXT_RIGHT}}", "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus id convallis justo, at fringilla ex. Mauris non lacus vestibulum magna interdum rhoncus. Nulla elementum justo non malesuada fringilla. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Vivamus turpis magna, interdum at tellus eu, consequat tempus felis. Suspendisse a justo odio. Donec aliquam, nunc a sodales pellentesque, ipsum risus vestibulum tellus, sed tristique enim nunc non nibh." },
                { "{{3_ITEM_ONE}}", "Healthcare market" },
                { "{{3_ITEM_TWO}}", "Market opportunities" },
                { "{{3_ITEM_THREE}}", "Market Challenges" }

            };

            Dictionary<string, string> imagePlaceholders = new Dictionary<string, string>
            {
                { "{{4_IMAGE}}", "https://sp-ao.shortpixel.ai/client/to_webp,q_lossy,ret_img,w_750/https://www.digitalauthority.me/wp-content/uploads/2018/12/shutterstock_400002673.jpg" }
            };

            BatchUpdatePresentationRequest request = new BatchUpdatePresentationRequest();
            request.Requests = new List<Request>();

            // Replace all instances of text placeholders in the presentation.
            foreach (KeyValuePair<string, string> placeholder in textPlaceholders)
            {
                request.Requests.Add(new Request
                {
                    ReplaceAllText = new ReplaceAllTextRequest
                    {
                        ContainsText = new SubstringMatchCriteria() { MatchCase = true, Text = placeholder.Key },
                        ReplaceText = placeholder.Value
                    }
                });
            }

            // Replace all instances of image placeholders in the presentation.
            foreach (KeyValuePair<string, string> placeholder in imagePlaceholders)
            {
                request.Requests.Add(new Request
                {
                    ReplaceAllShapesWithImage = new ReplaceAllShapesWithImageRequest()
                    {
                        ContainsText = new SubstringMatchCriteria() { MatchCase = true, Text = placeholder.Key },
                        ImageUrl = placeholder.Value,
                        ImageReplaceMethod = "CENTER_INSIDE",
                    }
                });
            }

            request.Requests.Add(new Request()
            {
                ReplaceAllShapesWithSheetsChart = new ReplaceAllShapesWithSheetsChartRequest()
                {
                    SpreadsheetId = "1J7sP682rkpLtGiRXcxJVkeCGKhCd8iUNktRomp2iEM0",
                    ChartId = 73277889,
                    LinkingMode = "LINKED",     // LINKED or NOT_LINKED_IMAGE
                    ContainsText = new SubstringMatchCriteria() { MatchCase = true, Text = "{{SALES_CHART}}" }
                }
            });

            var response = SlidesService.Presentations.BatchUpdate(request, copyPresentationId).Execute();
        }

        [TestMethod]
        public void CloneTemplateAndReplaceTextAndImageFromDrive()
        {

            string templatePresentationId = "16kkjG_XN190JJ8_Ob2yT--q6IM5DpsG_K5MSek4KUPw";

            //Get the file
            var originalPresentation = DriveService.Files.Get(templatePresentationId).Execute();

            // Create a copy of the presentation
            var copyPresentation = new Google.Apis.Drive.v3.Data.File
            {
                Name = "Copy of " + originalPresentation.Name,
                Parents = originalPresentation.Parents
            };

            var copyRequest = DriveService.Files.Copy(copyPresentation, templatePresentationId);
            var copyResponse = copyRequest.Execute();

            // Get the ID of the newly created presentation
            string copyPresentationId = copyResponse.Id;

            Dictionary<string, string> textPlaceholders = new Dictionary<string, string>
            {
                { "{{TITLE}}", "Creating a Digital \nEcosystem for \nHealthcare and \nLife Sciences" },
                { "{{2_TITLE}}", "Healthcare market opportunities and challenges" },
                { "{{2_TEXT_LEFT}}", "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus id convallis justo, at fringilla ex. Mauris non lacus vestibulum magna interdum rhoncus. Nulla elementum justo non malesuada fringilla. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Vivamus turpis magna, interdum at tellus eu, consequat tempus felis. Suspendisse a justo odio. Donec aliquam, nunc a sodales pellentesque, ipsum risus vestibulum tellus, sed tristique enim nunc non nibh." },
                { "{{2_TEXT_RIGHT}}", "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus id convallis justo, at fringilla ex. Mauris non lacus vestibulum magna interdum rhoncus. Nulla elementum justo non malesuada fringilla. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Vivamus turpis magna, interdum at tellus eu, consequat tempus felis. Suspendisse a justo odio. Donec aliquam, nunc a sodales pellentesque, ipsum risus vestibulum tellus, sed tristique enim nunc non nibh." },
                { "{{3_ITEM_ONE}}", "Healthcare market" },
                { "{{3_ITEM_TWO}}", "Market opportunities" },
                { "{{3_ITEM_THREE}}", "Market  Challenges" }

            };



            //Uploading the image to drive

            var filePath = "image.png";

            var fileName = Path.GetFileName(filePath);

            // Create the file metadata
            var fileMetadata = new Google.Apis.Drive.v3.Data.File
            {
                Name = fileName + Guid.NewGuid(),
                Parents = new List<string>() { "1HtwFuFKFe4XJH6IIc8bHKbJZoJ4Tmr8s" }, //This is the Id of a shared folder
                CopyRequiresWriterPermission = false
            };

            string imageLink;
            string fileId;

            // Read the file content
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                // Upload the file
                var driveRequest = DriveService.Files.Create(fileMetadata, stream, "image/jpeg");
                driveRequest.Fields = "id";

                var uploadResult = driveRequest.Upload();

                // Get the file ID
                fileId = driveRequest.ResponseBody?.Id;


                ////Updating the permission to anyone so we can use it in google slides
                //var permission = DriveService.Permissions.Create(new Permission()
                //{
                //    Type = "anyone",
                //    Role = "reader"

                //}, fileId).Execute();

                // Get the file link
                imageLink = $"https://drive.google.com/u/0/uc?id={fileId}";

                Console.WriteLine("File Link: " + imageLink);
            }


            Dictionary<string, string> imagePlaceholders = new Dictionary<string, string>
            {
                { "{{4_IMAGE}}", imageLink }
            };


            BatchUpdatePresentationRequest request = new BatchUpdatePresentationRequest();
            request.Requests = new List<Request>();

            // Replace all instances of text placeholders in the presentation.
            foreach (KeyValuePair<string, string> placeholder in textPlaceholders)
            {
                request.Requests.Add(new Request
                {
                    ReplaceAllText = new ReplaceAllTextRequest
                    {
                        ContainsText = new SubstringMatchCriteria() { MatchCase = true, Text = placeholder.Key },
                        ReplaceText = placeholder.Value
                    }
                });
            }

            // Replace all instances of image placeholders in the presentation.
            foreach (KeyValuePair<string, string> placeholder in imagePlaceholders)
            {
                request.Requests.Add(new Request
                {
                    ReplaceAllShapesWithImage = new ReplaceAllShapesWithImageRequest()
                    {
                        ContainsText = new SubstringMatchCriteria() { MatchCase = true, Text = placeholder.Key },
                        ImageUrl = placeholder.Value,
                        ImageReplaceMethod = "CENTER_INSIDE",
                    }
                });
            }

            //Send the request
            var response = SlidesService.Presentations.BatchUpdate(request, copyPresentationId).Execute();

            //Delete the image from google drive
            DriveService.Files.Delete(fileId).Execute();

        }


        [TestMethod]
        public void GetAllFolders()
        {

            FilesResource.ListRequest listRequest = DriveService.Files.List();

            listRequest.Q = "mimeType='application/vnd.google-apps.folder'";
            listRequest.Fields = "nextPageToken, files(id, name)";

            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;
            foreach (var file in files)
            {
                Console.WriteLine("Folder name: " + file.Name + ", ID: " + file.Id);
            }

        }


        [TestMethod]
        public void CreateSpreadSheet()
        {
            var spreadSheet = new Spreadsheet()
            {
                Properties = new SpreadsheetProperties()
                {
                    Title = "New spreadsheet from C#"
                }
            };

            var request = SheetsService.Spreadsheets.Create(spreadSheet);
            var response = request.Execute();

            Console.WriteLine(response.SpreadsheetUrl);
            Assert.IsTrue(response.SpreadsheetUrl != null);
        }

        [TestMethod]
        public void AddChartToSpreadSheet()
        {
            // Get the spreadsheet ID of the target spreadsheet
            var spreadsheetId = "1J7sP682rkpLtGiRXcxJVkeCGKhCd8iUNktRomp2iEM0";

            var request = SheetsService.Spreadsheets.Get(spreadsheetId).Execute();
            var json = System.Text.Json.JsonSerializer.Serialize(request.Sheets[0]);

            //var getRangeResponse = SheetsService.Spreadsheets.Values.Get(spreadsheetId, "Sheet1!A1:B4").Execute();
            //var values = getRangeResponse.Values;
            //var numRows = values.Count;
            //var numCols = values[0].Count;

            //var gridRange = new GridRange()
            //{
            //    SheetId = 0,
            //    StartRowIndex = 0,
            //    EndRowIndex = numRows,
            //    StartColumnIndex = 0,
            //    EndColumnIndex = numCols
            //};

            // Define the chart type and configuration
            var chartSpec = new ChartSpec()
            {
                Title = "Sample Chart",
                Subtitle = "Description",
                BasicChart = new BasicChartSpec()
                {
                    ChartType = "COLUMN",
                    HeaderCount = 1,
                    LegendPosition = "BOTTOM_LEGEND",
                    Axis = new List<BasicChartAxis>()
                    {
                        new BasicChartAxis()
                        {
                            Position = "BOTTOM_AXIS",
                            Title = "Year"
                        },
                        new BasicChartAxis()
                        {
                            Position = "LEFT_AXIS",
                            Title = "Value"
                        }
                    },
                    Domains = new List<BasicChartDomain>()
                    {
                        new BasicChartDomain()
                        {
                            Domain = new ChartData()
                            {
                                SourceRange = new ChartSourceRange()
                                {
                                    Sources = new List<GridRange>()
                                    {
                                        new GridRange()
                                        {
                                            SheetId = 843514080,
                                            EndColumnIndex = 1 ,
                                            EndRowIndex = 4,
                                            StartColumnIndex = 0,
                                            StartRowIndex = 0
                                        }
                                    }
                                },
                                AggregateType = "SUM"
                            }
                        }
                    },
                    Series = new List<BasicChartSeries>()
                    {
                        new BasicChartSeries() {
                            TargetAxis = "LEFT_AXIS",
                            Series = new ChartData()
                            {
                                SourceRange = new ChartSourceRange()
                                {
                                    Sources = new List<GridRange>()
                                    {
                                        new GridRange()
                                        {
                                            SheetId = 843514080,
                                            EndColumnIndex = 2,
                                            EndRowIndex = 4,

                                            StartColumnIndex = 1,
                                            StartRowIndex = 0
                                        }
                                    }
                                }
                            }
                        },
                        new BasicChartSeries() {
                            TargetAxis = "LEFT_AXIS",
                            Series = new ChartData()
                            {
                                SourceRange = new ChartSourceRange()
                                {
                                    Sources = new List<GridRange>()
                                    {
                                        new GridRange()
                                        {
                                            SheetId = 843514080,
                                            EndColumnIndex = 3,
                                            EndRowIndex = 4,

                                            StartColumnIndex = 2,
                                            StartRowIndex = 0
                                        }
                                    }
                                }
                            }
                        }
                    }
                },
            };

            // Define the data for the new series
            var newSeriesData = new List<IList<object>>()
            {
                new List<object>() { "Year", "Value" },
                new List<object>() { "[2020]", 100 },
                new List<object>() { "[2021]", 150 },
                new List<object>() { "[2022]", 200 },
            };

            // Add the new data to the spreadsheet
            var dataRange = new ValueRange()
            {
                Range = "Sheet2!A:B",
                Values = newSeriesData
            };

            //var valueRequest = SheetsService.Spreadsheets.Values.Update(dataRange, spreadsheetId, dataRange.Range);
            //valueRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            //valueRequest.Execute();


            SheetsService.Spreadsheets.Values.BatchUpdate(new BatchUpdateValuesRequest()
            {
                Data = new List<ValueRange>()
                {
                    dataRange
                },
                ValueInputOption = "USER_ENTERED"
            }, spreadsheetId).Execute();

            // Create the chart
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
                            AnchorCell = new GridCoordinate() { SheetId = 843514080 }
                        }
                    }

                }
            };

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Google.Apis.Sheets.v4.Data.Request>()
                {
                    new Google.Apis.Sheets.v4.Data.Request()
                    {
                        AddChart = addChartRequest,
                    },
                },
            };


            var batchUpdateResponse = SheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId).Execute();

        }

        [TestMethod]
        public void CrearSpreadSheet()
        {
            string spreadsheetId = "1J7sP682rkpLtGiRXcxJVkeCGKhCd8iUNktRomp2iEM0";
            string range = "Sheet1!A1:Z";

            // Create the update cells request
            var updateCellsRequest = new UpdateCellsRequest
            {
                Range = new GridRange { SheetId = 0, StartColumnIndex = 0, StartRowIndex = 0, EndColumnIndex = int.MaxValue, EndRowIndex = int.MaxValue },
                Fields = "userEnteredValue"
            };

            // Create the batch update request
            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Google.Apis.Sheets.v4.Data.Request>
                {
                    new Google.Apis.Sheets.v4.Data.Request
                    {
                        UpdateCells = updateCellsRequest
                    }
                }
            };

            // Execute the batch update request
            var request = SheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId);
            var response = request.Execute();
        }

        [TestMethod]
        public void CloneSlideAndPaste()
        {

            string sourcePresentationId = "1aWCcKOeWR6EvTB2ldZThn4fZVObbuoHoWfteISWPqxc";
            string targetPresentationId = "1SPIIB_vW1H_85CGFgQf5e56RBmF_LT1O__INLDckHV4";
            int sourceSlideIndex = 0; // Index of the source slide (zero-based)

            var sourcePresentation = SlidesService.Presentations.Get(sourcePresentationId).Execute();
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

            var batchUpdateResponse = SlidesService.Presentations.BatchUpdate(batchUpdateRequest, targetPresentationId).Execute();

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

                    SlidesService.Presentations.BatchUpdate(batchUpdateRequest2, targetPresentationId).Execute();
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

                    SlidesService.Presentations.BatchUpdate(batchUpdateRequest2, targetPresentationId).Execute();

                }


            }

            var targetPresentation = SlidesService.Presentations.Get(targetPresentationId).Execute();

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

                    SlidesService.Presentations.BatchUpdate(batchUpdateRequest3, targetPresentationId).Execute();

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

                    SlidesService.Presentations.BatchUpdate(batchUpdateRequest4, targetPresentationId).Execute();
                }


            }



        }

        [TestMethod]
        public void CloneSlideToNewPresentation()
        {
            var presentationIdA = "1aWCcKOeWR6EvTB2ldZThn4fZVObbuoHoWfteISWPqxc";
            var presentationIdB = "1YI_Ht2EhjALZj86t8oPed2O4FXbKukjAZ5-8_EvduKw";
            var slideId = "g253657358fe_0_1";

            var getRequest = SlidesService.Presentations.Get(presentationIdA);
            var presentation = getRequest.Execute();

            var slide = presentation.Slides.FirstOrDefault(s => s.ObjectId == slideId);

            if (slide != null)
            {
                var batchUpdateRequest = new BatchUpdatePresentationRequest
                {
                    Requests = new List<Request>
                    {
                        new Request
                        {
                            DuplicateObject = new DuplicateObjectRequest
                            {
                                ObjectId = slide.ObjectId,
                                ObjectIds = new Dictionary<string, string>
                                {
                                    { slide.ObjectId, Guid.NewGuid().ToString() }
                                }
                            }
                        }
                    }
                };

                var batchUpdateResponse = SlidesService.Presentations.BatchUpdate(batchUpdateRequest, presentationIdB).Execute();
            }

        }


        [TestMethod]
        public void ReorderSlides()
        {
            string presentationId = "1CQjIJLn3KbZOy22ZgHaCZiW_R0e0mkaHLGGpz62QyL0";
            string slideId = "g254835d2201_0_5"; // 
            int targetIndex = 3; // Índice donde deseas colocar la diapositiva movida

            // Crear la solicitud de actualización de la presentación
            var request = new UpdateSlidesPositionRequest
            {
                SlideObjectIds = new List<string> { slideId },
                InsertionIndex = targetIndex
            };

            var batchUpdateRequest = new BatchUpdatePresentationRequest
            {
                Requests = new List<Request>
                {
                    new Request
                    {
                            UpdateSlidesPosition = request
                    }
                }
            };


            // Enviar la solicitud a la API de Google Slides
            var updateRequest = SlidesService.Presentations.BatchUpdate(batchUpdateRequest, presentationId);
            updateRequest.Execute();
        }

        [TestMethod]
        public void AddPieChart()
        {

            //This chart uses 2 columns, the first column for categories and the second for the values

            var spreadsheetId = "1J7sP682rkpLtGiRXcxJVkeCGKhCd8iUNktRomp2iEM0";

            // Define the chart specification
            var chartSpec = new ChartSpec
            {
                Title = "My Pie Chart",
                Subtitle = "Subtitle",
                PieChart = new PieChartSpec
                {
                    /*BOTTOM_LEGEND: displays the legend below the chart.
                    LEFT_LEGEND: displays the legend on the left side of the chart.
                    RIGHT_LEGEND: displays the legend on the right side of the chart.
                    TOP_LEGEND: displays the legend above the chart.
                    NO_LEGEND: hides the legend from the chart.
                    LABELED_LEGEND: Use labels for each category*/
                    LegendPosition = "LABELED_LEGEND",
                    ThreeDimensional = false,
                    PieHole = 0.5,
                    Domain = new ChartData()
                    {
                        SourceRange = new ChartSourceRange()
                        {
                            Sources = new List<GridRange>()
                            {
                                new GridRange()
                                {
                                    EndColumnIndex = 1 ,
                                    EndRowIndex = 4,
                                    StartColumnIndex = 0,
                                    StartRowIndex = 0
                                }
                            }
                        },
                        //AggregateType = "SUM"
                    },
                    Series = new ChartData()
                    {
                        SourceRange = new ChartSourceRange()
                        {
                            Sources = new List<GridRange>()
                            {
                                new GridRange()
                                {
                                    EndColumnIndex = 2,
                                    EndRowIndex = 4,

                                    StartColumnIndex = 1,
                                    StartRowIndex = 0
                                }
                            }
                        }
                    }
                }
            };


            // Create the chart
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

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Google.Apis.Sheets.v4.Data.Request>()
                {
                    new Google.Apis.Sheets.v4.Data.Request()
                    {
                        AddChart = addChartRequest,
                    },
                },
            };


            var batchUpdateResponse = SheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId).Execute();


        }

        [TestMethod]
        public void AddStackedBarChart()
        {

            //This chart uses 2 columns, the first column for categories and the second for the values

            var spreadsheetId = "1J7sP682rkpLtGiRXcxJVkeCGKhCd8iUNktRomp2iEM0";

            // Define the chart specification
            var chartSpec = new ChartSpec
            {
                Title = "My stacked Chart",
                Subtitle = "Subtitle",
                BasicChart = new BasicChartSpec()
                {
                    ChartType = "COLUMN",
                    HeaderCount = 1,
                    StackedType = "STACKED",    // PERCENT_STACKED, STACKED, NOT_STACKED
                    LegendPosition = "BOTTOM_LEGEND",
                    Axis = new List<BasicChartAxis>()
                    {
                        new BasicChartAxis()
                        {
                            Position = "BOTTOM_AXIS",
                            Title = "Year"
                        },
                        new BasicChartAxis()
                        {
                            Position = "LEFT_AXIS",
                            Title = "Value"
                        }
                    },
                    Domains = new List<BasicChartDomain>()
                    {
                        new BasicChartDomain()
                        {
                            Domain = new ChartData()
                            {
                                SourceRange = new ChartSourceRange()
                                {
                                    Sources = new List<GridRange>()
                                    {
                                        new GridRange()
                                        {
                                            EndColumnIndex = 1 ,
                                            EndRowIndex = 4,
                                            StartColumnIndex = 0,
                                            StartRowIndex = 0
                                        }
                                    }
                                },
                                AggregateType = "SUM"
                            }
                        }
                    },
                    Series = new List<BasicChartSeries>()
                    {
                        new BasicChartSeries() {
                            TargetAxis = "LEFT_AXIS",
                            Series = new ChartData()
                            {
                                SourceRange = new ChartSourceRange()
                                {
                                    Sources = new List<GridRange>()
                                    {
                                        new GridRange()
                                        {
                                            EndColumnIndex = 2,
                                            EndRowIndex = 4,

                                            StartColumnIndex = 1,
                                            StartRowIndex = 0
                                        }
                                    }
                                }
                            }
                        },
                        new BasicChartSeries() {
                            TargetAxis = "LEFT_AXIS",
                            Series = new ChartData()
                            {
                                SourceRange = new ChartSourceRange()
                                {
                                    Sources = new List<GridRange>()
                                    {
                                        new GridRange()
                                        {
                                            EndColumnIndex = 3,
                                            EndRowIndex = 4,

                                            StartColumnIndex = 2,
                                            StartRowIndex = 0
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };


            // Create the chart
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

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Google.Apis.Sheets.v4.Data.Request>()
                {
                    new Google.Apis.Sheets.v4.Data.Request()
                    {
                        AddChart = addChartRequest,
                    },
                },
            };


            var batchUpdateResponse = SheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId).Execute();


        }

        [TestMethod]
        public void AddWaterfallChart()
        {
            var spreadsheetId = "1J7sP682rkpLtGiRXcxJVkeCGKhCd8iUNktRomp2iEM0";

            // Define the chart specification
            var chartSpec = new ChartSpec
            {
                Title = "My Waterfall Chart",
                Subtitle = "Subtitle",
                WaterfallChart = new WaterfallChartSpec()
                {

                }

            };


            // Create the chart
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

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Google.Apis.Sheets.v4.Data.Request>()
                {
                    new Google.Apis.Sheets.v4.Data.Request()
                    {
                        AddChart = addChartRequest,
                    },
                },
            };


            var batchUpdateResponse = SheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId).Execute();

        }

        [TestMethod]
        public void AddImageFullscreen()
        {

            string presentationId = "1aWCcKOeWR6EvTB2ldZThn4fZVObbuoHoWfteISWPqxc";
            string imageUrl = "https://googleslidesapi.azurewebsites.net/thumbnails/1yJzb5prVVdWZ4WAfKK6y0tOZ_pn0U8rR8KZ3UGvonfg/LARGE/slide0.png";


            var getRequest = SlidesService.Presentations.Get(presentationId);
            var presentation = getRequest.Execute();

            var createSlideRequest = new CreateSlideRequest
            {
                //InsertionIndex = presentation.Slides.Count,
                SlideLayoutReference = new LayoutReference { PredefinedLayout = "BLANK" }
            };
            var createSlideResponse = SlidesService.Presentations.BatchUpdate(new BatchUpdatePresentationRequest
            {
                Requests = new List<Request> { new Request { CreateSlide = createSlideRequest } }
            }, presentationId).Execute();

            // Obtener el ID de la nueva página creada
            var pageId = createSlideResponse.Replies[0].CreateSlide.ObjectId;

            // Realizar la solicitud para crear el shape en la nueva página
            SlidesService.Presentations.BatchUpdate(new BatchUpdatePresentationRequest
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
            }, presentationId).Execute();

        }

        [TestMethod]
        public void MakeReadOnly()
        {

            string presentationId = "1aWCcKOeWR6EvTB2ldZThn4fZVObbuoHoWfteISWPqxc";
            string slideId = "g1e468af8aa1_0_0";
            string elementObjectId = Guid.NewGuid().ToString();

            // Create a transparent shape element to cover the existing element
            var transparentShapeRequest = new CreateShapeRequest
            {
                ObjectId = elementObjectId,
                ShapeType = "TEXT_BOX",
                ElementProperties = new PageElementProperties
                {
                    PageObjectId = slideId,
                    Size = new Size
                    {
                        Height = new Dimension { Magnitude = 405.04, Unit = "PT" },
                        Width = new Dimension { Magnitude = 719.42, Unit = "PT" }
                    },
                    Transform = new AffineTransform
                    {
                        ScaleX = 1,
                        ScaleY = 1,
                        TranslateX = 0,
                        TranslateY = 0,
                        Unit = "PT"
                    }
                }
            };

            // Add the transparent shape request to the update requests
            var updateRequests = new List<Request>
            {
                new Request
                {
                    CreateShape = transparentShapeRequest
                }
            };

            var batchUpdateRequest = new BatchUpdatePresentationRequest
            {
                Requests = updateRequests
            };


            SlidesService.Presentations.BatchUpdate(batchUpdateRequest, presentationId).Execute();

        }

        [TestMethod]
        public void SendMail()
        {
            // Create email message.
            var to = "eladri-@live.com";
            var subject = "Test email";
            var body = "This is a test email sent from the Gmail API.";
            var email = new MimeKit.MimeMessage();
            email.From.Add(new MimeKit.MailboxAddress("Adrian", "velociraptor088@gmail.com"));
            email.To.Add(new MimeKit.MailboxAddress(Encoding.UTF8, "eladri", to));
            email.Subject = subject;
            email.Body = new MimeKit.TextPart("plain")
            {
                Text = body
            };
            var rawMessage = Convert.ToBase64String(Encoding.UTF8.GetBytes(email.ToString()));
            var message = new Message
            {
                Raw = rawMessage
            };

            // Send email message.
            var sendMessageRequest = GmailService.Users.Messages.Send(message, "me");
            var response = sendMessageRequest.Execute();
        }

        private (string? pageId, Shape shapeId) GetShapeId(string presentationId, string placeholderText)
        {
            //Get the presentation
            var presentation = SlidesService.Presentations.Get(presentationId).Execute();

            // Loop through the slides
            foreach (var page in presentation.Slides)
            {
                foreach (var element in page.PageElements)
                {
                    // Check if the element is a shape
                    if (element.Shape != null)
                    {
                        // Check if the shape has the specified placeholder text
                        if (element.Shape.Text != null &&
                            element.Shape.Text.TextElements[0].TextRun.Content.Contains(placeholderText))
                        {
                            // Return the ID of the shape
                            return (page.ObjectId, element.Shape);
                        }
                    }
                }
            }

            // Return null if the shape was not found
            return (null, null);
        }

        private UserCredential GetCredentials()
        {
            UserCredential credential;

            using (var stream = new FileStream("desk_veloci_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);

                credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                // Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }
    }
}