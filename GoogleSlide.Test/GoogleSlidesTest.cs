using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Slides.v1;
using Google.Apis.Slides.v1.Data;
using Google.Apis.Util.Store;
using System;
using System.Net.Http.Headers;
using System.Security;
using Request = Google.Apis.Slides.v1.Data.Request;

namespace GoogleSlide.Test
{
    [TestClass]
    public class GoogleSlidesTest
    {
        string[] Scopes = { SlidesService.Scope.Presentations, DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets };
        UserCredential? Credential;
        DriveService DriveService = null!;
        SlidesService SlidesService = null!;
        SheetsService SheetsService = null!;

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
                            AnchorCell =  new GridCoordinate() { SheetId = 843514080 }
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
        public void AddChartFromTable()
        {

            //// Replace with your desired chart data range
            //string dataRange = "Sheet1!A1:B10";


            //// Create a new spreadsheet
            //Spreadsheet newSpreadsheet = new Spreadsheet();
            //newSpreadsheet.Properties = new SpreadsheetProperties();
            //newSpreadsheet.Properties.Title = "New Spreadsheet";
            //Spreadsheet createdSpreadsheet = SheetsService.Spreadsheets.Create(newSpreadsheet).Execute();

            //// Get the ID of the first sheet in the new spreadsheet
            //int sheetId = createdSpreadsheet.Sheets[0].Properties.SheetId ?? 0;

            //// Create the chart data source
            //DataSource chartDataSource = new DataSource();
            //chartDataSource.Range = new GridRange { SheetId = sheetId, StartRowIndex = 0, EndRowIndex = 10, StartColumnIndex = 0, EndColumnIndex = 2 };

            //// Create the chart specifications
            //ChartSpec chartSpec = new ChartSpec();
            //chartSpec.Title = "My Chart";
            //chartSpec.BasicChart = new BasicChart();
            //chartSpec.BasicChart.ChartType = "COLUMN";
            //chartSpec.BasicChart.DataSource = chartDataSource;

            //// Create the chart request
            //AddChartRequest addChartRequest = new AddChartRequest();
            //addChartRequest.Chart = new EmbeddedChart();
            //addChartRequest.Chart.Spec = chartSpec;
            //addChartRequest.TargetSheetId = sheetId;

            //// Create the batch update request
            //BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            //{
            //    Requests = new List<Request> { new Request { AddChart = addChartRequest } }
            //};

            //// Execute the batch update request
            //BatchUpdateSpreadsheetResponse batchUpdateResponse = SheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, createdSpreadsheet.SpreadsheetId).Execute();

            //Console.WriteLine("Chart created successfully.");

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