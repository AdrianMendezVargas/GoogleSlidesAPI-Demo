using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Slides.v1;
using Google.Apis.Slides.v1.Data;
using Google.Apis.Util.Store;

namespace GoogleSlide.Test
{
    [TestClass]
    public class GoogleSlidesTest
    {
        string[] Scopes = { SlidesService.Scope.Presentations, DriveService.Scope.Drive };
        UserCredential? Credential;
        DriveService DriveService = null!;
        SlidesService SlidesService = null!;

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
        }

        [TestMethod]
        public void CreateGoogleSlideAndAddText()
        {

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
                { "{{3_ITEM_THREE}}", "Market  Challenges" }

            };

            Dictionary<string, string> imagePlaceholders = new Dictionary<string, string>
            {
                { "{{4_IMAGE}}", "https://docpromo.in/wp-content/uploads/2022/02/Healthcare-Market-Segment-1-750x400.png" }
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

            var response = SlidesService.Presentations.BatchUpdate(request, copyPresentationId).Execute();
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

            using (var stream = new FileStream("slidestest_client_secret.json", FileMode.Open, FileAccess.Read))
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