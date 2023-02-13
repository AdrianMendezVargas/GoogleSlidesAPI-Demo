using Google.Apis.Slides.v1;
using Google.Apis.Slides.v1.Data;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Util.Store;
using System.Net;
using Google.Apis.Drive.v3;

namespace GoogleSlidesExample
{
    class Program
    {

        static string[] Scopes = { SlidesService.Scope.Presentations, DriveService.Scope.Drive };
        static UserCredential? Credential;
        static DriveService DriveService = null!;
        static SlidesService SlidesService = null!;

        static void Main(string[] args)
        {
            InitVariables();

            Console.WriteLine("Enter the template presentation Id:");
            string? templatePresentationId = Console.ReadLine();

            if (string.IsNullOrEmpty(templatePresentationId))
            {
                throw new Exception("the template presentation Id can't be empty");
            }

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

        private static void InitVariables()
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

        private static UserCredential GetCredentials()
        {
            UserCredential credential;

            using (var stream = new FileStream("slidestest_client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);

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
