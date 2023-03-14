using Azure.Identity;
using HackTogetherMSGraphMAUI.Models;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace HackTogetherMSGraphMAUI.Services
{
    /// <summary>
    /// This provides access to Microsoft 365 with Microsoft Graph API
    /// </summary>
    public class GraphService
    {

        private const string TenantId = "<Add Tenant ID>";
        private const string ClientId = "<Add Client ID>";

        private readonly string[] _scopes = new[] { "User.Read", "Files.Read" };
        private GraphServiceClient _client;

        /// <summary>
        /// Default constructor
        /// </summary>
        public GraphService()
        {
            Initialize();
        }

        /// <summary>
        /// Initializes Graph Service Client
        /// </summary>
        private void Initialize()
        {
            // Assume Windows for this app
            if (OperatingSystem.IsWindows())
            {
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = TenantId,
                    ClientId = ClientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("https://localhost"),
                };

                InteractiveBrowserCredential interactiveCredential = new(options);
                _client = new GraphServiceClient(interactiveCredential, _scopes);
            }
            else
            {
                // No iOS/Android support
                throw new NotImplementedException();
            }
        }

        /// <summary>
        /// Gets list of Excel files from default users's drive
        /// </summary>
        /// <returns>List of Excel files</returns>
        public async Task<List<ExcelFile>> GetMyFilesAsync()
        {
            //Get Drive and Folder
            var drive = await _client.Me.Drive.GetAsync();
            var folder = await _client.Drives[drive.Id].Root.GetAsync();

            //Get Files
            var children = await _client.Drives[drive.Id].Items[folder.Id].Children.GetAsync(
                requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new string[] { "id", "name", "createdDateTime", "lastModifiedDateTime", "file" };
                }
            );

            //Filter only Excel files
            var files = children.Value
                .Where(x => x.FileObject != null)
                .Where(x => x.Name.EndsWith("xlsx"))
                .OrderByDescending(x => x.LastModifiedDateTime)
                .Take(10)
                .Select(x => new ExcelFile
                {
                    Id = x.Id,
                    Name = x.Name,
                    CreatedDateTime = x.CreatedDateTime?.ToString("g"),
                    LastModifiedDateTime = x.LastModifiedDateTime?.ToString("g")
                }).ToList();


            //Create requests to get number of Worksheets per Workbook
            var sheetsBatchRequest = new BatchRequestContent(_client);
            foreach (var file in files)
            {
                var sheetsRequest = _client.Drives[drive.Id].Items[file.Id]
                 .Workbook.Worksheets.ToGetRequestInformation(
                     requestConfiguration =>
                     {
                         requestConfiguration.QueryParameters.Select = new string[] { "name" };
                         requestConfiguration.QueryParameters.Count = true;
                     });

                var requestMessage = await _client.RequestAdapter.ConvertToNativeRequestAsync<HttpRequestMessage>(sheetsRequest);
                BatchRequestStep requestStep = new BatchRequestStep(file.Id, requestMessage, null);
                sheetsBatchRequest.AddBatchRequestStep(requestStep);
            }

            //Read responses
            BatchResponseContent response = await _client.Batch.PostAsync(sheetsBatchRequest);
            foreach (var file in files)
            {
                try
                {
                    var responseItem = await response.GetResponseByIdAsync<WorkbookWorksheetCollectionResponse>(file.Id);
                    file.Sheets = responseItem.Value.Count;
                }
                catch { }
            }


            //Create requests to get number of Tables per Workbook
            var tablesBatchRequest = new BatchRequestContent(_client);
            foreach (var file in files)
            {
                var tablesRequest = _client.Drives[drive.Id].Items[file.Id]
                  .Workbook.Tables.ToGetRequestInformation(
                       requestConfiguration =>
                       {
                           requestConfiguration.QueryParameters.Select = new string[] { "name" };
                           requestConfiguration.QueryParameters.Count = true;
                       });
                var requestMessage = await _client.RequestAdapter.ConvertToNativeRequestAsync<HttpRequestMessage>(tablesRequest);
                BatchRequestStep requestStep = new BatchRequestStep(file.Id, requestMessage, null);
                tablesBatchRequest.AddBatchRequestStep(requestStep);
            }

            //Read responses
            response = await _client.Batch.PostAsync(tablesBatchRequest);
            foreach (var file in files)
            {
                try
                {
                    var responseItem = await response.GetResponseByIdAsync<WorkbookTableCollectionResponse>(file.Id);
                    file.Tables = responseItem.Value.Count;
                }
                catch { }
            }


            return files;

        }
    }
}
