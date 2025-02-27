using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace reportBuilder.classes
{
    public class GoogleSheetsService
    {
        private SheetsService _service;
        private string _spreadsheetId;

        public GoogleSheetsService(string credentialsPath, string spreadsheetId, string applicationName = "ReportBuilderGoogleSheetsAPI")
        {
            _spreadsheetId = spreadsheetId;

            string[] scopes = { SheetsService.Scope.Spreadsheets };

            UserCredential credential;
            using (var stream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            _service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });
        }

        public void InsertData(Dictionary<string, int> data, string range = "Sheet1!A1")
        {
            var valueRange = new ValueRange();

            var objectList = new List<IList<object>>();
            foreach (var kvp in data)
            {
                objectList.Add(new List<object> { kvp.Key, kvp.Value });
            }

            valueRange.Values = objectList;

            var updateRequest = _service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();

            Console.WriteLine("Data inserted successfully.");
        }
    }
}
