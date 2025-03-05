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

        public void InsertData(Dictionary<string, int> data, string range, string updateRange)
        {
            var requset = _service.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = requset.Execute();
            var values = response.Values;

            if (values == null || !values.Any())
            {
                Console.WriteLine("No data found.");
                return;
            }

            var updateData = new List<IList<object>>();

            foreach (var row in values) 
            {
                if (row.Count > 0 && row[0] is string description && data.ContainsKey(description))
                {
                    var newRow = new List<object>() { data[description] };
                    updateData.Add(newRow);
                }
                else
                {
                    updateData.Add(new List<object> { ""});
                }
            }

            var updateValueRange = new ValueRange { Values = updateData};

            var updateRequest = _service.Spreadsheets.Values.Update(updateValueRange, _spreadsheetId, updateRange);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();

            Console.WriteLine("Data updated successfully.");
        }
    }
}
