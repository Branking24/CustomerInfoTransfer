using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Auth.OAuth2;
using System.IO;
using System.Threading;

namespace CustomerEmailMapper
{
    class CsvClient
    {

        public async void AddToSheet(string spreadsheetID, UIForm form, List<CustomerInfoTemplate> allInfo, string cid, string secret, string sheetName)
        {
            try
            {
                UserCredential credential;
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                            new ClientSecrets
                                            {
                                                ClientId = cid,
                                                ClientSecret = secret
                            },
                            new[] { SheetsService.Scope.Spreadsheets },
                            "user",
                            CancellationToken.None,
                            new FileDataStore("Spreadhseets.ListMyLibrary"));
                var userService = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "CsvClient",
                });


                // How the input data should be interpreted.
                SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;

                // How the input data should be inserted.
                SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum insertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;

                IList<IList<Object>> values = new List<IList<Object>>();
                foreach (CustomerInfoTemplate t in allInfo)
                {
                    IList<Object> obj = new List<Object>();
                    obj.Add(t.name);
                    obj.Add(t.phone);
                    obj.Add(t.email);
                    obj.Add(t.profileUrl);
                    values.Add(obj);
                }
                ValueRange requestBody = new ValueRange()
                {
                    Values = values
                };

                SpreadsheetsResource.ValuesResource.AppendRequest request = userService.Spreadsheets.Values.Append(requestBody, spreadsheetID, sheetName);
                request.ValueInputOption = valueInputOption;
                request.InsertDataOption = insertDataOption;

                var response = request.Execute();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                form.WriteStatus(ex.Message);
                throw ex;
            }
        }
    }
}
