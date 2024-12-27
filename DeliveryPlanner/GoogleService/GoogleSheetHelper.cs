using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using DeliveryPlanner.DataModel;
using DeliveryPlanner.Commons;

namespace DeliveryPlanner.GoogleService
{
    internal class GoogleSheetHelper
    {
        private readonly SheetsService _service;

        public GoogleSheetHelper()
        {
            // サービスアカウントの認証情報を取得
            string credentialPath = ConfigurationManager.AppSettings["ServiceAccountPath"] ?? "File None";
            GoogleCredential credential;
            using (var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(SheetsService.Scope.Spreadsheets);
            }

            // Sheets API サービスの初期化
            _service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = Const.ApplicationName,
            });

        }
        public async Task<List<TimeOff>> GetTimeOff()
        {
            // 3. スプレッドシート ID とシート名の設定
            var spreadsheetId = ConfigurationManager.AppSettings["TimeOffSheetId"] ?? "None";
            var sheetName = ConfigurationManager.AppSettings["TimeOffFormSheetName"] ?? "None"; // シート名を指定

            // 4. シート全体の有効データ範囲を取得
            var range = $"{sheetName}";  // 範囲を指定しない

            SpreadsheetsResource.ValuesResource.GetRequest request = _service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = await request.ExecuteAsync();
            var values = response.Values;

            var result = new List<TimeOff>();

            // 5. データを表示
            if (values != null && values.Count > 0)
            {
                var dataRows = values.Skip(1).ToList();  // 2行目以降がデータ行

                foreach (var row in dataRows)
                {
                    if (row[1] != null && row[2] != null && row[3] != null &&
                        DateTime.TryParse(row[2].ToString(), out DateTime startDate) &&
                        DateTime.TryParse(row[3].ToString(), out DateTime endDate) &&
                        startDate <= endDate && !(endDate < DateTime.Today))
                    {
                        result.Add(new TimeOff
                        {
                            Name = row[1].ToString(),
                            From = startDate,
                            To = endDate
                        });
                    }
                }
            }
            return result;
        }

        public async Task<List<OrderWorker>> GetOrderWorker()
        {
            // 3. スプレッドシート ID とシート名の設定
            var spreadsheetId = ConfigurationManager.AppSettings["OperationSheetId"] ?? "None";
            var sheetName = ConfigurationManager.AppSettings["OperationPlanSheetName"] ?? "None"; // シート名を指定

            // 4. シート全体の有効データ範囲を取得
            var range = $"{sheetName}";  // 範囲を指定しない

            SpreadsheetsResource.ValuesResource.GetRequest request = _service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = await request.ExecuteAsync();
            var values = response.Values;

            var result = new List<OrderWorker>();

            // 5. データを表示
            if (values != null && values.Count > 0)
            {
                var dataRows = values.Skip(1).ToList();  // 2行目以降がデータ行

                foreach (var row in dataRows)
                {                   
                    if (row[0] != null && row[1] != null && row[5] != null && row[11] != null && row[13] != null &&
                        row[14] != null && row[17] != null && row[20] != null && row[21] != null &&
                        (row[22] == null || string.IsNullOrWhiteSpace(row[22].ToString())) &&
                        (row[23] == null || string.IsNullOrWhiteSpace(row[23].ToString())) &&
                        int.TryParse(row[11].ToString(), out int containerNo) &&
                        DateTime.TryParse(row[20].ToString(), out DateTime startDate) &&
                        DateTime.TryParse(row[21].ToString(), out DateTime endDate) &&
                        startDate >= DateTime.Today)
                    {
                        result.Add(new OrderWorker(row[0].ToString(), row[1].ToString(), row[5].ToString(), containerNo, row[13].ToString(), row[14].ToString(), row[17].ToString(), startDate, endDate)); 
                    }
                }
            }
            return result;
        }

        // Google Sheets にデータを追記するメソッド
        public async Task AppendDataToGoogleSheet(string spreadsheetId, string sheetName, IList<IList<object>> tableData)
        {
            // 書き込みデータの作成
            var valueRange = new ValueRange
            {
                Values = tableData
            };

            // データの末尾に追記するリクエスト
            var appendRequest = _service.Spreadsheets.Values.Append(valueRange, spreadsheetId, sheetName);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            // リクエストの実行
            await appendRequest.ExecuteAsync();
            await ApplyBordersToUsedRange(spreadsheetId, sheetName);
        }

        public async Task<int> GetFirstEmptyRowAsync(string spreadsheetId, string range)
        {
            // A列のすべてのデータを取得
            var request = _service.Spreadsheets.Values.Get(spreadsheetId, range);
            var response = await request.ExecuteAsync();

            var values = response.Values;

            // 空白の行を探す
            if (values == null || values.Count == 0)
            {
                // A列が完全に空の場合、最初の行は1
                return 1;
            }

            // 最初の空白行を特定
            for (int i = 0; i < values.Count; i++)
            {
                if (values[i].Count == 0 || string.IsNullOrWhiteSpace(values[i][0]?.ToString()))
                {
                    return i + 1;  // 行番号は0ベースなので +1
                }
            }

            // 空白行が見つからない場合は、最終行の次の行を返す
            return values.Count + 1;
        }

        // シートの 2 行目以降を削除して新しいデータを書き込むメソッド
        public async Task ClearAndWriteNewData(string spreadsheetId, string sheetName, IList<IList<object>> tableData)
        {
            // スプレッドシートの ID を設定
            var useRow = await GetFirstEmptyRowAsync(spreadsheetId, $"{sheetName}!A:A");
            var range = $"{sheetName}!2:{useRow}";  // 2 行目以降を対象範囲に指定

            // 以降をクリア
            var clearRequest = _service.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadsheetId, range);
            await clearRequest.ExecuteAsync();

            // 新しいデータを書き込み
            var valueRange = new ValueRange
            {
                Values = tableData
            };

            var updateRequest = _service.Spreadsheets.Values.Update(valueRange, spreadsheetId, $"{sheetName}!A2");
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            await updateRequest.ExecuteAsync();
            await ApplyBordersToUsedRange(spreadsheetId, sheetName);
        }

        // 使用領域に罫線を引くメソッド
        public async Task ApplyBordersToUsedRange(string spreadsheetId, string sheetName)
        {
            // 1. 使用範囲の取得
            var response = await _service.Spreadsheets.Values.Get(spreadsheetId, $"{sheetName}").ExecuteAsync();
            int rowCount = response.Values.Count;
            int colCount = response.Values[0].Count;

            // 2. BatchUpdateリクエストの作成
            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>
            {
                new Request
                {
                    UpdateBorders = new UpdateBordersRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = await GetSheetIdByName(spreadsheetId, sheetName),
                            StartRowIndex = 0,
                            EndRowIndex = rowCount,
                            StartColumnIndex = 0,
                            EndColumnIndex = colCount
                        },
                        Top = new Border { Style = "SOLID", Width = 1, Color = new Color { Red = 0, Green = 0, Blue = 0 } },
                        Bottom = new Border { Style = "SOLID", Width = 1, Color = new Color { Red = 0, Green = 0, Blue = 0 } },
                        Left = new Border { Style = "SOLID", Width = 1, Color = new Color { Red = 0, Green = 0, Blue = 0 } },
                        Right = new Border { Style = "SOLID", Width = 1, Color = new Color { Red = 0, Green = 0, Blue = 0 } },
                        InnerHorizontal = new Border { Style = "SOLID", Width = 1, Color = new Color { Red = 0, Green = 0, Blue = 0 } },
                        InnerVertical = new Border { Style = "SOLID", Width = 1, Color = new Color { Red = 0, Green = 0, Blue = 0 } }
                    }
                }
            }
            };

            // 3. リクエスト送信
            var batchRequest = _service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId);
            await batchRequest.ExecuteAsync();

            Console.WriteLine("罫線が適用されました！");
        }

        // シート名から SheetId を取得するメソッド
        private async Task<int> GetSheetIdByName(string spreadsheetId, string sheetName)
        {
            var spreadsheet = await _service.Spreadsheets.Get(spreadsheetId).ExecuteAsync();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == sheetName);
            return sheet?.Properties.SheetId ?? 0;
        }
    }
}
