using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DeliveryPlanner.GoogleService;
using System.Configuration;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace DeliveryPlanner.UseCase
{
    internal class UploadPlan
    {
        public async Task Register(string planPath)
        {
            try
            {
                // データの準備（関数を含む）
                var newOrder = new List<IList<object>>();
                var newWorker = new List<IList<object>>();
                var newPlan = new List<IList<object>>();
                var newProcess = new List<IList<object>>();

                var sheetHelper = new GoogleSheetHelper();
                var planUseRow = await sheetHelper.GetFirstEmptyRowAsync(ConfigurationManager.AppSettings["OperationSheetId"], $"{ConfigurationManager.AppSettings["OperationPlanSheetName"]}!A:A");

                // ClosedXMLを使用してExcelを開く
                using (var workbook = new XLWorkbook(planPath))
                {
                    {
                        var worksheet = workbook.Worksheet("M_作業者");

                        var rows = worksheet.RowsUsed().Skip(1); // 最初の行はヘッダー行をスキップ
                        foreach (var row in rows)
                        {
                            if (!string.IsNullOrWhiteSpace(row.Cell(1).GetValue<string>()))
                            {

                                newWorker.Add(new List<object> {
                                    row.Cell(1).GetValue<string>(),
                                    row.Cell(2).GetValue<string>(),
                                    row.Cell(3).GetValue<int>()
                                });
                            }
                        }
                        await sheetHelper.ClearAndWriteNewData(ConfigurationManager.AppSettings["OperationSheetId"], ConfigurationManager.AppSettings["OperationWorkerSheetName"], newWorker);
                    }

                    {
                        var worksheet = workbook.Worksheet("T_生産計画表");

                        var rows = worksheet.RowsUsed().Skip(1); // 最初の行はヘッダー行をスキップ
                        var index = planUseRow;
                        var M_作業者 = ConfigurationManager.AppSettings["OperationWorkerSheetName"];
                        foreach (var row in rows)
                        {
                            if (!string.IsNullOrWhiteSpace(row.Cell(1).GetValue<string>()))
                            {
                                newPlan.Add(new List<object> {
                                        row.Cell(1).GetValue<string>(),
                                        row.Cell(2).GetValue<string>(),
                                        row.Cell(3).GetValue<DateTime>().ToString("yyyy/MM/dd"),
                                        row.Cell(4).GetValue<DateTime>().ToString("yyyy/MM/dd"),
                                        row.Cell(5).GetValue<DateTime>().ToString("yyyy/MM/dd"),
                                        row.Cell(6).GetValue<string>(),
                                        row.Cell(7).GetValue<string>(),
                                        row.Cell(8).GetValue<string>(),
                                        row.Cell(9).GetValue<string>(),
                                        row.Cell(10).GetValue<decimal >().ToString("F1"),
                                        row.Cell(11).GetValue<string>(),
                                        row.Cell(12).GetValue<string>(),
                                        row.Cell(13).GetValue<string>(),
                                        row.Cell(14).GetValue<string>(),
                                        row.Cell(15).GetValue<string>(),
                                        row.Cell(16).GetValue<string>(),
                                        row.Cell(17).GetValue<string>(),
                                        row.Cell(18).GetValue<string>(),
                                        $"=XLOOKUP(R{index},'{M_作業者}'!$A$2:$A${newWorker.Count() + 1},'{M_作業者}'!$B$2:$B${newWorker.Count() + 1})",
                                        $"=XLOOKUP(R{index},'{M_作業者}'!$A$2:$A${newWorker.Count() + 1},'{M_作業者}'!$C$2:$C${newWorker.Count() + 1})",
                                        row.Cell(21).GetValue<DateTime>().ToString("yyyy/MM/dd"),
                                        row.Cell(22).GetValue<DateTime>().ToString("yyyy/MM/dd"),
                                        "",
                                        "",
                                        $"=IF($X{index}<>\"\",\"回収済\",IF($W{index}<>\"\",\"回収予定\",\"配達予定\"))",
                                        $"=IF($X{index}<>\"\",\"\",IF($W{index}<>\"\",$V{index},$U{index}))"
                                    });
                                ++index;
                            }
                        }
                        await sheetHelper.AppendDataToGoogleSheet(ConfigurationManager.AppSettings["OperationSheetId"], ConfigurationManager.AppSettings["OperationPlanSheetName"], newPlan);
                    }

                    {
                        var worksheet = workbook.Worksheet("T_工程計画表");

                        var rows = worksheet.RowsUsed().Skip(1); // 最初の行はヘッダー行をスキップ
                        var T_生産計画表 = ConfigurationManager.AppSettings["OperationPlanSheetName"];
                        var lastPlanIndex = planUseRow + newPlan.Count() - 1;
                        var index = await sheetHelper.GetFirstEmptyRowAsync(ConfigurationManager.AppSettings["OperationSheetId"], $"{ConfigurationManager.AppSettings["OperationProcessSheetName"]}!A:A");
                        foreach (var row in rows)
                        {
                            if (!string.IsNullOrWhiteSpace(row.Cell(1).GetValue<string>()))
                            {
                                newProcess.Add(new List<object> {
                                        row.Cell(1).GetValue<string>(),
                                        row.Cell(2).GetValue<string>(),
                                        row.Cell(3).GetValue<DateTime>().ToString("yyyy/MM/dd"),
                                        row.Cell(4).GetValue<DateTime>().ToString("yyyy/MM/dd"),
                                        row.Cell(5).GetValue<DateTime>().ToString("yyyy/MM/dd"),
                                        row.Cell(6).GetValue<string>(),
                                        row.Cell(7).GetValue<string>(),
                                        row.Cell(8).GetValue<string>(),
                                        row.Cell(9).GetValue<string>(),
                                        row.Cell(10).GetValue<decimal >().ToString("F1"),
                                        row.Cell(11).GetValue<string>(),
                                        row.Cell(12).GetValue<string>(),
                                        row.Cell(13).GetValue<string>(),
                                        row.Cell(14).GetValue<string>(),
                                        $"=MINIFS('{T_生産計画表}'!$U$2:$U{lastPlanIndex},'{T_生産計画表}'!$A$2:$A{lastPlanIndex},$A{index},'{T_生産計画表}'!$B$2:$B{lastPlanIndex},$B{index},'{T_生産計画表}'!$N$2:$N{lastPlanIndex},$L{index},'{T_生産計画表}'!$O$2:$O{lastPlanIndex},$M{index})",
                                        $"=MAXIFS('{T_生産計画表}'!$V$2:$V{lastPlanIndex},'{T_生産計画表}'!$A$2:$A{lastPlanIndex},$A{index},'{T_生産計画表}'!$B$2:$B{lastPlanIndex},$B{index},'{T_生産計画表}'!$N$2:$N{lastPlanIndex},$L{index},'{T_生産計画表}'!$O$2:$O{lastPlanIndex},$M{index})",
                                        $"=MINIFS('{T_生産計画表}'!$W$2:$W{lastPlanIndex},'{T_生産計画表}'!$A$2:$A{lastPlanIndex},$A{index},'{T_生産計画表}'!$B$2:$B{lastPlanIndex},$B{index},'{T_生産計画表}'!$N$2:$N{lastPlanIndex},$L{index},'{T_生産計画表}'!$O$2:$O{lastPlanIndex},$M{index})",
                                        $"=MAXIFS('{T_生産計画表}'!$X$2:$X{lastPlanIndex},'{T_生産計画表}'!$A$2:$A{lastPlanIndex},$A{index},'{T_生産計画表}'!$B$2:$B{lastPlanIndex},$B{index},'{T_生産計画表}'!$N$2:$N{lastPlanIndex},$L{index},'{T_生産計画表}'!$O$2:$O{lastPlanIndex},$M{index})",
                                        $"=K{index}-(U{index}+V{index})",
                                        $"=K{index}-S{index}",
                                        $"=SUMIFS('{T_生産計画表}'!$M$2:$M{lastPlanIndex},'{T_生産計画表}'!$A$2:$A{lastPlanIndex},$A{index},'{T_生産計画表}'!$B$2:$B{lastPlanIndex},$B{index},'{T_生産計画表}'!$N$2:$N{lastPlanIndex},$L{index},'{T_生産計画表}'!$O$2:$O{lastPlanIndex},$M{index},'{T_生産計画表}'!$Y$2:$Y{lastPlanIndex},\"回収済\")",
                                        $"=SUMIFS('{T_生産計画表}'!$M$2:$M{lastPlanIndex},'{T_生産計画表}'!$A$2:$A{lastPlanIndex},$A{index},'{T_生産計画表}'!$B$2:$B{lastPlanIndex},$B{index},'{T_生産計画表}'!$N$2:$N{lastPlanIndex},$L{index},'{T_生産計画表}'!$O$2:$O{lastPlanIndex},$M{index},'{T_生産計画表}'!$Y$2:$Y{lastPlanIndex},\"回収予定\")",
                                    });
                                ++index;
                            }
                        }
                        await sheetHelper.AppendDataToGoogleSheet(ConfigurationManager.AppSettings["OperationSheetId"], ConfigurationManager.AppSettings["OperationProcessSheetName"], newProcess);
                    }

                    {
                        var worksheet = workbook.Worksheet("T_受注管理台帳");

                        var rows = worksheet.RowsUsed().Skip(1); // 最初の行はヘッダー行をスキップ
                        var T_生産計画表 = ConfigurationManager.AppSettings["OperationPlanSheetName"];
                        var lastPlanIndex = planUseRow + newPlan.Count() - 1;
                        var index = await sheetHelper.GetFirstEmptyRowAsync(ConfigurationManager.AppSettings["OperationSheetId"], $"{ConfigurationManager.AppSettings["OperationOrderSheetName"]}!A:A");
                        foreach (var row in rows)
                        {
                            if (!string.IsNullOrWhiteSpace(row.Cell(1).GetValue<string>()))
                            {
                                newOrder.Add(new List<object> {
                                        row.Cell(1).GetValue<string>(),
                                        row.Cell(2).GetValue<string>(),
                                        row.Cell(3).GetValue<DateTime>().ToString("yyyy/MM/dd"),
                                        row.Cell(4).GetValue<DateTime>().ToString("yyyy/MM/dd"),
                                        row.Cell(5).GetValue<DateTime>().ToString("yyyy/MM/dd"),
                                        row.Cell(6).GetValue<string>(),
                                        row.Cell(7).GetValue<string>(),
                                        row.Cell(8).GetValue<string>(),
                                        row.Cell(9).GetValue<string>(),
                                        row.Cell(10).GetValue<decimal >().ToString("F1"),
                                        row.Cell(11).GetValue<string>(),
                                        $"=MINIFS('{T_生産計画表}'!$U$2:$U{lastPlanIndex},'{T_生産計画表}'!$A$2:$A{lastPlanIndex},$A{index},'{T_生産計画表}'!$B$2:$B{lastPlanIndex},$B{index})",
                                        $"=MAXIFS('{T_生産計画表}'!$V$2:$V{lastPlanIndex},'{T_生産計画表}'!$A$2:$A{lastPlanIndex},$A{index},'{T_生産計画表}'!$B$2:$B{lastPlanIndex},$B{index})",
                                        $"=MAXIFS('{T_生産計画表}'!$X$2:$X{lastPlanIndex},'{T_生産計画表}'!$A$2:$A{lastPlanIndex},$A{index},'{T_生産計画表}'!$B$2:$B{lastPlanIndex},$B{index},'{T_生産計画表}'!$N$2:$N{lastPlanIndex},\"左\",'{T_生産計画表}'!$Q$2:$Q{lastPlanIndex},99,'{T_生産計画表}'!$Y$2:$Y{lastPlanIndex},\"回収済\")",
                                        $"=MAXIFS('{T_生産計画表}'!$M$2:$M{lastPlanIndex},'{T_生産計画表}'!$A$2:$A{lastPlanIndex},$A{index},'{T_生産計画表}'!$B$2:$B{lastPlanIndex},$B{index},'{T_生産計画表}'!$N$2:$N{lastPlanIndex},\"左\",'{T_生産計画表}'!$Q$2:$Q{lastPlanIndex},99,'{T_生産計画表}'!$Y$2:$Y{lastPlanIndex},\"回収済\")"
                                    });
                                ++index;
                            }
                        }
                        await sheetHelper.AppendDataToGoogleSheet(ConfigurationManager.AppSettings["OperationSheetId"], ConfigurationManager.AppSettings["OperationOrderSheetName"], newOrder);
                    }


                }
            }
            catch (Exception ex)
            {
                throw new Exception($"GoogleDriveへの登録に失敗しました。{ex.Message}");
            }

            return;
        }
    }
}