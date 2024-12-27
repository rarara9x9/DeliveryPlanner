using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DeliveryPlanner.ExcelDataModel;

namespace DeliveryPlanner.ExcelDataLoader
{
    internal class ProcessWorkerLoader
    {
        public static List<ProcessWorkerInfo> FromExcel(IXLWorksheet worksheet)
        {
            var processWorkerInfos = new List<ProcessWorkerInfo>();

            var rows = worksheet.RowsUsed().Skip(1); // 最初の行はヘッダー行をスキップ

            // 各行をループしてデータを取得し、クラスのインスタンスに変換
            foreach (var row in rows)
            {
                if (!string.IsNullOrWhiteSpace(row.Cell(3).GetValue<string>()))
                {
                    var processWorkerInfo = new ProcessWorkerInfo(
                        productId: row.Cell(1).GetValue<string>(),
                        processId: row.Cell(3).GetValue<string>(),
                        workerId: row.Cell(5).GetValue<string>(),
                        priority: row.Cell(7).GetValue<int>(),
                        unitPrice: row.Cell(8).GetValue<decimal>(),
                        maxContainerPerDay: row.Cell(10).GetValue<int>()
                    );
                    processWorkerInfos.Add(processWorkerInfo);
                }
            }
            return processWorkerInfos;
        }
    }
}
