using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DeliveryPlanner.ExcelDataModel;

namespace DeliveryPlanner.ExcelDataLoader
{
    internal class ProcessLoader
    {
        public static List<ProcessInfo> FromExcel(IXLWorksheet worksheet)
        {
            var processes = new List<ProcessInfo>();

            var rows = worksheet.RowsUsed().Skip(1); // 最初の行はヘッダー行をスキップ

            // 各行をループしてデータを取得し、クラスのインスタンスに変換
            foreach (var row in rows)
            {
                if (!string.IsNullOrWhiteSpace(row.Cell(2).GetValue<string>()))
                {
                    var process = new ProcessInfo(
                        processId: row.Cell(1).GetValue<string>(),
                        processName: row.Cell(2).GetValue<string>()
                    );
                    processes.Add(process);
                }
            }
            return processes;
        }
    }
}
