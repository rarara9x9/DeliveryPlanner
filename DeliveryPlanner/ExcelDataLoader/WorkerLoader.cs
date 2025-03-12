using ClosedXML.Excel;
using DeliveryPlanner.ExcelDataModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DeliveryPlanner.ExcelDataLoader
{
    internal class WorkerLoader
    {
        public static List<WorkerInfo> FromExcel(IXLWorksheet worksheet)
        {
            var workers = new List<WorkerInfo>();

            var rows = worksheet.RowsUsed().Skip(1); // 最初の行はヘッダー行をスキップ

            // 各行をループしてデータを取得し、クラスのインスタンスに変換
            foreach (var row in rows)
            {
                if (!string.IsNullOrWhiteSpace(row.Cell(2).GetValue<string>()))
                {
                    var worker = new WorkerInfo(
                        workerId: row.Cell(1).GetValue<string>(),
                        workerName: row.Cell(2).GetValue<string>(),
                        deliveryOrder: row.Cell(3).GetValue<int>(),
                        maxContainer: new List<(DayOfWeek dayOfweek, int Count)> { (DayOfWeek.Monday, row.Cell(4).GetValue<int>()),
                                                                                   (DayOfWeek.Tuesday, row.Cell(5).GetValue<int>()),
                                                                                   (DayOfWeek.Wednesday, row.Cell(6).GetValue<int>()),
                                                                                   (DayOfWeek.Thursday, row.Cell(7).GetValue<int>()),
                                                                                   (DayOfWeek.Friday, row.Cell(8).GetValue<int>()),
                                                                                   (DayOfWeek.Saturday, row.Cell(9).GetValue<int>()),
                                                                                   (DayOfWeek.Sunday, row.Cell(10).GetValue<int>())
                                                                                 }
                    );
                    workers.Add(worker);
                }

            }
            return workers;
        }
    }
}
