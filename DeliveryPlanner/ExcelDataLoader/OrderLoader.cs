using ClosedXML.Excel;
using DeliveryPlanner.ExcelDataModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DeliveryPlanner.ExcelDataLoader
{
    internal class OrderLoader
    {
        public static List<OrderInfo> FromExcel(IXLWorksheet worksheet)
        {
            var orders = new List<OrderInfo>();

            var rows = worksheet.RowsUsed().Skip(2); // 最初の行はヘッダー行をスキップ

            // 各行をループしてデータを取得し、クラスのインスタンスに変換
            foreach (var row in rows)
            {
                if (!string.IsNullOrWhiteSpace(row.Cell(2).GetValue<string>()))
                {
                    var order = new OrderInfo(
                        orderNumber: row.Cell(1).GetValue<string>(),
                        orderDetailNumber: row.Cell(2).GetValue<string>(),
                        orderDate: row.Cell(3).GetValue<DateTime>(),
                        workStartDate: row.Cell(4).GetValue<DateTime>(),
                        dueDate: row.Cell(5).GetValue<DateTime>(),
                        productId: row.Cell(6).GetValue<string>(),
                        lotNumber: row.Cell(8).GetValue<string>(),
                        color: row.Cell(9).GetValue<string>(),
                        size: row.Cell(10).GetValue<string>(),
                        quantity: row.Cell(11).GetValue<int>()
                    );
                    orders.Add(order);
                }
            }
            return orders;
        }
    }
}
