using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DeliveryPlanner.ExcelDataModel;

namespace DeliveryPlanner.ExcelDataLoader
{
    internal class ProductProcessLoader
    {
        public static List<ProductProcessInfo> FromExcel(IXLWorksheet worksheet)
        {
            var productProcesses = new List<ProductProcessInfo>();

            var rows = worksheet.RowsUsed().Skip(1); // 最初の行はヘッダー行をスキップ

            // 各行をループしてデータを取得し、クラスのインスタンスに変換
            foreach (var row in rows)
            {
                if (!string.IsNullOrWhiteSpace(row.Cell(1).GetValue<string>()))
                {
                    var productProcess = new ProductProcessInfo(
                        productId: row.Cell(1).GetValue<string>(),
                        processId: row.Cell(3).GetValue<string>(),
                        taskOrder: row.Cell(5).GetValue<int>(),
                        isSimultaneous: row.Cell(6).GetValue<string>()
                    );
                    productProcesses.Add(productProcess);
                }
            }
            return productProcesses;
        }
    }
}
