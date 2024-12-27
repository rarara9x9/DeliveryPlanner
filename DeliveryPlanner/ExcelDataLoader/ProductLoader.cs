using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DeliveryPlanner.ExcelDataModel;

namespace DeliveryPlanner.ExcelDataLoader
{
    internal class ProductLoader
    {
        public static List<ProductInfo> FromExcel(IXLWorksheet worksheet)
        {
            var products = new List<ProductInfo>();

            var rows = worksheet.RowsUsed().Skip(1); // 最初の行はヘッダー行をスキップ

            // 各行をループしてデータを取得し、クラスのインスタンスに変換
            foreach (var row in rows)
            {
                if (!string.IsNullOrWhiteSpace(row.Cell(3).GetValue<string>()))
                {
                    var product = new ProductInfo(
                        productId: row.Cell(1).GetValue<string>(),
                        client: row.Cell(2).GetValue<string>(),
                        productName: row.Cell(3).GetValue<string>(),
                        containerSplitCount: row.Cell(4).GetValue<int>()
                    );
                    products.Add(product);
                }
            }
            return products;
        }
    }
}
