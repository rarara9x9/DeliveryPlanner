using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using DeliveryPlanner.Commons;
using DeliveryPlanner.ExcelDataLoader;
using DeliveryPlanner.ExcelDataModel;

namespace DeliveryPlanner.UseCase
{
    internal class InputOrder
    {
        public static string CreateFile()
        {

            string templateFile = Utilities.GetTemplateOrderPath();

            // Check if Template file exists, throw exception if not
            if (!File.Exists(templateFile))
            {
                throw new FileNotFoundException("テンプレートファイルが存在しません: " + templateFile);
            }

            // Check if Work directory exists, create if not
            if (!Directory.Exists(Utilities.GetWorkDir()))
            {
                Directory.CreateDirectory(Utilities.GetWorkDir());
            }

            // Search for existing files in Work directory
            var existingFiles = Utilities.GetOrderFile();

            if (existingFiles.Any())
            {
                DialogResult result = MessageBox.Show(
                    "作業中ファイルが存在します。新たに受注入力ファイルを作成しますか？",
                    "確認",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );

                if (result == DialogResult.Yes)
                {
                    // Delete existing files
                    foreach (var file in existingFiles)
                    {
                        File.Delete(file);
                    }
                }
                else
                {
                    // Return full path of the first existing file
                    return existingFiles.First();
                }
            }

            // Generate new file name with timestamp
            string newFileName = Const.TemplateOrder.Replace(".xlsx", $"_{DateTime.Now:yyyyMMdd-HHmmss}.xlsx");
            string newFilePath = Path.Combine(Utilities.GetWorkDir(), newFileName);

            List<ProductInfo> products = null;
            try
            {
                // ClosedXMLを使用してExcelを開く
                using (var workbook = new XLWorkbook(ConfigurationManager.AppSettings["MasterPath"] ?? "File None"))
                {
                    products = ProductLoader.FromExcel(workbook.Worksheet("M_商品"));
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"マスタファイルの読み込みで失敗しました。{ex.Message}");
            }

            try
            {
                // ClosedXMLを使用してExcelを開く
                using (var workbook = new XLWorkbook(templateFile))
                {
                    var worksheet = workbook.Worksheet("M_商品");
                    // A3からデータを書き込む
                    int startRow = 2;
                    int row = startRow;

                    foreach (var order in products)
                    {
                        worksheet.Cell(row, 1).Value = order.ProductId;
                        worksheet.Cell(row, 2).Value = order.Client;
                        worksheet.Cell(row, 3).Value = order.ProductName;
                        ++row;
                    }
                    workbook.SaveAs(newFilePath);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"受注入力ファイルの作成で失敗しました。{ex.Message}");
            }

            return newFilePath;
        }
    }
}
