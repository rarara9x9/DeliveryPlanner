using ClosedXML.Excel;
using DeliveryPlanner.ExcelDataModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DeliveryPlanner.ExcelDataLoader
{
    internal class HolidayLoader
    {

        public static List<HolidayInfo> FromExcel(IXLWorksheet worksheet)
        {
            var holidays = new List<HolidayInfo>();

            var rows = worksheet.RowsUsed().Skip(1); // 最初の行はヘッダー行をスキップ

            // 各行をループしてデータを取得し、クラスのインスタンスに変換
            foreach (var row in rows)
            {
                if (!string.IsNullOrWhiteSpace(row.Cell(2).GetValue<string>()))
                {
                    var holiday = new HolidayInfo(
                        holidayDate: row.Cell(1).GetValue<DateTime>(),
                        holidayName: row.Cell(2).GetValue<string>()
                    );
                    holidays.Add(holiday);
                }
            }
            return holidays;
        }
    }
}
