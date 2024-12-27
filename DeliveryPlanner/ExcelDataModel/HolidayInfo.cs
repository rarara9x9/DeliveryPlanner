using System;

namespace DeliveryPlanner.ExcelDataModel
{
    internal class HolidayInfo
    {
        // プロパティ
        public DateTime HolidayDate { get; set; } // 休日
        public string HolidayName { get; set; }   // 名称

        // コンストラクタ
        public HolidayInfo(DateTime holidayDate, string holidayName)
        {
            HolidayDate = holidayDate;
            HolidayName = holidayName ?? throw new ArgumentNullException(nameof(holidayName), "名称は必須です。");
        }

        // オーバーライド（デバッグや表示用）
        public override string ToString()
        {
            return $"休日: {HolidayDate:yyyy-MM-dd}, 名称: {HolidayName}";
        }
    }
}
