using System;

namespace DeliveryPlanner.ExcelDataModel
{
    internal class OrderInfo
    {
        // プロパティ
        public string OrderNumber { get; set; }        // 受注番号
        public string OrderDetailNumber { get; set; }  // 受注詳細番号
        public DateTime OrderDate { get; set; }        // 発注日
        public DateTime WorkStartDate { get; set; }    // 作業開始日
        public DateTime DueDate { get; set; }          // 納期
        public string ProductId { get; set; }          // 商品ID
        public string LotNumber { get; set; }          // Lot番号
        public string Color { get; set; }              // 色
        public string Size { get; set; }               // サイズ
        public int Quantity { get; set; }              // 個数

        // コンストラクタ
        public OrderInfo(string orderNumber, string orderDetailNumber, DateTime orderDate, DateTime workStartDate, DateTime dueDate, string productId, string lotNumber, string color, string size, int quantity)
        {
            OrderNumber = orderNumber ?? throw new ArgumentNullException(nameof(orderNumber), "受注番号は必須です。");
            OrderDetailNumber = orderDetailNumber ?? throw new ArgumentNullException(nameof(orderDetailNumber), "受注詳細番号は必須です。");
            OrderDate = orderDate;
            WorkStartDate = workStartDate;
            DueDate = dueDate;
            ProductId = productId ?? throw new ArgumentNullException(nameof(productId), "商品IDは必須です。");
            LotNumber = lotNumber ?? throw new ArgumentNullException(nameof(lotNumber), "Lot番号は必須です。");
            Color = color ?? throw new ArgumentNullException(nameof(color), "色は必須です。");
            Size = size ?? throw new ArgumentNullException(nameof(size), "サイズは必須です。");
            Quantity = quantity >= 0 ? quantity : throw new ArgumentOutOfRangeException(nameof(quantity), "個数は0以上である必要があります。");
        }

        // オーバーライド（デバッグや表示用）
        public override string ToString()
        {
            return $"受注番号: {OrderNumber}, 受注詳細番号: {OrderDetailNumber}, 発注日: {OrderDate:yyyy-MM-dd}, 作業開始日: {WorkStartDate:yyyy-MM-dd}, 納期: {DueDate:yyyy-MM-dd}, 商品ID: {ProductId}, Lot番号: {LotNumber}, 色: {Color}, サイズ: {Size}, 個数: {Quantity}";
        }
    }
}
