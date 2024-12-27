using System;

namespace DeliveryPlanner.DataModel
{
    internal class OrderWorker
    {
        public string OrderNumber { get; }        // 受注番号
        public string OrderDetailNumber { get; }  // 受注詳細番号
        public string ProductId { get; }          // 商品ID
        public int ContainerNo { get; }              // コンテナ番号
        public string ProcessSide { get; }        // 左右
        public string ProcessId { get; }          // 工程ID
        public string WorkerId { get; }           // 作業者ID
        public DateTime DeliveryDay { get; }      // 配送予定日
        public DateTime CollectionDay { get; }    // 回収予定日

        // コンストラクタ
        public OrderWorker(string orderNumber, string orderDetailNumber, string productId, int containerNo, string processSide, string processId, string workerId, DateTime deliveryDay, DateTime collectionDay)
        {
            OrderNumber = orderNumber ?? throw new ArgumentNullException(nameof(orderNumber), "受注番号は必須です。");
            OrderDetailNumber = orderDetailNumber ?? throw new ArgumentNullException(nameof(orderDetailNumber), "受注詳細番号は必須です。");
            ProductId = productId ?? throw new ArgumentNullException(nameof(productId), "商品IDは必須です。");
            ContainerNo = containerNo;
            ProcessSide = processSide ?? throw new ArgumentNullException(nameof(processSide), "左右は必須です。");
            ProcessId = processId ?? throw new ArgumentNullException(nameof(processId), "工程IDは必須です。");
            WorkerId = workerId ?? throw new ArgumentNullException(nameof(workerId), "作業者IDは必須です。");
            DeliveryDay = deliveryDay;
            CollectionDay = collectionDay;
        }

        // オーバーライド（デバッグや表示用）
        public override string ToString()
        {
            return $"受注番号: {OrderNumber}, 受注詳細番号: {OrderDetailNumber}, 商品ID: {ProductId}, コンテナ番号: {ContainerNo}, 左右: {ProcessSide}, 工程ID: {ProcessId}, 作業者ID: {WorkerId}, 配送予定日:  {DeliveryDay:yyyy-MM-dd}, 回収予定日:  {CollectionDay:yyyy-MM-dd}";
        }
    }
}
