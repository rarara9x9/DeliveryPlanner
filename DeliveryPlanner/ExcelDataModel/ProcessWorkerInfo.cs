using System;

namespace DeliveryPlanner.ExcelDataModel
{
    internal class ProcessWorkerInfo
    {
        // プロパテ
        public string ProductId { get; }              // 商品ID
        public string ProcessId { get; set; }         // 工程ID
        public string WorkerId { get; set; }          // 作業者ID
        public int Priority { get; set; }             // 優先
        public decimal UnitPrice { get; set; }        // 単価
        public int MaxContainer { get; set; }      // 最大数(コンテナ/日)

        // コンストラクタ
        public ProcessWorkerInfo(string productId, string processId, string workerId, int priority, decimal unitPrice, int maxContainerPerDay)
        {
            ProductId = productId ?? throw new ArgumentNullException(nameof(productId), "商品IDは必須です。");
            ProcessId = processId ?? throw new ArgumentNullException(nameof(processId), "工程IDは必須です。");
            WorkerId = workerId ?? throw new ArgumentNullException(nameof(workerId), "作業者IDは必須です。");
            Priority = priority >= 0 ? priority : throw new ArgumentOutOfRangeException(nameof(priority), "優先は0以上の整数である必要があります。");
            UnitPrice = unitPrice >= 0 ? unitPrice : throw new ArgumentOutOfRangeException(nameof(unitPrice), "単価は0以上である必要があります。");
            MaxContainer = maxContainerPerDay >= 0 ? maxContainerPerDay : throw new ArgumentOutOfRangeException(nameof(maxContainerPerDay), "最大数(コンテナ/日)は0以上である必要があります。");
        }

        // オーバーライド（デバッグや表示用）
        public override string ToString()
        {
            return $"商品ID: {ProductId}, 工程ID: {ProcessId}, 作業者ID: {WorkerId}, 優先: {Priority}, 単価: {UnitPrice:C}, 最大数(コンテナ/日): {MaxContainer}";
        }
    }
}
