using System;

namespace DeliveryPlanner.ExcelDataModel
{
    internal class ProductProcessInfo
    {
        // プロパティ
        public string ProductId { get; }         // 商品ID
        public string ProcessId { get; }         // 工程ID
        public int TaskOrder { get; }            // 作業順序
        public bool IsSimultaneous { get; }      // 左右同期

        // コンストラクタ
        public ProductProcessInfo(string productId, string processId, int taskOrder, string isSimultaneous)
        {
            ProductId = productId ?? throw new ArgumentNullException(nameof(productId), "商品IDは必須です。");
            ProcessId = processId ?? throw new ArgumentNullException(nameof(processId), "工程IDは必須です。");
            TaskOrder = taskOrder > 0 ? taskOrder : throw new ArgumentOutOfRangeException(nameof(taskOrder), "作業順序は正の整数である必要があります。");
            IsSimultaneous = isSimultaneous == "同期";
        }

        // オーバーライド（デバッグや表示用）
        public override string ToString()
        {
            return $"商品ID: {ProductId}, 工程ID: {ProcessId}, 作業順序: {TaskOrder}, 左右同期: {IsSimultaneous}";
        }
    }
}
