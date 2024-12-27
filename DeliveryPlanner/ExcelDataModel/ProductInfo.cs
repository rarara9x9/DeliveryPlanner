using System;

namespace DeliveryPlanner.ExcelDataModel
{
    internal class ProductInfo
    {
        // プロパティ
        public string ProductId { get; set; }    // 商品ID
        public string Client { get; set; }       // 取引先
        public string ProductName { get; set; }  // 品名
        public int ContainerSplitCount { get; }     // コンテナ分割個数

        // コンストラクタ
        public ProductInfo(string productId, string client, string productName, int containerSplitCount)
        {
            ProductId = productId ?? throw new ArgumentNullException(nameof(productId), "商品IDは必須です。");
            Client = client ?? throw new ArgumentNullException(nameof(client), "取引先は必須です。");
            ProductName = productName ?? throw new ArgumentNullException(nameof(productName), "品名は必須です。");
            ContainerSplitCount = containerSplitCount >= 0 ? containerSplitCount : throw new ArgumentOutOfRangeException(nameof(containerSplitCount), "コンテナ分割個数は0以上である必要があります。");
        }

        // オーバーライド（デバッグや表示用）
        public override string ToString()
        {
            return $"商品ID: {ProductId}, 取引先: {Client}, 品名: {ProductName}, コンテナ分割個数: {ContainerSplitCount}";
        }
    }
}
