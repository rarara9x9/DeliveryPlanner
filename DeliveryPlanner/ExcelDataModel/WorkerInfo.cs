using System;
using System.Collections.Generic;

namespace DeliveryPlanner.ExcelDataModel
{
    internal class WorkerInfo
    {
        // プロパティ
        public string WorkerId { get; }               // 作業者ID
        public string WorkerName { get; }             // 作業者名
        public int DeliveryOrder { get; }             // 配送順
        public List<(DayOfWeek dayOfweek, int Count)> MaxContainer { get; set; } // 最大数(コンテナ/日)

        // コンストラクタ
        public WorkerInfo(string workerId, string workerName, int deliveryOrder, List<(DayOfWeek dayOfweek, int Count)> maxContainer)
        {
            WorkerId = workerId ?? throw new ArgumentNullException(nameof(workerId));
            WorkerName = workerName ?? throw new ArgumentNullException(nameof(workerName));
            DeliveryOrder = deliveryOrder > 0 ? deliveryOrder : throw new ArgumentOutOfRangeException(nameof(deliveryOrder), "配送順は正の整数である必要があります。");
            MaxContainer = maxContainer.Count == 7  ? maxContainer : throw new ArgumentOutOfRangeException(nameof(maxContainer), "最大数(コンテナ/日)は7である必要があります。");
        }

        // オーバーライド（デバッグや表示用）
        public override string ToString()
        {
            return $"作業者ID: {WorkerId}, 作業者名: {WorkerName}, 配送順: {DeliveryOrder}, 最大数(コンテナ/日): {string.Join(",", MaxContainer)}";
        }
    }
}
