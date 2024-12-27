using System;

namespace DeliveryPlanner.ExcelDataModel
{
    internal class ProcessInfo
    {
        // プロパティ
        public string ProcessId { get; set; }   // 工程ID
        public string ProcessName { get; set; } // 工程名

        // コンストラクタ
        public ProcessInfo(string processId, string processName)
        {
            ProcessId = processId ?? throw new ArgumentNullException(nameof(processId), "工程IDは必須です。");
            ProcessName = processName ?? throw new ArgumentNullException(nameof(processName), "工程名は必須です。");
        }

        // オーバーライド（デバッグや表示用）
        public override string ToString()
        {
            return $"工程ID: {ProcessId}, 工程名: {ProcessName}";
        }
    }
}
