using System.Collections.Generic;
using System.Linq;

namespace DeliveryPlanner.DataModel
{
    internal class AssignProcess
    {
        public string ProductId { get; }                       // 商品ID
        public string ProcessId { get; }                       // 工程ID
        public List<AssignContainer> AssignContainers { get; set; } = new List<AssignContainer>();
        public int MaxContainerCount { get; }                     // 宅配最大コンテナ数
        public int ContainerCount                                 // 割り当て可能コンテナ数
        {
            get => MaxContainerCount - AssignContainers?.Count() ?? 0;
        }

        public AssignProcess(string productId, string processId, int maxContainerCount)
        {
            ProductId = productId;
            ProcessId = processId;
            MaxContainerCount = maxContainerCount;
        }
    }
}
