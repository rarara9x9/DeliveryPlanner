using System;
using System.Collections.Generic;
using System.Linq;

namespace DeliveryPlanner.DataModel
{
    internal class WorkerManager
    {
        public string WorkerId { get; }          // 作業者ID
        public DateTime WorkDay { get; }         // 作業日
        public List<AssignProcess> AssignProcesses { get; set; } = new List<AssignProcess>();
        public int MaxContainerCount { get; }    // 宅配最大コンテナ数
        public int ContainerCount                // 割り当て可能コンテナ数
        {
            get => MaxContainerCount - AssignProcesses?.Sum(process => process.AssignContainers.Count()) ?? 0;
        }

        public WorkerManager(string workerId, DateTime workDay, int maxContainerCount)
        {
            WorkerId = workerId;
            WorkDay = workDay;
            MaxContainerCount = maxContainerCount;
        }
    }
}
