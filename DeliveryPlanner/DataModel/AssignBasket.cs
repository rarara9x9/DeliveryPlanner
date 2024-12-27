namespace DeliveryPlanner.DataModel
{
    internal class AssignContainer
    {
        public string OrderNumber { get; }        // 受注番号
        public string OrderDetailNumber { get; }  // 受注詳細番号
        public int ContainerNo { get; }              // コンテナ番号
        public string ProcessSide { get; }        // 左右

        public AssignContainer(string orderNumber, string orderDetailNumber, int containerNo, string processSide)
        {
            OrderNumber = orderNumber;
            OrderDetailNumber = orderDetailNumber;
            ContainerNo = containerNo;
            ProcessSide = processSide;
        }
    }
}
