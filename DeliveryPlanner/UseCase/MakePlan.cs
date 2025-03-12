using ClosedXML.Excel;
using DeliveryPlanner.DataModel;
using DeliveryPlanner.Commons;
using DeliveryPlanner.ExcelDataLoader;
using DeliveryPlanner.ExcelDataModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using System.Threading.Tasks;
using DeliveryPlanner.GoogleService;

namespace DeliveryPlanner.UseCase
{
    internal class MakePlan
    {
        public async Task<string> CreateFile(string orderPath)
        {
            string templateFile = Utilities.GetTemplatePlanPath();

            // Check if Template file exists, throw exception if not
            if (!File.Exists(templateFile))
            {
                throw new FileNotFoundException("テンプレートファイルが存在しません: " + templateFile);
            }

            // Check if Work directory exists, create if not
            if (!Directory.Exists(Utilities.GetWorkDir()))
            {
                Directory.CreateDirectory(Utilities.GetWorkDir());
            }

            // Search for existing files in Work directory
            var existingFiles = Utilities.GetPlanFile();

            if (existingFiles.Any())
            {
                DialogResult result = MessageBox.Show(
                    "作業中ファイルが存在します。新たに生産計画ファイルを作成しますか？",
                    "確認",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );

                if (result == DialogResult.Yes)
                {
                    // Delete existing files
                    foreach (var file in existingFiles)
                    {
                        File.Delete(file);
                    }
                }
                else
                {
                    // Return full path of the first existing file
                    return existingFiles.First();
                }
            }

            List<WorkerInfo> workers = null;
            List<ProductInfo> products = null;
            List<HolidayInfo> holidays = null;
            List<OrderInfo> orders = null;
            List<ProcessInfo> processs = null;
            List<ProcessWorkerInfo> processWorkers = null;
            List<ProductProcessInfo> productProcessrs = null;
            try
            {
                // ClosedXMLを使用してExcelを開く
                using (var workbook = new XLWorkbook(ConfigurationManager.AppSettings["MasterPath"] ?? "File None"))
                {
                    workers = WorkerLoader.FromExcel(workbook.Worksheet("M_作業者"));
                    products = ProductLoader.FromExcel(workbook.Worksheet("M_商品"));
                    processs = ProcessLoader.FromExcel(workbook.Worksheet("M_工程"));
                    productProcessrs = ProductProcessLoader.FromExcel(workbook.Worksheet("M_加工"));
                    processWorkers = ProcessWorkerLoader.FromExcel(workbook.Worksheet("M_計画"));
                    holidays = HolidayLoader.FromExcel(workbook.Worksheet("M_休日"));
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"マスタファイルの読み込みで失敗しました。{ex.Message}");
            }

            try
            {
                // ClosedXMLを使用してExcelを開く
                using (var workbook = new XLWorkbook(orderPath))
                {
                    orders = OrderLoader.FromExcel(workbook.Worksheet("T_受注入力表"));
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"受注入力の読み込みで失敗しました。{ex.Message}");
            }

            try
            {
                var containers = (from order in orders
                                  join product in products
                                  on order.ProductId equals product.ProductId
                                  select new
                                  {
                                      order.OrderNumber,
                                      order.OrderDetailNumber,
                                      order.ProductId,
                                      order.Quantity,
                                      product.ContainerSplitCount
                                  })
                                  .SelectMany(record =>
                                  {
                                      int count = (int)Math.Ceiling((double)record.Quantity / record.ContainerSplitCount);
                                      return Enumerable.Range(1, count).Select(i =>
                                      new
                                      {
                                          record.OrderNumber,
                                          record.OrderDetailNumber,
                                          record.ProductId,
                                          ContainerNo = i,
                                          ItemCount = (record.Quantity > (record.ContainerSplitCount * i)) ? record.ContainerSplitCount : (record.Quantity - record.ContainerSplitCount * (i - 1)),
                                          ProcessSide = "左右"
                                      });
                                  });

                var containerProcesses = (from order in orders
                                          join container in containers
                                          on new { order.OrderNumber, order.OrderDetailNumber, order.ProductId } equals new { container.OrderNumber, container.OrderDetailNumber, container.ProductId }
                                          join productProcessr in productProcessrs
                                          on order.ProductId equals productProcessr.ProductId
                                          join productProcessrFirstLast in productProcessrs.GroupBy(item => item.ProductId)
                                                                                           .Select(g => new { ProductId = g.Key, FirstTaskOrder = g.Min(x => x.TaskOrder), LastTaskOrder = g.Max(x => x.TaskOrder) })
                                          on productProcessr.ProductId equals productProcessrFirstLast.ProductId
                                          join process in processs
                                          on productProcessr.ProcessId equals process.ProcessId
                                          select new
                                          {
                                              order.OrderNumber,
                                              order.OrderDetailNumber,
                                              order.WorkStartDate,
                                              order.ProductId,
                                              container.ContainerNo,
                                              container.ItemCount,
                                              container.ProcessSide,
                                              productProcessr.ProcessId,
                                              productProcessr.TaskOrder,
                                              productProcessr.IsSimultaneous,
                                              productProcessrFirstLast.FirstTaskOrder,
                                              productProcessrFirstLast.LastTaskOrder
                                          })
                    .OrderBy(x => x.OrderNumber.PadLeft(20, ' '))
                    .ThenBy(x => x.OrderDetailNumber.PadLeft(20, ' '))
                    .ThenBy(x => x.ContainerNo)
                    .ThenBy(x => x.ProcessSide)
                    .ThenBy(x => x.TaskOrder);

                // 営業日・作業者のリストを作成
                // 今日の日付
                DateTime startDate = DateTime.Today;
                // 1年後の日付
                DateTime endDate = startDate.AddYears(1);

                var holiday = holidays.Select(x => x.HolidayDate).OrderBy(x => x).ToList();
                var sheetApi = new GoogleSheetHelper();
                var workOffAll = await sheetApi.GetTimeOff();

                var companyWorkday = workers.Where(x => x.DeliveryOrder == Const.CompanyDeliveryOrder)
                                            .OrderByDescending(x => x.WorkerId).First();
                var workerManager = new List<WorkerManager>();
                foreach (var worker in workers)
                {
                    var processes = processWorkers.Where(x => x.WorkerId == worker.WorkerId);
                    var workOff = workOffAll.Where(x => x.Name.Equals(worker.WorkerName));

                    var skipDays = new List<DayOfWeek>();
                    for (DateTime date = startDate; date < endDate; date = date.AddDays(1))
                    {
                        var 配送数 = worker.MaxContainer.Where(x => x.dayOfweek == date.DayOfWeek).First().Count;

                        if (companyWorkday.MaxContainer.Where(x => x.dayOfweek == date.DayOfWeek).First().Count <= 0 ||
                             workOff.Where(x => x.From <= date && date <= x.To).Count() >= 1 ||
                            holiday.Contains(date))
                        {
                            // 会社(作業者IDの最大)の作業数が0の曜日
                            // 休業範囲
                            // 祝日・指定休日
                            var manager = new WorkerManager(worker.WorkerId, date, 0);

                            foreach (var process in processes)
                            {
                                manager.AssignProcesses.Add(new AssignProcess(process.ProductId, process.ProcessId, 0));
                            }
                            workerManager.Add(manager);

                            if (holiday.Contains(date) && 配送数 != 0)
                            {
                                // 祝日・指定休日
                                skipDays.Add(date.DayOfWeek);
                            }
                        }
                        else
                        {
                            if (配送数 != 0 || skipDays.Count == 0)
                            {
                                var manager = new WorkerManager(worker.WorkerId, date, 配送数);
                                foreach (var process in processes)
                                {
                                    manager.AssignProcesses.Add(new AssignProcess(process.ProductId, process.ProcessId, process.MaxContainer));
                                }
                                workerManager.Add(manager);
                            }
                            else
                            {
                                //配送数が0の曜日でも直前に祝日・指定休日がある場合は配送する
                                var manager = new WorkerManager(worker.WorkerId, date, worker.MaxContainer.Where(x => skipDays.Last() == x.dayOfweek).First().Count);
                                foreach (var process in processes)
                                {
                                    manager.AssignProcesses.Add(new AssignProcess(process.ProductId, process.ProcessId, process.MaxContainer));
                                }
                                workerManager.Add(manager);
                            }
                            skipDays.Clear();
                        }
                    }
                };

                var 割当済コンテナ = await sheetApi.GetOrderWorker();
                foreach (var orderWorker in 割当済コンテナ)
                {
                    var assignProcess = workerManager.Find(x => x.WorkDay == orderWorker.DeliveryDay && x.WorkerId == orderWorker.WorkerId)
                                                     .AssignProcesses.Find(x => x.ProductId == orderWorker.ProductId && x.ProcessId == orderWorker.ProcessId);
                    if (assignProcess == null)
                    {
                        // 存在しないプロセスが割り当たっている場合は追加する。最大割り当ては0
                        workerManager.Find(x => x.WorkDay == orderWorker.DeliveryDay && x.WorkerId == orderWorker.WorkerId).AssignProcesses.Add(new AssignProcess(orderWorker.ProductId, orderWorker.ProcessId, 0));
                        assignProcess = workerManager.Find(x => x.WorkDay == orderWorker.DeliveryDay && x.WorkerId == orderWorker.WorkerId)
                                                     .AssignProcesses.Find(x => x.ProductId == orderWorker.ProductId && x.ProcessId == orderWorker.ProcessId);
                    }
                    assignProcess.AssignContainers.Add(new AssignContainer(orderWorker.OrderNumber, orderWorker.OrderDetailNumber, orderWorker.ContainerNo, orderWorker.ProcessSide));
                }

                // コンテナに割り付ける
                var collectionDay = new DateTime(2000, 1, 1);
                var deliveryOrder = Const.CompanyDeliveryOrder;
                var orderWorkers = new List<OrderWorker>();
                foreach (var containerProcesse in containerProcesses)
                {
                    var workerList = (from processWorker in (processWorkers.Where(x => x.ProductId == containerProcesse.ProductId && x.ProcessId == containerProcesse.ProcessId))
                                      join worker in workers
                                      on processWorker.WorkerId equals worker.WorkerId
                                      select new
                                      {
                                          processWorker.WorkerId,
                                          processWorker.Priority,
                                          worker.DeliveryOrder
                                      });

                    if (!workerList.Any())
                    {
                        throw new Exception($"作業者の取得に失敗しました。商品ID:{containerProcesse.ProductId}/工程ID:{containerProcesse.ProcessId}");
                    }

                    if (containerProcesse.TaskOrder == containerProcesse.FirstTaskOrder)
                    {
                        // 最初の工程は、作業開始日を回収日にする
                        collectionDay = containerProcesse.WorkStartDate;
                    }

                    // 会社は出発・到着となるため置き換える
                    if (deliveryOrder == Const.CompanyDeliveryOrder)
                    {
                        deliveryOrder = 0;
                    }

                    // 担当者・配送日を取得する
                    var assignWorkerDay = workerManager.Where(x => x.ContainerCount > 0)
                                                       .Where(x => workerList.Where(y => y.WorkerId == x.WorkerId).Any())
                                                       .Where(x => x.WorkDay >= collectionDay.AddDays((workerList.Where(y => y.WorkerId == x.WorkerId).First().DeliveryOrder < deliveryOrder) ? 1 : 0))
                                                       .Where(x => x.AssignProcesses.Where(y => y.ProductId == containerProcesse.ProductId && y.ProcessId == containerProcesse.ProcessId)
                                                                                    .Where(y => y.ContainerCount > 0)
                                                                                    .Any())
                                                       .OrderBy(x => x.WorkDay)
                                                       .ThenBy(x => workerList.Where(y => y.WorkerId == x.WorkerId).First().Priority)
                                                       .ThenByDescending(x => x.ContainerCount)
                                                       .First();

                    // 回収日(次回配送日)
                    collectionDay = workerManager.Where(x => x.WorkerId == assignWorkerDay.WorkerId)
                                                 .Where(x => x.WorkDay > assignWorkerDay.WorkDay)
                                                 .First()
                                                 .WorkDay;

                    // 次配送開始位置
                    deliveryOrder = workerList.Where(x => x.WorkerId == assignWorkerDay.WorkerId).First().DeliveryOrder;

                    orderWorkers.Add(new OrderWorker(containerProcesse.OrderNumber, containerProcesse.OrderDetailNumber, containerProcesse.ProductId, containerProcesse.ContainerNo, containerProcesse.ProcessSide, containerProcesse.ProcessId, assignWorkerDay.WorkerId, assignWorkerDay.WorkDay, collectionDay));

                    var assignProcess = workerManager.Find(x => x.WorkDay == assignWorkerDay.WorkDay && x.WorkerId == assignWorkerDay.WorkerId)
                                                     .AssignProcesses.Find(x => x.ProductId == containerProcesse.ProductId && x.ProcessId == containerProcesse.ProcessId);
                    assignProcess.AssignContainers.Add(new AssignContainer(containerProcesse.OrderNumber, containerProcesse.OrderDetailNumber, containerProcesse.ContainerNo, containerProcesse.ProcessSide));

                };

                var writePlanWokers = (from orderWorker in orderWorkers
                                       join order in orders
                                       on new { orderWorker.OrderNumber, orderWorker.OrderDetailNumber, orderWorker.ProductId } equals
                                          new { order.OrderNumber, order.OrderDetailNumber, order.ProductId }
                                       join process in processs
                                         on orderWorker.ProcessId equals process.ProcessId
                                       join product in products
                                       on orderWorker.ProductId equals product.ProductId
                                       join worker in workers
                                         on orderWorker.WorkerId equals worker.WorkerId
                                       join container in containers
                                         on new { orderWorker.OrderNumber, orderWorker.OrderDetailNumber, orderWorker.ProductId, orderWorker.ContainerNo, orderWorker.ProcessSide } equals
                                            new { container.OrderNumber, container.OrderDetailNumber, container.ProductId, container.ContainerNo, container.ProcessSide }
                                       join productProcessr in productProcessrs
                                         on new { orderWorker.ProductId, orderWorker.ProcessId } equals new { productProcessr.ProductId, productProcessr.ProcessId }
                                       select new
                                       {
                                           orderWorker.OrderNumber,
                                           orderWorker.OrderDetailNumber,
                                           order.OrderDate,
                                           order.WorkStartDate,
                                           order.DueDate,
                                           orderWorker.ProductId,
                                           product.ProductName,
                                           order.LotNumber,
                                           order.Color,
                                           order.Size,
                                           order.Quantity,
                                           orderWorker.ContainerNo,
                                           container.ItemCount,
                                           orderWorker.ProcessSide,
                                           orderWorker.ProcessId,
                                           process.ProcessName,
                                           orderWorker.WorkerId,
                                           worker.WorkerName,
                                           orderWorker.DeliveryDay,
                                           orderWorker.CollectionDay,
                                           productProcessr.TaskOrder
                                       })
                                 .OrderBy(x => x.OrderNumber.PadLeft(20, ' '))
                                 .ThenBy(x => x.OrderDetailNumber.PadLeft(20, ' '))
                                 .ThenBy(x => x.ContainerNo)
                                 .ThenBy(x => x.ProcessSide)
                                 .ThenBy(x => x.TaskOrder)
                                 ;

                var writePlans = writePlanWokers.Select(x => new
                {
                    x.OrderNumber,
                    x.OrderDetailNumber,
                    x.OrderDate,
                    x.WorkStartDate,
                    x.DueDate,
                    x.ProductId,
                    x.ProductName,
                    x.LotNumber,
                    x.Color,
                    x.Size,
                    x.Quantity,
                    x.ProcessSide,
                    x.ProcessId,
                    x.ProcessName,
                    x.TaskOrder
                })
                   .Distinct()
                   .OrderBy(x => x.OrderNumber.PadLeft(20, ' '))
                   .ThenBy(x => x.OrderDetailNumber.PadLeft(20, ' '))
                   .ThenBy(x => x.ProcessSide)
                   .ThenBy(x => x.TaskOrder)
                 ;

                var writeOrders = writePlanWokers.Select(x => new
                {
                    x.OrderNumber,
                    x.OrderDetailNumber,
                    x.OrderDate,
                    x.WorkStartDate,
                    x.DueDate,
                    x.ProductId,
                    x.ProductName,
                    x.LotNumber,
                    x.Color,
                    x.Size,
                    x.Quantity
                })
                   .Distinct()
                   .OrderBy(x => x.OrderNumber.PadLeft(20, ' '))
                   .ThenBy(x => x.OrderDetailNumber.PadLeft(20, ' '))
                   ;

                // Generate new file name with timestamp
                string newFileName = Const.TemplatePlan.Replace(".xlsx", $"_{DateTime.Now:yyyyMMdd-HHmmss}.xlsx");
                string newFilePath = Path.Combine(Utilities.GetWorkDir(), newFileName);

                try
                {
                    // ClosedXMLを使用してExcelを開く
                    using (var workbook = new XLWorkbook(templateFile))
                    {
                        {
                            var worksheet = workbook.Worksheet("T_受注管理台帳");
                            // A3からデータを書き込む
                            int startRow = 2;
                            int row = startRow;

                            foreach (var order in writeOrders)
                            {
                                worksheet.Cell(row, 1).Value = order.OrderNumber;
                                worksheet.Cell(row, 2).Value = order.OrderDetailNumber;
                                worksheet.Cell(row, 3).Value = order.OrderDate;
                                worksheet.Cell(row, 4).Value = order.WorkStartDate;
                                worksheet.Cell(row, 5).Value = order.DueDate;
                                worksheet.Cell(row, 6).Value = order.ProductId;
                                worksheet.Cell(row, 7).Value = order.ProductName;
                                worksheet.Cell(row, 8).Value = order.LotNumber;
                                worksheet.Cell(row, 9).Value = order.Color;
                                worksheet.Cell(row, 10).Value = order.Size;
                                worksheet.Cell(row, 11).Value = order.Quantity;
                                ++row;
                            }
                        }

                        {
                            var worksheet = workbook.Worksheet("T_工程計画表");
                            // A3からデータを書き込む
                            int startRow = 2;
                            int row = startRow;

                            foreach (var plan in writePlans)
                            {
                                worksheet.Cell(row, 1).Value = plan.OrderNumber;
                                worksheet.Cell(row, 2).Value = plan.OrderDetailNumber;
                                worksheet.Cell(row, 3).Value = plan.OrderDate;
                                worksheet.Cell(row, 4).Value = plan.WorkStartDate;
                                worksheet.Cell(row, 5).Value = plan.DueDate;
                                worksheet.Cell(row, 6).Value = plan.ProductId;
                                worksheet.Cell(row, 7).Value = plan.ProductName;
                                worksheet.Cell(row, 8).Value = plan.LotNumber;
                                worksheet.Cell(row, 9).Value = plan.Color;
                                worksheet.Cell(row, 10).Value = plan.Size;
                                worksheet.Cell(row, 11).Value = plan.Quantity;
                                worksheet.Cell(row, 12).Value = plan.ProcessSide;
                                worksheet.Cell(row, 13).Value = plan.ProcessId;
                                worksheet.Cell(row, 14).Value = plan.ProcessName;
                                ++row;
                            }
                        }

                        {
                            var worksheet = workbook.Worksheet("T_生産計画表");
                            // A3からデータを書き込む
                            int startRow = 2;
                            int row = startRow;

                            foreach (var planWoker in writePlanWokers)
                            {
                                worksheet.Cell(row, 1).Value = planWoker.OrderNumber;
                                worksheet.Cell(row, 2).Value = planWoker.OrderDetailNumber;
                                worksheet.Cell(row, 3).Value = planWoker.OrderDate;
                                worksheet.Cell(row, 4).Value = planWoker.WorkStartDate;
                                worksheet.Cell(row, 5).Value = planWoker.DueDate;
                                worksheet.Cell(row, 6).Value = planWoker.ProductId;
                                worksheet.Cell(row, 7).Value = planWoker.ProductName;
                                worksheet.Cell(row, 8).Value = planWoker.LotNumber;
                                worksheet.Cell(row, 9).Value = planWoker.Color;
                                worksheet.Cell(row, 10).Value = planWoker.Size;
                                worksheet.Cell(row, 11).Value = planWoker.Quantity;
                                worksheet.Cell(row, 12).Value = planWoker.ContainerNo;
                                worksheet.Cell(row, 13).Value = planWoker.ItemCount;
                                worksheet.Cell(row, 14).Value = planWoker.ProcessSide;
                                worksheet.Cell(row, 15).Value = planWoker.ProcessId;
                                worksheet.Cell(row, 16).Value = planWoker.ProcessName;
                                worksheet.Cell(row, 17).Value = planWoker.TaskOrder;
                                worksheet.Cell(row, 18).Value = planWoker.WorkerId;
                                worksheet.Cell(row, 21).Value = planWoker.DeliveryDay;
                                worksheet.Cell(row, 22).Value = planWoker.CollectionDay;
                                ++row;
                            }
                        }

                        {
                            var worksheet = workbook.Worksheet("M_作業者");
                            // A3からデータを書き込む
                            int startRow = 2;
                            int row = startRow;

                            foreach (var woker in workers)
                            {
                                worksheet.Cell(row, 1).Value = woker.WorkerId;
                                worksheet.Cell(row, 2).Value = woker.WorkerName;
                                worksheet.Cell(row, 3).Value = woker.DeliveryOrder;
                                ++row;
                            }
                        }
                        workbook.SaveAs(newFilePath);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"生産計画表の作成で失敗しました。{ex.Message}");
                }

                return newFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception($"コンテナ割り付けで失敗しました。{ex.Message}");
            }

        }
    }
}
