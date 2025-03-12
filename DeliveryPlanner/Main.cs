using DeliveryPlanner.Commons;
using DeliveryPlanner.UseCase;
using System;
using System.Configuration;
using System.Linq;
using System.Windows.Forms;

namespace DeliveryPlanner
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            this.lblMaster.Text = ConfigurationManager.AppSettings["MasterPath"] ?? "File None";
        }

        private void btnInputOrder_Click(object sender, EventArgs e)
        {
            try
            {
                // 開きたいExcelファイルのパスを指定
                var masterFilePath = ConfigurationManager.AppSettings["MasterPath"] ?? "File None";
                if (!System.IO.File.Exists(masterFilePath))
                {
                    MessageBox.Show($"指定したファイルが見つかりません。{masterFilePath}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var filePath = InputOrder.CreateFile();
                Utilities.RunExcel(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラー: {ex.Message}", "例外", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnMakePlan_Click(object sender, EventArgs e)
        {
            try
            {
                // 開きたいExcelファイルのパスを指定
                var masterFilePath = ConfigurationManager.AppSettings["MasterPath"] ?? "File None";
                if (!System.IO.File.Exists(masterFilePath))
                {
                    MessageBox.Show($"指定したファイルが見つかりません。{masterFilePath}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Search for existing files in Work directory
                var orderFiles = Utilities.GetOrderFile();

                if (orderFiles.Any())
                {
                    var orderFile = orderFiles.First();
                    var fileName = System.IO.Path.GetFileName(orderFile);
                    DialogResult result = MessageBox.Show(
                        $"{fileName}から配送計画書を作成しますか？",
                        "確認",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information
                    );

                    if (result == DialogResult.Yes)
                    {
                        // Return full path of the first existing file
                        var makePlan = new MakePlan();
                        var filePath = await makePlan.CreateFile(orderFile);
                        Utilities.MoveFileToBackup(orderFile);
                        Utilities.RunExcel(filePath);
                    }
                }
                else
                {
                    MessageBox.Show($"受注入力表が見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラー: {ex.Message}", "例外", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnUploadDrive_Click(object sender, EventArgs e)
        {
            try
            {
                // 開きたいExcelファイルのパスを指定
                var masterFilePath = ConfigurationManager.AppSettings["MasterPath"] ?? "File None";
                if (!System.IO.File.Exists(masterFilePath))
                {
                    MessageBox.Show($"指定したファイルが見つかりません。{masterFilePath}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Search for existing files in Work directory
                var planFiles = Utilities.GetPlanFile();

                if (planFiles.Any())
                {
                    var planFile = planFiles.First();
                    var fileName = System.IO.Path.GetFileName(planFile);
                    DialogResult result = MessageBox.Show(
                        $"{fileName}をGoogleDriveに登録しますか？",
                        "確認",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information
                    );

                    if (result == DialogResult.Yes)
                    {
                        // Return full path of the first existing file
                        var uploadPlan = new UploadPlan();
                        await uploadPlan.Register(planFile);
                        Utilities.MoveFileToBackup(planFile);
                    }

                    MessageBox.Show(
                         $"GoogleDriveに登録しました",
                        "確認",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                     );
                }
                else
                {
                    MessageBox.Show($"受注入力表が見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラー: {ex.Message}", "例外", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSetting_Click(object sender, EventArgs e)
        {
            var frmSetting = new DeliveryPlanner.Setting();
            frmSetting.ShowDialog();

            this.lblMaster.Text = ConfigurationManager.AppSettings["MasterPath"] ?? "File None";
        }
    }
}