using System;
using System.Windows.Forms;
using System.Configuration;

namespace DeliveryPlanner
{
    public partial class Setting : Form
    {
        public Setting()
        {
            InitializeComponent();
        }

        private void Setting_Load(object sender, EventArgs e)
        {
            var appSettings = ConfigurationManager.AppSettings;
            this.txtMasterPath.Text = appSettings["MasterPath"] ?? "";
            this.txtServiceAccountPath.Text = appSettings["ServiceAccountPath"] ?? "";
            this.txtOperationSheetId.Text = appSettings["OperationSheetId"] ?? "";
            this.txtOperationOrderSheetName.Text = appSettings["OperationOrderSheetName"] ?? "";
            this.txtOperationProcessSheetName.Text = appSettings["OperationProcessSheetName"] ?? "";
            this.txtOperationPlanSheetName.Text = appSettings["OperationPlanSheetName"] ?? "";
            this.txtOperationWorkerSheetName.Text = appSettings["OperationWorkerSheetName"] ?? "";
            this.txtTimeOffSheetId.Text = appSettings["TimeOffSheetId"] ?? "";
            this.txtTimeOffFormSheetName.Text = appSettings["TimeOffFormSheetName"] ?? "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(this.txtMasterPath.Text))
            {
                MessageBox.Show("マスタ情報 ファイルパスが空白です", "ERROR", MessageBoxButtons.OK);
                return;
            }
            if (string.IsNullOrWhiteSpace(this.txtServiceAccountPath.Text))
            {
                MessageBox.Show("接続情報情報 ファイルパスが空白です", "ERROR", MessageBoxButtons.OK);
                return;
            }
            if (string.IsNullOrWhiteSpace(this.txtOperationSheetId.Text))
            {
                MessageBox.Show("工程進捗管理表 シートIDが空白です", "ERROR", MessageBoxButtons.OK);
                return;
            }
            if (string.IsNullOrWhiteSpace(this.txtOperationOrderSheetName.Text))
            {
                MessageBox.Show("工程進捗管理表 シート名 受注管理台帳が空白です", "ERROR", MessageBoxButtons.OK);
                return;
            }
            if (string.IsNullOrWhiteSpace(this.txtOperationProcessSheetName.Text))
            {
                MessageBox.Show("工程進捗管理表 シート名 工程計画表が空白です", "ERROR", MessageBoxButtons.OK);
                return;
            }
            if (string.IsNullOrWhiteSpace(this.txtOperationPlanSheetName.Text))
            {
                MessageBox.Show("工程進捗管理表 シート名 生産計画表が空白です", "ERROR", MessageBoxButtons.OK);
                return;
            }
            if (string.IsNullOrWhiteSpace(this.txtOperationWorkerSheetName.Text))
            {
                MessageBox.Show("工程進捗管理表 シート名 作業者が空白です", "ERROR", MessageBoxButtons.OK);
                return;
            }
            if (string.IsNullOrWhiteSpace(this.txtTimeOffSheetId.Text))
            {
                MessageBox.Show("休業日報告 シートIDが空白です", "ERROR", MessageBoxButtons.OK);
                return;
            }
            if (string.IsNullOrWhiteSpace(this.txtTimeOffFormSheetName.Text))
            {
                MessageBox.Show("休業日報告 シート名 休業申請が空白です", "ERROR", MessageBoxButtons.OK);
                return;
            }

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["MasterPath"].Value = txtMasterPath.Text;
            config.AppSettings.Settings["ServiceAccountPath"].Value = txtServiceAccountPath.Text;
            config.AppSettings.Settings["OperationSheetId"].Value = txtOperationSheetId.Text;
            config.AppSettings.Settings["OperationOrderSheetName"].Value = txtOperationOrderSheetName.Text;
            config.AppSettings.Settings["OperationProcessSheetName"].Value = txtOperationProcessSheetName.Text;
            config.AppSettings.Settings["OperationPlanSheetName"].Value = txtOperationPlanSheetName.Text;
            config.AppSettings.Settings["OperationWorkerSheetName"].Value = txtOperationWorkerSheetName.Text;
            config.AppSettings.Settings["TimeOffSheetId"].Value = txtTimeOffSheetId.Text;
            config.AppSettings.Settings["TimeOffFormSheetName"].Value = txtTimeOffFormSheetName.Text;

            // 保存とリロード
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
            
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();

            //はじめのファイル名を指定する
            //はじめに「ファイル名」で表示される文字列を指定する
            ofd.FileName = this.txtMasterPath.Text;
            //はじめに表示されるフォルダを指定する
            //指定しない（空の文字列）の時は、現在のディレクトリが表示される
            ofd.InitialDirectory = System.IO.Path.GetDirectoryName(this.txtMasterPath.Text);
            //[ファイルの種類]に表示される選択肢を指定する
            //指定しないとすべてのファイルが表示される
            ofd.Filter = "Excelファイル(*.xlsx)|*.xlsx|すべてのファイル(*.*)|*.*";
            //[ファイルの種類]ではじめに選択されるものを指定する
            //2番目の「すべてのファイル」が選択されているようにする
            ofd.FilterIndex = 1;
            //タイトルを設定する
            ofd.Title = "ファイルを選択してください";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            ofd.RestoreDirectory = true;
            //存在しないファイルの名前が指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckFileExists = true;
            //存在しないパスが指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckPathExists = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                this.txtMasterPath.Text = ofd.FileName;
            }

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            //OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();

            //はじめのファイル名を指定する
            //はじめに「ファイル名」で表示される文字列を指定する
            ofd.FileName = this.txtServiceAccountPath.Text;
            //はじめに表示されるフォルダを指定する
            //指定しない（空の文字列）の時は、現在のディレクトリが表示される
            ofd.InitialDirectory = System.IO.Path.GetDirectoryName(this.txtServiceAccountPath.Text);
            //[ファイルの種類]に表示される選択肢を指定する
            //指定しないとすべてのファイルが表示される
            ofd.Filter = "サービスアカウントファイル(*.json)|*.json|すべてのファイル(*.*)|*.*";
            //[ファイルの種類]ではじめに選択されるものを指定する
            //2番目の「すべてのファイル」が選択されているようにする
            ofd.FilterIndex = 1;
            //タイトルを設定する
            ofd.Title = "ファイルを選択してください";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            ofd.RestoreDirectory = true;
            //存在しないファイルの名前が指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckFileExists = true;
            //存在しないパスが指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckPathExists = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                this.txtServiceAccountPath.Text = ofd.FileName;
            }
        }
    }
}
