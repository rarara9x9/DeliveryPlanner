namespace DeliveryPlanner
{
    partial class Main
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.btnSetting = new System.Windows.Forms.Button();
            this.btnInputOrder = new System.Windows.Forms.Button();
            this.lblMaster = new System.Windows.Forms.Label();
            this.btnMakePlan = new System.Windows.Forms.Button();
            this.btnUploadDrive = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnSetting
            // 
            this.btnSetting.Image = ((System.Drawing.Image)(resources.GetObject("btnSetting.Image")));
            this.btnSetting.Location = new System.Drawing.Point(16, 408);
            this.btnSetting.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnSetting.Name = "btnSetting";
            this.btnSetting.Size = new System.Drawing.Size(47, 45);
            this.btnSetting.TabIndex = 0;
            this.btnSetting.TabStop = false;
            this.btnSetting.UseVisualStyleBackColor = true;
            this.btnSetting.Click += new System.EventHandler(this.btnSetting_Click);
            // 
            // btnInputOrder
            // 
            this.btnInputOrder.Font = new System.Drawing.Font("MS UI Gothic", 36F);
            this.btnInputOrder.Location = new System.Drawing.Point(16, 22);
            this.btnInputOrder.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnInputOrder.Name = "btnInputOrder";
            this.btnInputOrder.Size = new System.Drawing.Size(850, 122);
            this.btnInputOrder.TabIndex = 10;
            this.btnInputOrder.Text = "受注情報を入力する";
            this.btnInputOrder.UseVisualStyleBackColor = true;
            this.btnInputOrder.Click += new System.EventHandler(this.btnInputOrder_Click);
            // 
            // lblMaster
            // 
            this.lblMaster.AutoSize = true;
            this.lblMaster.Font = new System.Drawing.Font("MS UI Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblMaster.Location = new System.Drawing.Point(68, 438);
            this.lblMaster.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblMaster.Name = "lblMaster";
            this.lblMaster.Size = new System.Drawing.Size(75, 14);
            this.lblMaster.TabIndex = 11;
            this.lblMaster.Text = "MasterPath";
            // 
            // btnMakePlan
            // 
            this.btnMakePlan.Font = new System.Drawing.Font("MS UI Gothic", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnMakePlan.Location = new System.Drawing.Point(16, 149);
            this.btnMakePlan.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnMakePlan.Name = "btnMakePlan";
            this.btnMakePlan.Size = new System.Drawing.Size(850, 125);
            this.btnMakePlan.TabIndex = 12;
            this.btnMakePlan.Text = "生産計画表を作成する";
            this.btnMakePlan.UseVisualStyleBackColor = true;
            this.btnMakePlan.Click += new System.EventHandler(this.btnMakePlan_Click);
            // 
            // btnUploadDrive
            // 
            this.btnUploadDrive.Font = new System.Drawing.Font("MS UI Gothic", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnUploadDrive.Location = new System.Drawing.Point(16, 278);
            this.btnUploadDrive.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnUploadDrive.Name = "btnUploadDrive";
            this.btnUploadDrive.Size = new System.Drawing.Size(850, 125);
            this.btnUploadDrive.TabIndex = 13;
            this.btnUploadDrive.Text = "生産計画表をGoogleDriveに登録する";
            this.btnUploadDrive.UseVisualStyleBackColor = true;
            this.btnUploadDrive.Click += new System.EventHandler(this.btnUploadDrive_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(889, 458);
            this.Controls.Add(this.btnUploadDrive);
            this.Controls.Add(this.btnMakePlan);
            this.Controls.Add(this.lblMaster);
            this.Controls.Add(this.btnInputOrder);
            this.Controls.Add(this.btnSetting);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "Main";
            this.Text = "配送計画作成ツール";
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSetting;
        private System.Windows.Forms.Button btnInputOrder;
        private System.Windows.Forms.Label lblMaster;
        private System.Windows.Forms.Button btnMakePlan;
        private System.Windows.Forms.Button btnUploadDrive;
    }
}

