using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace DeliveryPlanner.Commons
{
    internal class Utilities
    {
        public static string GetTemplateOrderPath()
        {
            string templateFile = Path.Combine(GetTemplateDir(), Const.TemplateOrder);

            // Return full path of the first existing file
            return templateFile;
        }

        public static string GetTemplatePlanPath()
        {
            string templateFile = Path.Combine(GetTemplateDir(), Const.TemplatePlan);

            // Return full path of the first existing file
            return templateFile;
        }

        public static string GetTemplateDir()
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string templateDir = Path.Combine(exeDir, Const.TemplateDir);

            // Return full path of the first existing file
            return templateDir;
        }

        public static string GetWorkDir()
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string workDir = Path.Combine(exeDir, Const.WorkDir);

            // Return full path of the first existing file
            return workDir;
        }

        public static List<string> GetOrderFile()
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;

            // Search for existing files in Work directory
            var existingFiles = Directory.GetFiles(Const.WorkDir, Const.TemplateOrder.Replace(".xlsx", "*.xlsx")).OrderBy(x => x).ToList();

            // Return full path of the first existing file
            return existingFiles;
        }

        public static List<string> GetPlanFile()
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;

            // Search for existing files in Work directory
            var existingFiles = Directory.GetFiles(Const.WorkDir, Const.TemplatePlan.Replace(".xlsx", "*.xlsx")).OrderBy(x => x).ToList();

            // Return full path of the first existing file
            return existingFiles;
        }

        public static void RunExcel(string filePath)
        {
            // 開きたいExcelファイルのパスを指定
            if (!System.IO.File.Exists(filePath))
            {
                MessageBox.Show("指定したファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // Excel の実行可能ファイルパス
            string excelPath = GetExcelExecutablePath();
            if (string.IsNullOrEmpty(excelPath))
            {
                MessageBox.Show("Excel の実行可能ファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Excel を指定ファイルとともに起動
            Process.Start(new ProcessStartInfo
            {
                FileName = excelPath,
                Arguments = $"\"{filePath}\"", // 引数としてファイルパスを指定
                UseShellExecute = true
            });
        }

        /// <summary>
        /// 指定したファイルをバックアップフォルダに移動する
        /// </summary>
        /// <param name="sourceFilePath">移動元のファイルパス</param>
        public static void MoveFileToBackup(string sourceFilePath)
        {
            // ファイルが存在するか確認
            if (!File.Exists(sourceFilePath))
            {
                MessageBox.Show("指定したファイルが存在しません: " + sourceFilePath, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string backupFolderPath = Path.Combine(exeDir, Const.BackupDir);

            // バックアップフォルダが存在しない場合は作成
            if (!Directory.Exists(backupFolderPath))
            {
                Directory.CreateDirectory(backupFolderPath);
            }

            // 移動先のファイルパスを生成
            string fileName = Path.GetFileName(sourceFilePath);
            string destinationFilePath = Path.Combine(backupFolderPath, fileName);

            // ファイルを移動（同名ファイルが存在する場合は上書き）
            File.Move(sourceFilePath, destinationFilePath);
        }

        private static string GetExcelExecutablePath()
        {
            // レジストリから Excel のインストールパスを取得 (64bit/32bitに対応)
            string[] registryKeys = new[]
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
            };

            foreach (var key in registryKeys)
            {
                using (var regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(key))
                {
                    if (regKey != null)
                    {
                        var value = regKey.GetValue(null) as string;
                        if (!string.IsNullOrEmpty(value))
                        {
                            return value;
                        }
                    }
                }
            }

            return null;
        }
    }
}
