using MyTool;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
//using System.Windows.Shapes;

namespace ChangeName
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public ExcelHelper eh;
        private List<string> cfgList = new List<string>();
        private List<FileForRename> fileList = new List<FileForRename>();
        private bool repeatKeyFlag;
        private List<string> repeatKeyList = new List<string>(0);
        public MainWindow()
        {
            InitializeComponent();
            this.GetConfig();
        }

        private void GetConfig()
        {
            try
            {
                this.cfgList.AddRange(File.ReadAllLines("config.txt", Encoding.Default));
            }
            catch (FileNotFoundException)
            {
                ShowError("读取配置文件config.txt错误：文件可能不存在\r\n请立即修复配置文件，否则无法拖入文件");
            }
            catch (Exception ex)
            {
                this.Log(ex.ToString());
            }
        }

        private void ShowMessage(string msg)
        {
            MessageBox.Show(this, msg, "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ShowError(string msg)
        {
            MessageBox.Show(this, msg, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void ShowWarn(string msg)
        {
            MessageBox.Show(this, msg, "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void Log(string str)
        {
            File.AppendAllText("log.log", $"{DateTime.Now.ToString()}\r\n{str}\r\n");
        }

        private void txtExcelSource_PreviewDrag(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        private void txtExcelSource_PreviewDrop(object sender, DragEventArgs e)
        {
            object data = e.Data.GetData(DataFormats.FileDrop);
            TextBox box = sender as TextBox;
            box.Text = ((string[])data)[0];
            this.InitTitle(box.Text);
        }

        private void InitTitle(string path)
        {
            eh = new ExcelHelper(path);
            cmbKeyToPair.ItemsSource = eh.GetRowWithColumnIndex(eh.FirstRowNum);
            cmbKeyToPair.SelectedIndex = 0;
        }

        private void btnReloadExcel_Click(object sender, RoutedEventArgs e)
        {
            eh = new ExcelHelper(txtExcelSource.Text);
            cmbKeyToPair.ItemsSource = eh.GetRowWithColumnIndex(eh.FirstRowNum);
            cmbKeyToPair.SelectedIndex = 0;
            btnRefresh_Click(sender, e);
            GC.Collect();
        }

        private void lbFileList_PreviewDrop(object sender, DragEventArgs e)
        {
            StringBuilder builder = new StringBuilder();
            string[] data = (string[])e.Data.GetData(DataFormats.FileDrop);
            if ((data.Length == 1) & Directory.Exists(data[0]))
            {
                data = Directory.GetFiles(data[0]);
            }
            for (int i = 0; i < data.Length; i++)
            {
                if (this.cfgList.Contains(Path.GetExtension(data[i])))
                {
                    this.fileList.Add(new FileForRename(data[i]));
                }
                else
                {
                    builder.AppendLine(Path.GetFileName(data[i]));
                }
            }
            this.lbFileList.ItemsSource = null;
            this.lbFileList.ItemsSource = this.fileList;
            string str = builder.ToString();
            if (!string.IsNullOrEmpty(str))
            {
                this.Log($"已排除的文件：\r\n{str}");
                ShowWarn($"已排除以下文件\r\n\r\n{str}");
            }
            builder.Clear();
            builder = null;
            str = null;
            GC.Collect();
        }

        private void lbFileList_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                foreach (FileForRename selectedFile in (sender as ListBox).SelectedItems)
                {
                    this.fileList.Remove(selectedFile);
                }
                this.lbFileList.ItemsSource = null;
                this.lbFileList.ItemsSource = this.fileList;
                GC.Collect();
            }
        }

        private bool FileIsOK()
        {
            if (eh == null)
            {
                ShowError("Excel表格异常，没有工作表");
                return false;
            }
            if (this.fileList.Count <= 0)
            {
                ShowError("没有要匹配的文件");
                return false;
            }
            return true;
        }

        private void btnChangeName_Click(object sender, RoutedEventArgs e)
        {
            if (!this.FileIsOK())
            {
                return;
            }
            int keyIndex = Convert.ToInt32(this.cmbKeyToPair.SelectedValue);
            int num1 = Convert.ToInt32(this.cmbName1.SelectedValue ?? -1);
            int num2 = Convert.ToInt32(this.cmbName2.SelectedValue ?? -1);
            int num3 = Convert.ToInt32(this.cmbName3.SelectedValue ?? -1);
            if (((num2 <= -1) && (num3 <= -1)) && (num1 <= -1))
            {
                ShowError("至少选择一个标题作为新的文件名");
            }
            else
            {
                string[] arrKeys = eh.GetColumn(keyIndex, eh.FirstRowNum + 1);
                string[] arrName1 = num1 <= -1 ? null : eh.GetColumn(num1, eh.FirstRowNum + 1);
                string[] arrName2 = num2 <= -1 ? null : eh.GetColumn(num2, eh.FirstRowNum + 1);
                string[] arrName3 = num3 <= -1 ? null : eh.GetColumn(num3, eh.FirstRowNum + 1);
                string filter1 = txtFilter1.Text;
                string filter2 = txtFilter2.Text;
                repeatKeyList.AddRange(arrKeys.Except(arrKeys.Distinct()));
                this.repeatKeyFlag = false;
                StringBuilder builder = new StringBuilder(0);
                foreach (var item in fileList)
                {
                    FileForRename file = item as FileForRename;
                    for (int i = 0; i < arrKeys.Length; i++)
                    {
                        if (file.NewFileName.ToUpper().Contains(arrKeys[i].ToUpper()))
                        {
                            if (!this.repeatKeyList.Contains(arrKeys[i]))
                            {
                                int rowIndex = eh.FirstRowNum + 1 + i;
                                file.ChangeName($"{arrName1?[i] ?? ""}{filter1}{arrName2?[i] ?? ""}{filter2}{arrName3?[i] ?? ""}");
                                break;
                            }
                            else
                            {
                                this.repeatKeyFlag = true;
                                builder.AppendLine(arrKeys[i]);
                                break;
                            }
                        }
                    }
                }
                lbFileList.ItemsSource = null;
                lbFileList.ItemsSource = fileList;
                if (this.repeatKeyFlag)
                {
                    ShowWarn(string.Format("以下键在Excel表格中存在多个值，无法修改其文件名, 请修改Excel表格中的键\r\n{0}", builder.ToString()));
                }
                arrKeys = null;
                arrName1 = null;
                arrName2 = null;
                arrName3 = null;
                repeatKeyList.Clear();
                GC.Collect();
            }
        }

        private void btnSaveFile_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder sb = new StringBuilder(0);
            foreach (FileForRename file in fileList)
            {
                if (file.needRename())
                {
                    if (!file.Rename())
                    {
                        sb.AppendLine($"{file.OldFileName} ---> {file.NewFileName}");
                    }
                }
            }
            if (sb.Length > 0)
            {
                ShowWarn($"出现错误, 以下名称修改失败: \r\n\r\n{sb.ToString()}");
            }
            else
            {
                ShowMessage("保存成功");
            }
            GC.Collect();
        }

        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            this.cmbName1.SelectedIndex = -1;
            this.cmbName2.SelectedIndex = -1;
            this.cmbName3.SelectedIndex = -1;
        }

        private void btnMark_Click(object sender, RoutedEventArgs e)
        {
            if (!this.FileIsOK())
            {
                return;
            }
            int keyIndex = Convert.ToInt32(this.cmbKeyToPair.SelectedValue);
            string[] arrKeys = eh.GetColumn(keyIndex);
            string markLabel = txtMarkLabel.Text;
            arrKeys[0] = txtNewTitle.Text;
            for (int i = 1; i < arrKeys.Length; i++)
            {
                string key = arrKeys[i];
                arrKeys[i] = string.Empty;
                foreach (FileForRename file in fileList)
                {
                    if (file.NewFileName.ToUpper().Contains(key.ToUpper()))
                    {
                        arrKeys[i] = markLabel;
                    }
                }
            }
            eh.AppendColumn(arrKeys);
            string savePath = txtExcelSource.Text;
            var res = MessageBox.Show(this, "标记完成，请选择保存方式\r\n是 -- 覆盖原文件\r\n否 -- 生成新文件", "请选择", MessageBoxButton.YesNo, MessageBoxImage.Question);
            try
            {
                if (res == MessageBoxResult.Yes)
                {
                    eh.Save(savePath, true);
                }
                else
                {
                    savePath = Path.Combine(Path.GetDirectoryName(savePath), $"{Path.GetFileNameWithoutExtension(savePath)}_{DateTime.Now.ToString("MM-dd_HH-mm-ss")}{Path.GetExtension(savePath)}");
                    eh.Save(savePath, true);
                }
                ShowMessage($"保存成功\r\n{savePath}");
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                ShowError("文件保存异常，请查看日志文件log.log");
            }
            GC.Collect();
        }
    }
}
