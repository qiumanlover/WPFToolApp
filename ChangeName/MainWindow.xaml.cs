using Microsoft.Win32;
using MyTool;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

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
        private int keyIndex = -1, num1 = -1, num2 = -1, num3 = -1;
        string[] arrKeys, arrName1, arrName2, arrName3;
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
            catch (Exception ex)
            {
                ShowError("读取配置文件config.txt错误：文件可能不存在\r\n修复配置文件后才能打开应用程序", ex);
                Application.Current.Shutdown();
            }
        }

        private void ShowMessage(string msg)
        {
            MessageBox.Show(this, msg, "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ShowError(string msg, Exception ex)
        {
            this.Log(msg + "\r\n" + ex.ToString());
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
            string path = ((string[])data)[0];
            if (Path.GetExtension(path).Equals(".xls") || Path.GetExtension(path).Equals(".xlsx"))
            {
                (sender as TextBox).Text = path;
            }
            else
            {
                ShowMessage("该文件不是Excel文件");
                e.Handled = true;
                return;
            }

        }

        private void txtExcelSource_Drop(object sender, DragEventArgs e)
        {
            ThreadPool.QueueUserWorkItem(new WaitCallback(InitTitle), txtExcelSource.Text);
            btnRefresh_Click(sender, e);
            ResetCache();
            new Thread(() => GC.Collect()).Start();
        }

        private void txtExcelSource_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "Excel文件|*.xls;*.xlsx"
            };
            TextBox box = sender as TextBox;
            if (dialog.ShowDialog() == true)
            {
                box.Text = dialog.FileName;
                try
                {
                    ThreadPool.QueueUserWorkItem(new WaitCallback(InitTitle), txtExcelSource.Text);
                }
                catch (Exception ex)
                {
                    ShowError("获取Excel标题失败", ex);
                }
                
                btnRefresh_Click(sender, e);
                ResetCache();
            }
            dialog = null;
            new Thread(() => GC.Collect()).Start();
        }

        private void InitTitle(object path)
        {
            try
            {
                eh = new ExcelHelper(path.ToString());
                this.cmbKeyToPair.Dispatcher.Invoke(new Action(() =>
                {
                    cmbKeyToPair.ItemsSource = eh.GetRowWithColumnIndex(eh.FirstRowNum);
                    cmbKeyToPair.SelectedIndex = 0;
                }));
            }
            catch (Exception ex)
            {
                eh = null;
                this.cmbKeyToPair.Dispatcher.Invoke(new Action(() =>
                {
                    cmbKeyToPair.ItemsSource = null;
                    cmbKeyToPair.SelectedIndex = -1;
                    ShowError(ex.Message, ex);
                }));
                ShowError("获取Excel标题失败", ex);
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btnReloadExcel_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtExcelSource.Text))
            {
                ShowWarn("没有拖入Excel文件");
                return;
            }
            InitTitle(txtExcelSource.Text);
            btnRefresh_Click(sender, e);
            ResetCache();
            new Thread(() => GC.Collect()).Start();
        }

        private void ResetCache()
        {
            this.keyIndex = -1;
            this.num1 = -1;
            this.num2 = -1;
            this.num3 = -1;
            this.arrKeys = null;
            this.arrName1 = null;
            this.arrName2 = null;
            this.arrName3 = null;
        }

        private void lbFileList_PreviewDrop(object sender, DragEventArgs e)
        {
            List<string> excludeFiles = new List<string>();
            string[] data = (string[])e.Data.GetData(DataFormats.FileDrop);
            List<string> files = this.GetAllFiles(data, this.cfgList, ref excludeFiles);
            this.fileList.AddRange(from file in files select new FileForRename(file));

            this.lbFileList.ItemsSource = null;
            this.lbFileList.ItemsSource = this.fileList;
            if (excludeFiles.Count > 0)
            {
                string str = string.Join(", ", from excFile in excludeFiles select Path.GetFileName(excFile));
                this.Log($"已排除的文件：\r\n{str}");
                ShowWarn($"已排除以下文件\r\n\r\n{str}");
            }
            excludeFiles.Clear();
            excludeFiles = null;
            data = null;
            new Thread(() => GC.Collect()).Start();
        }

        private List<string> GetAllFiles(string[] paths, List<string> extList, ref List<string> excludeFiles)
        {
            List<string> files = new List<string>();
            List<string> folders = new List<string>();
            List<string> subFolders = new List<string>();
            foreach (string path in paths)
            {
                if (File.Exists(path))
                {
                    if (extList.Contains(Path.GetExtension(path)))
                    {
                        files.Add(path);
                    }
                    else
                    {
                        excludeFiles.Add(path);
                    }
                }
                else
                {
                    folders.Add(path);
                }
            }
            while (folders.Count > 0)
            {
                foreach (string folder in folders)
                {
                    files.AddRange(Directory.GetFiles(folder).Where(file => extList.Contains(Path.GetExtension(file))));
                    excludeFiles.AddRange(Directory.GetFiles(folder).SkipWhile(file => extList.Contains(Path.GetExtension(file))));
                    subFolders.AddRange(Directory.GetDirectories(folder));
                }
                folders.Clear();
                folders.AddRange(subFolders);
                subFolders.Clear();
            }
            new Thread(() => GC.Collect()).Start();
            return files;
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
                new Thread(() => GC.Collect()).Start();
            }
        }

        private bool FileIsOK()
        {
            if (eh == null)
            {
                ShowWarn("Excel表格异常，没有工作表");
                return false;
            }
            if (this.fileList.Count <= 0)
            {
                ShowWarn("没有要匹配的文件");
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
                ShowWarn("至少选择一个标题作为新的文件名");
                return;
            }
            if (keyIndex != this.keyIndex)
            {
                this.arrKeys = eh.GetColumn(keyIndex, eh.FirstRowNum);
                this.keyIndex = keyIndex;
            }
            if (num1 != this.num1)
            {
                this.arrName1 = num1 <= -1 ? null : eh.GetColumn(num1, eh.FirstRowNum);
                this.num1 = num1;
            }
            if (num2 != this.num2)
            {
                this.arrName2 = num2 <= -1 ? null : eh.GetColumn(num2, eh.FirstRowNum);
                this.num2 = num2;
            }
            if (num3 != this.num3)
            {
                this.arrName3 = num3 <= -1 ? null : eh.GetColumn(num3, eh.FirstRowNum);
                this.num3 = num3;
            }
            string filter1 = txtFilter1.Text;
            string filter2 = txtFilter2.Text;
            List<string> repeatKeyList = new List<string>(0);
            repeatKeyList.AddRange(this.arrKeys.Except(this.arrKeys.Distinct()));
            bool repeatKeyFlag = false;
            StringBuilder builder = new StringBuilder(0);
            foreach (var item in fileList)
            {
                FileForRename file = item as FileForRename;
                for (int i = 1; i < this.arrKeys.Length; i++)
                {
                    if (file.NewFileName.ToUpper().Contains(this.arrKeys[i].ToUpper()))
                    {
                        if (!repeatKeyList.Contains(this.arrKeys[i]))
                        {
                            int rowIndex = eh.FirstRowNum + i;
                            file.ChangeName($"{arrName1?[i] ?? ""}{filter1}{arrName2?[i] ?? ""}{filter2}{arrName3?[i] ?? ""}");
                            break;
                        }
                        else
                        {
                            repeatKeyFlag = true;
                            builder.AppendLine(this.arrKeys[i]);
                            break;
                        }
                    }
                }
            }
            lbFileList.ItemsSource = null;
            lbFileList.ItemsSource = fileList;
            if (repeatKeyFlag)
            {
                ShowWarn(string.Format("以下键在Excel表格中存在多个值，无法修改其文件名, 请修改Excel表格中的键\r\n{0}", builder.ToString()));
            }
            repeatKeyList.Clear();
            repeatKeyList = null;
            new Thread(() => GC.Collect()).Start();
        }

        private void btnSaveFile_Click(object sender, RoutedEventArgs e)
        {
            if (!FileIsOK())
            {
                return;
            }
            StringBuilder sb = new StringBuilder(0);
            foreach (FileForRename file in fileList)
            {
                if (file.needRename())
                {
                    try
                    {
                        file.Rename();
                    }
                    catch (Exception ex)
                    {
                        this.Log(ex.ToString());
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
            sb.Clear();
            new Thread(() => GC.Collect()).Start();
        }

        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            this.cmbName1.SelectedIndex = -1;
            this.cmbName2.SelectedIndex = -1;
            this.cmbName3.SelectedIndex = -1;
        }

        private void btnMark_Click(object sender, RoutedEventArgs e)
        {
            new OfficeHelper.ExcelOLEHelper().ReadExcel();
            if (!this.FileIsOK())
            {
                return;
            }
            int keyIndex = Convert.ToInt32(this.cmbKeyToPair.SelectedValue);
            string[] arrKeys;
            if (keyIndex != this.keyIndex)
            {
                arrKeys = eh.GetColumn(keyIndex);
            }
            else
            {
                arrKeys = new string[this.arrKeys.Length];
                this.arrKeys.CopyTo(arrKeys, 0);
            }
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
            var res = MessageBox.Show(this, "标记完成，请选择保存方式\r\n\r\n是 -- 覆盖原文件\r\n否 -- 生成新文件", "请选择", MessageBoxButton.YesNo, MessageBoxImage.Question);
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
                ShowError("文件保存异常，请查看日志文件log.log", ex);
            }
            new Thread(() => GC.Collect()).Start();
        }

        private void btnUndo_Click(object sender, RoutedEventArgs e)
        {
            var selectFiles = lbFileList.SelectedItems;
            if (selectFiles.Count <= 0)
            {
                var result = MessageBox.Show(this, "全部撤销？", "提示", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    foreach (var item in this.fileList)
                    {
                        FileForRename file = item as FileForRename;
                        file.ResetInfo();
                    }
                }
            }
            else
            {
                foreach (var item in selectFiles)
                {
                    FileForRename file = item as FileForRename;
                    file.ResetInfo();
                }
            }
            lbFileList.ItemsSource = null;
            lbFileList.ItemsSource = this.fileList;
        }
    }
}
