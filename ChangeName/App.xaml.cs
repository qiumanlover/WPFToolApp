using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;

namespace ChangeName
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        private void Application_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show(MainWindow, "程序启动失败，缺少必要的组件", "警告", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }
    }
}
