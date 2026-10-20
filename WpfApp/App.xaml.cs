using System.Windows;
using System.Windows.Threading;

namespace AirDefenseLoad;

public partial class App : Application
{
    private static bool _errorShown;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // 兜底：任何没被捕获的 UI 异常都提示出来，而不是让进程直接闪退
        DispatcherUnhandledException += OnDispatcherUnhandledException;

        var window = new MainWindow();
        MainWindow = window;
        window.Show();
    }

    private static void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        e.Handled = true;
        if (_errorShown) return;
        _errorShown = true;

        MessageBox.Show(
            "程序遇到意外错误，已跳过该操作：\r\n\r\n" + e.Exception.Message,
            "错误",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
    }
}
