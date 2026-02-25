using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

public static class MessageBoxHelper
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    private const uint BM_CLICK = 0x00F5;

    public static async Task ShowMessageAsync(string title, string message, int autoOkMs = 0)
    {
        var tcs = new TaskCompletionSource<bool>();

        var thread = new Thread(() =>
        {
            if (autoOkMs > 0)
            {
                var timer = new System.Windows.Forms.Timer { Interval = autoOkMs };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();

                    // Find the message box window
                    var hWnd = FindWindow("#32770", title); // #32770 = dialog class
                    if (hWnd != IntPtr.Zero)
                    {
                        var btn = FindWindowEx(hWnd, IntPtr.Zero, "Button", "OK");
                        if (btn != IntPtr.Zero)
                        {
                            SendMessage(btn, BM_CLICK, IntPtr.Zero, IntPtr.Zero); // auto-press OK
                        }
                    }
                };
                timer.Start();
            }

            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);

            tcs.TrySetResult(true); // complete when closed manually or auto-OK pressed
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();

        await tcs.Task;
    }
}