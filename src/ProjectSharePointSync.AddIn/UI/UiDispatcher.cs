using System;
using System.Windows.Forms;
using ProjectSharePointSync.Core.Diagnostics;

namespace ProjectSharePointSync.AddIn.UI
{
    /// <summary>
    /// Marshals callbacks from background threads to the MS Project UI thread.
    /// Must be created on the UI thread.
    /// </summary>
    internal sealed class UiDispatcher : IDisposable
    {
        private readonly Control _control;

        public UiDispatcher()
        {
            _control = new Control();
            _control.CreateControl();
            // Force handle creation so BeginInvoke works before the control is shown.
            var handle = _control.Handle;
        }

        public void Post(Action action)
        {
            if (_control.IsDisposed) return;
            try
            {
                _control.BeginInvoke((MethodInvoker)delegate
                {
                    try
                    {
                        action();
                    }
                    catch (Exception ex)
                    {
                        Log.Error("UI callback failed.", ex);
                    }
                });
            }
            catch (InvalidOperationException)
            {
                // Handle destroyed during shutdown.
            }
        }

        public void Dispose()
        {
            _control.Dispose();
        }
    }

    /// <summary>Wraps the MS Project main window handle so it can own dialogs.</summary>
    internal sealed class Win32Window : IWin32Window
    {
        public Win32Window(IntPtr handle)
        {
            Handle = handle;
        }

        public IntPtr Handle { get; private set; }

        public static IWin32Window ProjectMainWindow
        {
            get { return new Win32Window(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle); }
        }
    }
}
