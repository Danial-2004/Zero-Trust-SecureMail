using Microsoft.Office.Tools.Ribbon;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using SecureMailAddin.MailHandler;

namespace SecureMailAddin
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors inspectors;
        private LoginForm _loginFormInstance;
        private bool _isLoginShown = false;
        private bool _isShuttingDown = false;
        private MailLifecycleManager _mailLifecycleManager;

        private const string CompanyDomain = "cuiatd.edu.pk";

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Logger.Initialize();
            Logger.Log("SecureMailAddin startup initiated.");

            try
            {
                var explorer = this.Application.ActiveExplorer();
                if (explorer == null)
                {
                    Logger.Log("ActiveExplorer is null. Continuing without hiding Outlook window.");
                }

                IntPtr handle = IntPtr.Zero;

                if (explorer != null)
                {
                    foreach (var proc in Process.GetProcessesByName("OUTLOOK"))
                    {
                        try
                        {
                            if (proc.MainWindowTitle == explorer.Caption)
                            {
                                handle = proc.MainWindowHandle;
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Log("Error while locating Outlook window handle: " + ex.Message);
                        }
                    }
                }

                if (handle != IntPtr.Zero)
                {
                    ShowWindow(handle, SW_HIDE);
                    Logger.Log("Outlook window hidden until login completes.");
                }

                ShowMandatoryLogin(handle);

                if (handle != IntPtr.Zero)
                {
                    ShowWindow(handle, SW_SHOW);
                    Logger.Log("Outlook window restored after successful login.");
                }

                inspectors = this.Application.Inspectors;

                // Only MailLifecycleManager handles send-time mail processing
                _mailLifecycleManager = new MailLifecycleManager(this.Application, CompanyDomain);
                _mailLifecycleManager.Start();

                Logger.Log("MailLifecycleManager started successfully for domain: " + CompanyDomain);
            }
            catch (Exception ex)
            {
                Logger.Log("Error during ThisAddIn_Startup: " + ex.ToString());
                MessageBox.Show(
                    "SecureMail Add-in failed to initialize.\n\n" + ex.Message,
                    "SecureMail Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void ShowMandatoryLogin(IntPtr outlookHandle)
        {
            if (_isShuttingDown)
                return;

            try
            {
                if (_loginFormInstance == null || _loginFormInstance.IsDisposed)
                    _loginFormInstance = new LoginForm();

                _isLoginShown = true;

                DialogResult loginResult;

                if (outlookHandle != IntPtr.Zero)
                {
                    loginResult = _loginFormInstance.ShowDialog(new WindowWrapper(outlookHandle));
                }
                else
                {
                    loginResult = _loginFormInstance.ShowDialog();
                }

                _isLoginShown = false;

                if (loginResult != DialogResult.OK)
                {
                    Logger.Log("Login failed or dialog closed incorrectly. Shutting down Outlook.");
                    _isShuttingDown = true;

                    foreach (var proc in Process.GetProcessesByName("OUTLOOK"))
                    {
                        try
                        {
                            proc.Kill();
                        }
                        catch (Exception ex)
                        {
                            Logger.Log("Failed to kill Outlook process: " + ex.Message);
                        }
                    }

                    Environment.Exit(0);
                }

                if (_loginFormInstance.AuthResult == null ||
                    string.IsNullOrWhiteSpace(_loginFormInstance.AuthResult.AccessToken))
                {
                    Logger.Log("Login completed but access token is missing. Shutting down Outlook.");
                    _isShuttingDown = true;

                    foreach (var proc in Process.GetProcessesByName("OUTLOOK"))
                    {
                        try
                        {
                            proc.Kill();
                        }
                        catch (Exception ex)
                        {
                            Logger.Log("Failed to kill Outlook process: " + ex.Message);
                        }
                    }

                    Environment.Exit(0);
                }

                Logger.Log("Mandatory login completed successfully.");
            }
            catch (Exception ex)
            {
                Logger.Log("Error during ShowMandatoryLogin: " + ex.ToString());
                MessageBox.Show(
                    "Login process failed.\n\n" + ex.Message,
                    "SecureMail Login Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                Environment.Exit(0);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _isShuttingDown = true;
            Logger.Log("Shutdown initiated.");

            try
            {
                if (_mailLifecycleManager != null)
                {
                    _mailLifecycleManager.Stop();
                    _mailLifecycleManager = null;
                    Logger.Log("MailLifecycleManager stopped successfully.");
                }

                if (_loginFormInstance != null && !_loginFormInstance.IsDisposed)
                {
                    LoginForm.AllowCloseOnShutdown = true;
                    _loginFormInstance.Close();
                    _loginFormInstance.Dispose();
                    _loginFormInstance = null;
                }

                if (inspectors != null)
                {
                    Marshal.ReleaseComObject(inspectors);
                    inspectors = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                Logger.Log("SecureMailAddin shutdown completed.");
            }
            catch (Exception ex)
            {
                Logger.Log("Error during ThisAddIn_Shutdown: " + ex.ToString());
            }
        }

        public class WindowWrapper : IWin32Window
        {
            private readonly IntPtr _hwnd;

            public WindowWrapper(IntPtr handle)
            {
                _hwnd = handle;
            }

            public IntPtr Handle => _hwnd;
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}