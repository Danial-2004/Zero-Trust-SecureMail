using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SecureMailAddin.MailHandler
{
    public class MailLifecycleManager
    {
        private readonly Outlook.Application _application;
        private readonly MailHandlerModule _mailHandlerModule;

        public MailLifecycleManager(Outlook.Application application, string companyDomain)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _mailHandlerModule = new MailHandlerModule(companyDomain);
        }

        public void Start()
        {
            _application.ItemSend += OnItemSend;
        }

        public void Stop()
        {
            _application.ItemSend -= OnItemSend;
        }

        private void OnItemSend(object item, ref bool cancel)
        {
            try
            {
                if (item is Outlook.MailItem mail)
                {
                    bool allowSend = _mailHandlerModule.HandleBeforeSend(mail);
                    if (!allowSend)
                        cancel = true;
                }
            }
            catch (Exception ex)
            {
                cancel = true;

                MessageBox.Show(
                    "An error occurred while processing the email before sending.\n\n" + ex.Message,
                    "Secure Mail",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}