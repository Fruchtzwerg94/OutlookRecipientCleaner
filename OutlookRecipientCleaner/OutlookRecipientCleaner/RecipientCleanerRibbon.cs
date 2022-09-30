using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

using OutlookRecipientCleaner.Forms;

namespace OutlookRecipientCleaner
{
    public partial class RecipientCleanerRibbon
    {
        private void RecipientCleanerRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Debug.WriteLine("Loading ribbon", nameof(RecipientCleanerRibbon));
        }

        private void SplitButton_Clean_Click(object sender, RibbonControlEventArgs e)
        {
            Debug.WriteLine("Start cleaning recipients", nameof(RecipientCleanerRibbon));

            Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector?.CurrentItem is MailItem mail)
            {
                mail.Recipients.ResolveAll();

                //Order recipents by OlMailRecipientType, which allows to remove by priority: To --> CC --> BCC
                IEnumerable<Recipient> recipents = mail.Recipients.Cast<Recipient>().OrderBy(r => r.Type);

                List<Recipient> recipientsToRemove = new List<Recipient>();
                IEnumerable<IGrouping<string, Recipient>> groupedRecipients = recipents.GroupBy(g => g.Address);
                IEnumerable<IGrouping<string, Recipient>> nonUniqueRecipients = groupedRecipients.Where(gr => gr.Count() > 1);
                foreach (IGrouping<string, Recipient> nonUniqueRecipient in nonUniqueRecipients)
                {
                    //Remove all non unique recipients, keep the first one
                    recipientsToRemove.AddRange(nonUniqueRecipient.Skip(1));
                }

                //Remove non unique recipients
                Debug.WriteLine($"Removing {recipientsToRemove.Count} recipients", nameof(RecipientCleanerRibbon));
                foreach (Recipient recipientToRemove in recipientsToRemove)
                {
                    Debug.WriteLine($"Removing {recipientToRemove.Name}: {recipientToRemove.Address}", nameof(RecipientCleanerRibbon));
                    mail.Recipients.Remove(recipientToRemove.Index);
                }
            }
            else
            {
                Debug.WriteLine("Failed to get mail item", nameof(RecipientCleanerRibbon));
                MessageBox.Show("Failed to access a mail", "Failed to clean participants", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Debug.WriteLine("Finished cleaning recipients", nameof(RecipientCleanerRibbon));
        }

        private void Button_About_Click(object sender, RibbonControlEventArgs e)
        {
            using (AboutWindow aboutWindow = new AboutWindow())
            {
                aboutWindow.ShowDialog();
            }
        }
    }
}
