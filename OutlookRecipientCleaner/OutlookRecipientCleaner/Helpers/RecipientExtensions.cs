using System.Diagnostics;

using Microsoft.Office.Interop.Outlook;

namespace OutlookRecipientCleaner.Helpers
{
    public static class RecipientExtensions
    {
        public static string GetSmtpAddress(this Recipient recipient)
        {
            //See: https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-e-mail-address-of-a-recipient
            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            PropertyAccessor pa = recipient.PropertyAccessor;
            try
            {
                return pa.GetProperty(PR_SMTP_ADDRESS).ToString();
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("Failed to get SMTP address via property accessor: " + ex, nameof(RecipientExtensions));
                return recipient.Address;
            }
        }
    }
}
