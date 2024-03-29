﻿using System;
using System.Windows.Forms;

namespace OutlookRecipientCleaner.Forms
{
    public static class FormsExtensions
    {
        public static void InvokeIfRequired(this Form form, Action action)
        {
            if (form.InvokeRequired)
            {
                form.Invoke(action);
            }
            else
            {
                action();
            }
        }
    }
}
