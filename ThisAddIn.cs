using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace FilterAddin
{
    public partial class ThisAddIn
    {
        private Outlook.MailItem previousMailItem;
        private Outlook.Explorer currentExplorer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Helper.LoadSettings();
            //this.Application.NewMailEx += Application_NewMailEx;
            currentExplorer = this.Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook
                .ExplorerEvents_10_SelectionChangeEventHandler
            (CurrentExplorer_Event);
        }
    
        private void CurrentExplorer_Event()
        {
            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];
                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem =
                            (selObject as Outlook.MailItem);
                        
                        if (mailItem != null)
                        {
                            string SenderAddress = GetSenderSMTPAddress(mailItem);
                            
                            if (previousMailItem != null)
                            {
                                BlackListHandler.RestoreAttachments(previousMailItem);
                            }
                            if (IsBlacklistedSender(SenderAddress))
                            {
                                        
                                BlackListHandler.RemoveAttachmentscurrent(mailItem);
                                mailItem.HTMLBody = $"{CONFIGS.Settings.blacklistWarning}";
                                previousMailItem = mailItem;

                            }
                            else if (!IsWhitelistedSender(SenderAddress))
                            {
                                if (IsBlacklistedSender(SenderAddress, true))
                                {

                                    BlackListHandler.RemoveAttachmentscurrent(mailItem);
                                    mailItem.HTMLBody = $"{CONFIGS.Settings.blacklistWarning}";
                                    previousMailItem = mailItem;
                                    
                                }
                                else if (!IsWhitelistedSender(SenderAddress, true))
                                {
                                    BlackListHandler.RemoveAttachmentscurrent(mailItem);
                                    mailItem.HTMLBody = $"{CONFIGS.Settings.unkownSenderWarning}";
                                    previousMailItem = mailItem;
                                    
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Application_NewMailEx(string EntryIDCollection)
        {
            Outlook.NameSpace outlookNamespace = Application.GetNamespace("MAPI");
            Outlook.MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            foreach (string entryID in EntryIDCollection.Split(','))
            {
                if (!string.IsNullOrEmpty(entryID))
                {
                    Outlook.MailItem newMail = outlookNamespace.GetItemFromID(entryID) as Outlook.MailItem;

                    if (newMail != null)
                    {
                        var mailItem = newMail as Outlook.MailItem;
                        if (IsBlacklistedSender(mailItem.Sender.Address))
                        {

                            BlackListHandler.RemoveAttachments(mailItem);
                            BlackListHandler.RemoveHyperlinksFromHTMLBody(mailItem);

                            mailItem.HTMLBody = $"{CONFIGS.Settings.blacklistWarning}" + newMail.HTMLBody;
                        }
                        else if (!IsWhitelistedSender(mailItem.Sender.Address))
                        {
                            if (IsBlacklistedSender(mailItem.Sender.Address, true))
                            {

                                BlackListHandler.RemoveAttachments(mailItem);
                                BlackListHandler.RemoveHyperlinksFromHTMLBody(mailItem);

                                mailItem.HTMLBody = $"{CONFIGS.Settings.blacklistWarning}" + newMail.HTMLBody;
                            }
                            else if (!IsWhitelistedSender(mailItem.Sender.Address, true))
                            {
                                BlackListHandler.RemoveAttachments(mailItem);
                                BlackListHandler.RemoveHyperlinksFromHTMLBody(mailItem);
                                mailItem.HTMLBody = $"{CONFIGS.Settings.unkownSenderWarning}" + newMail.HTMLBody;
                            }
                        }
                    }
                }
            }
        }
        private bool IsBlacklistedSender(string senderAddress, bool domainMatch = false)
        {
            if (domainMatch)
            {
                string whString = string.Join("; ", CONFIGS.Settings.blacklist);
                MessageBox.Show($"The blacklist email is: {whString}.");
                foreach (var item in CONFIGS.Settings.blacklist)
                {
                    //if (senderAddress.ToLower().Contains(item.ToLower()))
                    //    return true;
                    MessageBox.Show($"The sender's email is: {senderAddress}. Compare it to the blacklisted email: {item}.");
                    if (senderAddress.ToLower() == item.ToLower())
                        return true;
                }
            }
            else
            {
                foreach (var item in CONFIGS.Settings.blacklist)
                {
                    if (senderAddress.ToLower() == item.ToLower())
                        return true;
                }
            }

            return false;
        }

        private bool IsWhitelistedSender(string senderAddress, bool domainMatch = false)
        {
            if (domainMatch)
            {
                string whString = string.Join(",", CONFIGS.Settings.whitelist);
                MessageBox.Show($"The whitelist email is: {whString}.");
                foreach (var item in CONFIGS.Settings.whitelist)
                {
                    MessageBox.Show($"The sender's email is: {senderAddress}. Compare it to the whitelist email: {item}.");
                    if (senderAddress.ToLower().Contains(item.ToLower()))
                        return true;
                }
            }
            else
            {
                foreach (var item in CONFIGS.Settings.whitelist)
                {
                    if (senderAddress.ToLower() == item.ToLower())
                        return true;
                }
            }
            return false;
        }

        private string GetSenderSMTPAddress(Outlook.MailItem mail)
        {
            string PR_SMTP_ADDRESS =
                @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            if (mail == null)
            {
                throw new ArgumentNullException();
            }
            if (mail.SenderEmailType == "EX")
            {
                Outlook.AddressEntry sender = mail.Sender;
                if (sender != null)
                {
                    //Now we have an AddressEntry representing the Sender
                    if (sender.AddressEntryUserType ==
                        Outlook.OlAddressEntryUserType.
                        olExchangeUserAddressEntry
                        || sender.AddressEntryUserType ==
                        Outlook.OlAddressEntryUserType.
                        olExchangeRemoteUserAddressEntry)
                    {
                        //Use the ExchangeUser object PrimarySMTPAddress
                        Outlook.ExchangeUser exchUser =
                            sender.GetExchangeUser();
                        if (exchUser != null)
                        {
                            return exchUser.PrimarySmtpAddress;
                        }
                        else
                        {
                            return null;
                        }
                    }
                    else
                    {
                        return sender.PropertyAccessor.GetProperty(
                            PR_SMTP_ADDRESS) as string;
                    }
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return mail.SenderEmailAddress;
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
