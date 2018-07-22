using System;
using outlook = Microsoft.Office.Interop.Outlook;
using System.Net.Mail;
using System.Collections.Generic;

namespace PDAImport
{
    public class email
    {

        // Sendemail
        public static void Sendemail(string strSubject, string message, string fullFileName, string txtEmail, string txtEmailcc, string sSalesPerson)
            {

            if (Program.sPrd != "YES")
            {
                strSubject = strSubject + " - TEST";
            }

            if (Program.sourceSystem == "JFC")
                //Call emailsmtp(strSubject, Message, sFilename)
                emailsmtp(strSubject, message, fullFileName, txtEmail, txtEmailcc, sSalesPerson);
            else
                SendOutlookMail(strSubject, message, fullFileName, txtEmail, txtEmailcc, sSalesPerson);
            }

            static void emailsmtp(string strSubject, string message, string txtFilePath, string txtEmail, string txtEmailcc, string sSalesPerson)
            {
                const char LF = '\n';
                const char CR = '\r';

                SmtpClient Smtp_Server = new SmtpClient();
                MailMessage emailmessage = new MailMessage();

                Smtp_Server.UseDefaultCredentials = true;
                Smtp_Server.Port = 25;
                Smtp_Server.EnableSsl = true;
                Smtp_Server.Host = "smtp.kmsnet.com";

            if (!string.IsNullOrWhiteSpace(txtEmail))
            {
                string[] arrAddTos = txtEmail.Split(new char[] { ';', ',' });
                foreach (string strAddr in arrAddTos)
                {
                    if (!string.IsNullOrWhiteSpace(strAddr) &&
                        strAddr.IndexOf('@') != -1)
                    {
                        emailmessage.To.Add(strAddr.Trim());
                    }
                    else
                        throw new Exception("Bad to-address: " + strAddr);
                }
            }
            else
                throw new Exception("Must specify to-address");

            //Parse 'sCc'
            if (!string.IsNullOrWhiteSpace(txtEmailcc))
            {
                string[] arrAddTos = txtEmailcc.Split(new char[] { ';', ',' });
                foreach (string strAddr in arrAddTos)
                {
                    if (!string.IsNullOrWhiteSpace(strAddr) &&
                        strAddr.IndexOf('@') != -1)
                    {
                        emailmessage.CC.Add(strAddr.Trim());
                    }
                    else
                        throw new Exception("Bad CC-address: " + strAddr);
                }
            }
            else
                throw new Exception("Must specify to-address");

            emailmessage.From = new MailAddress("jetssoadmin@jfc.ca");
            emailmessage.Subject = strSubject;
            emailmessage.IsBodyHtml = false;
            emailmessage.Body = message;

            try
                {
                    Smtp_Server.Send(emailmessage);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception caught in emailsmtp(): {0}",
                                ex.ToString());
                }

                Smtp_Server.Dispose();
            }

            static void SendOutlookMail(string strSubject, string message, string fullFileName, string txtEmail, string txtEmailcc, string sSalesPerson)
            {
                outlook.MailItem OutlookMessage;
                outlook.Application AppOutlook = new outlook.Application();

            try
            {
                OutlookMessage = AppOutlook.CreateItem(outlook.OlItemType.olMailItem);
                outlook.Recipients Recipients = OutlookMessage.Recipients;
                //OutlookMessage.From = "jetssoadmin@jfc.ca";
                OutlookMessage.Subject = strSubject;
                OutlookMessage.Body = message;
                OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML;
                OutlookMessage.Attachments.Add(fullFileName, outlook.OlAttachmentType.olByValue);

                if (!string.IsNullOrWhiteSpace(txtEmail))
                {
                    string[] arrAddTos = txtEmail.Split(new char[] { ';', ',' });
                    foreach (string strAddr in arrAddTos)
                    {
                        if (!string.IsNullOrWhiteSpace(strAddr) &&
                            strAddr.IndexOf('@') != -1)
                        {
                            Recipients.Add(strAddr.Trim());
                        }
                        else
                            throw new Exception("Bad to-address: " + txtEmail);
                    }
                }
                else
                    throw new Exception("Must specify to-address");

                //Parse 'sCc'
                if (!string.IsNullOrWhiteSpace(txtEmailcc))
                {
                    string[] arrAddTos = txtEmailcc.Split(new char[] { ';', ',' });
                    foreach (string strAddr in arrAddTos)
                    {
                        if (!string.IsNullOrWhiteSpace(strAddr) &&
                            strAddr.IndexOf('@') != -1)
                        {
                            Recipients.Add(strAddr.Trim());
                        }
                        else
                            throw new Exception("Bad CC-address: " + txtEmailcc);
                    }
                }
                else
                    throw new Exception("Must specify to-address");

                OutlookMessage.Send();
                }
                catch (Exception ex)
                {
                }

                // MessageBox.Show("Mail could not be sent") 'if you dont want this message, simply delete this line  
                finally
                {
                    OutlookMessage = null/* TODO Change to default(_) if this is not a reference type */;
                    AppOutlook = null/* TODO Change to default(_) if this is not a reference type */;
                }

                return;
            }
        }

}

