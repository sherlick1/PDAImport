using System;
using outlook = Microsoft.Office.Interop.Outlook;
using System.Net.Mail;
using System.Collections.Generic;

namespace PDAImport
{
    public class email
    {
        // Sendemail
        public static void Sendemail(string strSubject, string message, string fullFileName, string txtEmail, string txtEmailcc, string sSalesPerson, string[][] arr)
        {

            int iSalesRep = -1;

            if (!string.IsNullOrWhiteSpace(sSalesPerson) && Program.emailSalesRep)
            {
                iSalesRep = Utilities.IsInArray(sSalesPerson, arr);

                if (txtEmail != null)
                    txtEmail = txtEmail + ";" + arr[iSalesRep][1];
                else
                    txtEmail = arr[iSalesRep][1];
            }

            if (Program.sPrd.ToUpper() != "PRD")
            {
                strSubject = strSubject + " - TEST";
            }

            if (Program.sourceSystem == "JFC")
                //Call emailsmtp(strSubject, Message, sFilename)
                emailsmtp(strSubject, message, fullFileName, txtEmail, txtEmailcc);
            else
                SendOutlookMail(strSubject, message, fullFileName, txtEmail, txtEmailcc);
            }

            static void emailsmtp(string strSubject, string message, string txtFilePath, string txtEmail, string txtEmailcc)
            {

                SmtpClient Smtp_Server = new SmtpClient();
                MailMessage emailmessage = new MailMessage();
                LogWriter log = new LogWriter();

                Smtp_Server.UseDefaultCredentials = true;
                Smtp_Server.Port = 25;
                Smtp_Server.EnableSsl = false;
                Smtp_Server.Host = "smtp.kmsnet.com";

            if (!string.IsNullOrWhiteSpace(txtEmail) && Program.emailSalesRep)
            {
                string[] arrAddTos = txtEmail.Split(new char[] { ';', ',' });
                foreach (string strAddr in arrAddTos)
                {
                    if (!string.IsNullOrWhiteSpace(strAddr) && strAddr.IndexOf('@') != -1)
                    {
                        emailmessage.To.Add(strAddr.Trim());
                    }
                    else
                    {
                        if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                            log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + txtFilePath + "," + "Bad to-address: " + strAddr);

                        if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                            log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + txtFilePath + "," + "Bad to-address: " + strAddr);
                    }
                    //throw new Exception("Bad to-address: " + strAddr);
                }
            }
            else
                if (string.IsNullOrWhiteSpace(txtEmail))
                {
                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                        log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + txtFilePath + "," + "Must specify to-address.");

                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + txtFilePath + "," + "Must specify to-address.");
                }
                //throw new Exception("Must specify to-address");

            //Parse 'cc'
            if (!string.IsNullOrWhiteSpace(txtEmailcc))
            {
                string[] arrAddTos = txtEmailcc.Split(new char[] { ';', ',' });
                foreach (string strAddr in arrAddTos)
                {
                    if (!string.IsNullOrWhiteSpace(strAddr) && strAddr.IndexOf('@') != -1)
                    {
                        emailmessage.CC.Add(strAddr.Trim());
                    }
                    else
                    {
                        if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                            log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + txtFilePath + "," + "Bad cc-address: " + strAddr);

                        if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                            log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + txtFilePath + "," + "Bad cc-address: " + strAddr);
                    }
                    //throw new Exception("Bad CC-address: " + strAddr);
                }
            }
            else
            {
                if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                    log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + txtFilePath + "," + "Must specify cc-address.");

                if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                    log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + txtFilePath + "," + "Must specify cc-address.");
            }
            //throw new Exception("Must specify cc-address");

            emailmessage.From = new MailAddress("jetssoadmin@jfc.ca");
            emailmessage.Subject = strSubject;
            emailmessage.IsBodyHtml = false;
            emailmessage.Body = message;

            //emailmessage.Body = htmlMessageBody;
            //emailmessage.BodyEncoding = Encoding.UTF8;
            //emailmessage.IsBodyHtml = true;

            try
                {
                    Smtp_Server.Send(emailmessage);
                }
                catch (Exception ex)
                {
                    if (Program.sMode != "Silent")
                        Console.WriteLine("Exception caught in emailsmtp(): {0}",ex.Message.ToString());

                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                        log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + txtFilePath + "," + "Exception caught in emailsmtp(): " + ex.ToString());

                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + txtFilePath + "," + "Exception caught in emailsmtp(): " + ex.ToString());
            }

            Smtp_Server.Dispose();
            }

            static void SendOutlookMail(string strSubject, string message, string fullFileName, string txtEmail, string txtEmailcc)
            {
                outlook.MailItem OutlookMessage;
                outlook.Application AppOutlook = new outlook.Application();
                LogWriter log = new LogWriter();

            try
            {
                OutlookMessage = AppOutlook.CreateItem(outlook.OlItemType.olMailItem);
                outlook.Recipients Recipients = OutlookMessage.Recipients;
                //OutlookMessage.From = "jetssoadmin@jfc.ca";
                OutlookMessage.Subject = strSubject;
                OutlookMessage.Body = message;
                OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML;
                if (!string.IsNullOrWhiteSpace(fullFileName))
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
                        {
                            if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "Bad to-address: " + strAddr);

                            if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "Bad to-address: " + strAddr);
                            // throw new Exception("Bad to-address: " + txtEmail);
                        }
                    }
                }
                else
                {
                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                        log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "Must specify to-address.");

                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "Must specify to-address.");
                }
                //throw new Exception("Must specify to-address");

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
                        {
                            if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "Bad cc-address: " + strAddr);

                            if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "Bad cc-address: " + strAddr);
                        }
                        //throw new Exception("Bad CC-address: " + txtEmailcc);
                    }
                }
                else
                {
                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                        log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "Must specify cc-address.");

                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "Must specify cc-address.");
                }
                //throw new Exception("Must specify to-address");

                OutlookMessage.Send();
                }

                catch (Exception ex)
                {
                }

                // MessageBox.Show("Mail could not be sent") 'if you dont want this message, simply delete this line  
                finally
                {
                    OutlookMessage = null   /* TODO Change to default(_) if this is not a reference type */;
                    AppOutlook = null       /* TODO Change to default(_) if this is not a reference type */;
                }

                return;
            }
        }
}

