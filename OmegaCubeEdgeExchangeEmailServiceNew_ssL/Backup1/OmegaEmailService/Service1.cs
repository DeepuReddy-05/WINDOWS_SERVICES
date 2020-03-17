using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.Data.OleDb;
using System.Net;
using System.Net.Mail;
using System.IO;
using System.Configuration;
using System.Configuration.Internal;
using Microsoft.ServiceModel.Channels.Mail.ExchangeWebService.Exchange2007;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Threading;


namespace OmegaEmailService
{
    public partial class OmegaEmailServiceForTaskSystem : ServiceBase
    {
        public string strServerName = string.Empty;
        public string strUserName = string.Empty;
        public string strPassWord = string.Empty;
        public string strFromName = string.Empty;
        public string strSmtpServer = string.Empty;
        public string strSmtpPort = string.Empty;
        public string strSmtpUid = string.Empty;
        public string strSmtpPwd = string.Empty;
        public string strFromEmailID = string.Empty;
        public string strDomain = string.Empty;
        public string strExorSm = string.Empty;
        public string strOracleRun = string.Empty;
        public string strOracleReportPath = string.Empty;
        public string strEnableSSL = string.Empty;

        public string fName = string.Empty;

        protected static string AppPath = System.AppDomain.CurrentDomain.BaseDirectory;
        protected static string strLogFilePath = AppPath + "\\" + "OmegacubeEmailLog.txt";
        private static StreamWriter sw = null;

        public static int LoopCount;

        public OmegaEmailServiceForTaskSystem()
        {
            InitializeComponent();

            // Remove Event Source if already there    
            if (EventLog.SourceExists("OmegacubeEmailSource"))
            {
                EventLog.DeleteEventSource("OmegacubeEmailSource");
            }

            // Initialize eventLogSimple 
            if (!System.Diagnostics.EventLog.SourceExists("OmegacubeEmailSource"))
            {
                System.Diagnostics.EventLog.CreateEventSource("OmegacubeEmailSource", "OmegacubeEmailLog");
            }
            // SimpleSource will appear as a column in eventviewer.
            eventLogFileOEE.Source = "OmegacubeEmailSource";
            // SimpleLog will appear as an event folder.
            eventLogFileOEE.Log = "OmegacubeEmailLog";

            strServerName = ConfigurationManager.AppSettings["Server"];
            strUserName = ConfigurationManager.AppSettings["User"];
            strPassWord = ConfigurationManager.AppSettings["Pwd"];

            strFromName = ConfigurationManager.AppSettings["frmName"];
            strFromEmailID = ConfigurationManager.AppSettings["frmEmailID"];

            strSmtpServer = ConfigurationManager.AppSettings["smtpServer"];
            strSmtpPort = ConfigurationManager.AppSettings["smtpPort"];

            strSmtpUid = ConfigurationManager.AppSettings["smtpUid"];
            strSmtpPwd = ConfigurationManager.AppSettings["smtpPwd"];

            strDomain = ConfigurationManager.AppSettings["Domain"];

            strExorSm = ConfigurationManager.AppSettings["ExhangeOrSMTP"];

            strOracleRun = ConfigurationManager.AppSettings["oracleRun"];
            strOracleReportPath = ConfigurationManager.AppSettings["oracleReportsPath"];

            strEnableSSL = ConfigurationManager.AppSettings["EnableSSL"];
        }

        OleDbDataAdapter da;
        OleDbConnection dbcon;

        string strSno = string.Empty;
        string strClient = string.Empty;
        string strToAdd = string.Empty;
        string strCcAdd = string.Empty;
        string strBccAdd = string.Empty;
        string strFromAdd = string.Empty;
        string strMailSub = string.Empty;
        string strMailBody = string.Empty;
        string strAttachments = string.Empty;
        string strNewsLetter = string.Empty;

        public string strReportName = string.Empty;
        public string strViewName = string.Empty;
        public string strCondParam = string.Empty;
        public string strCondValue = string.Empty;
        public string strReportFileName = string.Empty;

        string strUpdate = string.Empty;

        System.Timers.Timer testTimer = new System.Timers.Timer();

        OleDbCommand cmd = new OleDbCommand();

        protected override void OnStart(string[] args)
        {
            try
            {
                LogExceptionsN(strLogFilePath, null, "s Start");
                //SendBulkEmail();
                testTimer.Interval = 9999;
                testTimer.Elapsed += new System.Timers.ElapsedEventHandler(testTimer_Elapsed);
                testTimer.Enabled = true;
                LogExceptionsN(strLogFilePath, null, "s end Start");
            }
            catch (Exception ex)
            {
                eventLogFileOEE.WriteEntry(ex.Message.ToString());
                LogExceptions(strLogFilePath, ex);
            }
        }

        private void testTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
               
                LoopCount = LoopCount + 1;
                if (LoopCount == 2)
                {
                    LogExceptionsN(strLogFilePath, null, "in side timer");
                    LoopCount = 0;
                    testTimer.Enabled = false;
                    SendBulkEmail();
                    testTimer.Enabled = true;
                    LogExceptionsN(strLogFilePath, null, "out side timer");
                }
                
            }
            catch (Exception ex)
            {
                eventLogFileOEE.WriteEntry(ex.Message.ToString());
                LogExceptions(strLogFilePath, ex);
            }
        }

        protected override void OnStop()
        {
            Thread.Sleep(1000);
            testTimer.Enabled = false;            
        }

        protected override void  OnPause()
        {
            Thread.Sleep(1000);
            testTimer.Enabled = false;            
        }

        public void SendBulkEmail()
        {
            try
            {
                DataSet dsNew = new DataSet();
                string selectQuery = "SELECT * FROM A_GENERIC_EMAIL WHERE SEND_CONFIRMATION='N'";
                dbcon = new OleDbConnection("Provider=MSDAORA;Data Source=" + strServerName + ";User Id=" + strUserName + ";Password=" + strPassWord + ";");
                da = new OleDbDataAdapter(selectQuery, dbcon);
                try
                {
                     
                    da.Fill(dsNew, "Email");
                }
                catch (Exception ex1)
                {
                    LogExceptionsN(strLogFilePath, ex1,"");
                    LogExceptionsN(strLogFilePath, null , ex1.InnerException.ToString());
                }
                for (int i = 0; i < dsNew.Tables[0].Rows.Count; i++)
                {
                    Thread.Sleep(10000);
                    bool chkEmail = false;

                    strSno = dsNew.Tables[0].Rows[i]["SNO"].ToString();
                    strClient = dsNew.Tables[0].Rows[i]["CLIENT"].ToString();

                    strToAdd = dsNew.Tables[0].Rows[i]["TO_ADDRESS"].ToString();
                    strCcAdd = dsNew.Tables[0].Rows[i]["CC_ADDRESS"].ToString();
                    strBccAdd = dsNew.Tables[0].Rows[i]["BCC_ADDRESS"].ToString();
                    strFromAdd = dsNew.Tables[0].Rows[i]["FROM_ID"].ToString();
                    strMailSub = dsNew.Tables[0].Rows[i]["MAIL_SUBJECT"].ToString();
                    strMailBody = dsNew.Tables[0].Rows[i]["MAIL_BODY"].ToString();
                    strAttachments = dsNew.Tables[0].Rows[i]["ATTACHMENT"].ToString();
                    strNewsLetter = dsNew.Tables[0].Rows[i]["NEWS_LETTER"].ToString();

                    strReportName = dsNew.Tables[0].Rows[i]["REPORT_NAME"].ToString();
                    strCondParam = dsNew.Tables[0].Rows[i]["CONDITION_PARAMETER"].ToString();
                    strCondValue = dsNew.Tables[0].Rows[i]["CONDITION_VALUE"].ToString();
                    strReportFileName = dsNew.Tables[0].Rows[i]["REPORT_FILE_NAME"].ToString();
                    strViewName = dsNew.Tables[0].Rows[i]["VIEW_NAME"].ToString();



                    //strFromAdd = strFromEmailID;

                    //chkEmail = SendEmail(dsNew.Tables[0].Rows[i]["TO_ADDRESS"].ToString(),dsNew.Tables[0].Rows[i]["CC_ADDRESS"].ToString(),dsNew.Tables[0].Rows[i]["BCC_ADDRESS"].ToString(),
                    if (!string.IsNullOrEmpty(strToAdd))
                    {
                        if (strExorSm == "E")
                        {
                            // Using Exchange
                            chkEmail = SendEmailWithExchange(strToAdd, strCcAdd, strBccAdd, strFromAdd, strMailSub, strMailBody, strAttachments);
                        }
                        else
                        {
                            // Using  SMTP
                            chkEmail = SendEmailWithSMTP(strToAdd, strCcAdd, strBccAdd, strFromAdd, strMailSub, strMailBody, strAttachments);
                        }

                        if (chkEmail == true)
                        {
                            strUpdate = "UPDATE A_GENERIC_EMAIL  SET SEND_CONFIRMATION='Y' WHERE  SNO =" + strSno;
                            eventLogFileOEE.Clear();
                            eventLogFileOEE.WriteEntry("Mail has been Send to:" + strSno);
                        }
                        else
                        {
                            strUpdate = "UPDATE A_GENERIC_EMAIL  SET SEND_CONFIRMATION='R' WHERE  SNO =" + strSno;
                        }
                    }
                    else
                    {
                        strUpdate = "UPDATE A_GENERIC_EMAIL SET SEND_CONFIRMATION='R',CHANGED_BY='ADMN', CHANGED_DATE=SYSDATE  WHERE SNO=" + strSno;
                    }
                    cmd.Connection = dbcon;
                    cmd.CommandText = strUpdate;
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection.Open();
                    int count = cmd.ExecuteNonQuery();
                    cmd.Connection.Close();
                }
            }
            catch (Exception ex)
            {
                eventLogFileOEE.WriteEntry(ex.Message.ToString());
                LogExceptions(strLogFilePath, ex);
            }
        }     


        public static void LogExceptions(string filePath, Exception ex)
        {
            if (false == File.Exists(filePath))
            {
                FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                fs.Close();
            }
            WriteExceptionLog(filePath, ex);
        }


        private static void WriteExceptionLog(string strPathName, Exception objException)
        {
            sw = new StreamWriter(strPathName, true);
            sw.WriteLine("Source		: " + objException.Source.ToString().Trim());
            sw.WriteLine("Method		: " + objException.TargetSite.Name.ToString());
            sw.WriteLine("Date		: " + DateTime.Now.ToLongTimeString());
            sw.WriteLine("Time		: " + DateTime.Now.ToShortDateString());
            sw.WriteLine("Error		: " + objException.Message.ToString().Trim());
            sw.WriteLine("Stack Trace	: " + objException.StackTrace.ToString().Trim());
            sw.WriteLine("^^-------------------------------------------------------------------^^");
            sw.Flush();
            sw.Close();
        }



        public static void LogExceptionsN(string filePath, Exception ex, string msg1)
        {
            if (false == File.Exists(filePath))
            {
                FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                fs.Close();
            }
            WriteExceptionLogN(filePath, ex, msg1);
        }


        private static void WriteExceptionLogN(string strPathName, Exception objException, string msg)
        {
            if (string.IsNullOrEmpty(msg))
            {
                sw = new StreamWriter(strPathName, true);
                sw.WriteLine("Source		: " + objException.Source.ToString().Trim());
                sw.WriteLine("Method		: " + objException.TargetSite.Name.ToString());
                sw.WriteLine("Date		: " + DateTime.Now.ToLongTimeString());
                sw.WriteLine("Time		: " + DateTime.Now.ToShortDateString());
                sw.WriteLine("Error		: " + objException.Message.ToString().Trim());
                sw.WriteLine("Stack Trace	: " + objException.StackTrace.ToString().Trim());
                sw.WriteLine("^^-------------------------------------------------------------------^^");
            }
            else
            {
                sw = new StreamWriter(strPathName, true);
                sw.WriteLine(msg);
            }
            sw.Flush();
            sw.Close();
        }


        public void Cert()
        {
            ServicePointManager.ServerCertificateValidationCallback =
                  delegate(Object obj, X509Certificate certificate,
                  X509Chain chain, SslPolicyErrors errors)
                  {
                      // Replace this line with code to validate server    
                      // certificate.   
                      return true;
                  };
        }

        public bool SendEmailWithExchange(string strTo, string strCc, string strBcc, string strFrom, string strSubject, string strBody, string strAttachmentPath)
        {
            char[] splitter = { ',' };

            ExchangeServiceBinding ewsServiceBinding = new ExchangeServiceBinding();
            ewsServiceBinding.Credentials = new NetworkCredential(strSmtpUid, strSmtpPwd, strDomain);
            ewsServiceBinding.Url = @"https://" + strSmtpServer +"/EWS/exchange.asmx";
            Cert();
            MessageType emMessage = new MessageType();

            List<EmailAddressType> recipients = new List<EmailAddressType>();

            // For To Field
            Array arrTo;
            arrTo = strTo.Split(splitter);

            foreach (string s in arrTo)
            {
                EmailAddressType recipient = new EmailAddressType();
                recipient.EmailAddress = s;
                recipients.Add(recipient);
            }

            emMessage.ToRecipients = recipients.ToArray();

            // For CC              
            if (!string.IsNullOrEmpty(strCc))
            {
                Array arrCc;
                arrCc = strCc.Split(splitter);

                recipients = new List<EmailAddressType>();
                foreach (string s in arrCc)
                {
                    EmailAddressType recipient = new EmailAddressType();
                    recipient.EmailAddress = s;
                    recipients.Add(recipient);
                }
                emMessage.CcRecipients = recipients.ToArray();
            }

            // For CC              
            if (!string.IsNullOrEmpty(strBcc))
            {
                Array arrBcc;
                arrBcc = strBcc.Split(splitter);

                recipients = new List<EmailAddressType>();
                foreach (string s in arrBcc)
                {
                    EmailAddressType recipient = new EmailAddressType();
                    recipient.EmailAddress = s;
                    recipients.Add(recipient);
                }
                emMessage.BccRecipients = recipients.ToArray();
            }

           
            emMessage.Subject = strSubject;
            emMessage.Body = new BodyType();
            emMessage.Body.BodyType1 = BodyTypeType.HTML;
            emMessage.Body.Value = strBody;
            
            //emMessage.Body.Value = strBody;
            //if (!string.IsNullOrEmpty(strAttachmentPath))
            //{
            //    emMessage.Body.Value = strBody + "<BR>" + ReadHtmlFile(strAttachmentPath) + "<BR>" + UnSubscribe(strTo);
            //}
            //else
            //{
            //    emMessage.Body.Value = strBody;
            //}

            emMessage.ItemClass = "IPM.Note";
            emMessage.Sensitivity = SensitivityChoicesType.Normal;

            try
            {
                ItemIdType iiCreateItemid = CreateDraftMessage(ewsServiceBinding, emMessage);

                // For Attachments
                if (!string.IsNullOrEmpty(strAttachmentPath))
                {
                    Array arrAttach;
                    arrAttach = strAttachmentPath.Split(splitter);

                    foreach (string s in arrAttach)
                    {
                        iiCreateItemid = CreateAttachment(ewsServiceBinding, s, iiCreateItemid);
                    }
                }

                // For NewsLetterAttachment
                if (strNewsLetter == "R")
                {
                    string FileName = string.Empty;
                    FileName = NewsLetterAttachment();

                    if (!string.IsNullOrEmpty(FileName))
                    {
                        if (File.Exists(FileName))
                        {
                            iiCreateItemid = CreateAttachment(ewsServiceBinding, FileName, iiCreateItemid);
                        }
                    }
                }

                SendMessage(ewsServiceBinding, iiCreateItemid);
                return true;
            }
            catch (Exception ex)
            {
                eventLogFileOEE.WriteEntry(ex.Message.ToString());
                LogExceptions(strLogFilePath, ex);
                return false;
            }
        }

        public bool SendEmailWithSMTP(string strTo, string strCc, string strBcc, string strFrom, string strSubject, string strBody, string strAttachmentPath)
        {
            try
            {
                char[] splitter = { ',' };
                MailMessage mm = new MailMessage();

                //mm.From = new MailAddress(strFrom);
                if (!string.IsNullOrEmpty(strFrom.Trim()))
                {
                    mm.ReplyTo = new MailAddress(strFrom);
                }
                else
                {
                    mm.ReplyTo = new MailAddress(strFromEmailID);
                }

                // For To Field
                Array arrTo;
                arrTo = strTo.Split(splitter);
                foreach (string s in arrTo)
                {
                    mm.To.Add(s);
                }

                // For Cc Field
                if (!string.IsNullOrEmpty(strCc))
                {
                    Array arrCc;
                    arrCc = strCc.Split(splitter);
                    foreach (string s in arrCc)
                    {
                        mm.CC.Add(s);
                    }
                }

                // For Bcc Field
                if (!string.IsNullOrEmpty(strBcc))
                {
                    Array arrBcc;
                    arrBcc = strBcc.Split(splitter);
                    foreach (string s in arrBcc)
                    {
                        mm.Bcc.Add(s);
                    }
                }
                // For Subject
                mm.Subject = strSubject;

                // For Attachments
                if (!string.IsNullOrEmpty(strAttachmentPath))
                {
                    Array arrToAttach;
                    arrToAttach = strAttachmentPath.Split(splitter);

                    if (arrToAttach.Length != 0)
                    {
                        foreach (string s1 in arrToAttach)
                        {
                            if (File.Exists(s1))
                            {
                                Attachment attachFile = new Attachment(s1);
                                mm.Attachments.Add(attachFile);
                            }
                        }
                    }
                }

                // For Body
                mm.IsBodyHtml = true;
                mm.Body = strBody;

                // For News Letters            
                if (strNewsLetter == "R")
                {
                    #region Old
                    //string shellScript = strOracleRun + "  " + strOracleReportPath + strReportName + "  " +
                    //                     strCondParam + "=\"" + strViewName + " " + strCondValue + "\"" + "  " +
                    //                     "PClient=" + "\'" + strClient + "\'" + " " +
                    //                     "userid=" + strUserName + "/" + strPassWord + "@" + strServerName + "  " +
                    //                     "destype=file desformat=pdf batch=yes mode=BITMAP desname=" + "\"" + AppPath + "\\" + strReportFileName + ".pdf" + "\"";
                    ////"destype=file desformat=pdf batch=yes mode=BITMAP desname=" +"\"" + @"C:\R" + "\\" + strReportFileName + ".pdf" + "\"";
                    //eventLogFile.WriteEntry(shellScript);
                    //System.Diagnostics.Process process1;
                    //process1 = new System.Diagnostics.Process();

                    //string strCmdLine;
                    //strCmdLine = "/C " + shellScript;

                    //process1.StartInfo.FileName = "cmd.exe";
                    //process1.StartInfo.Arguments = strCmdLine;
                    //process1.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                    //process1.StartInfo.CreateNoWindow = false;
                    //process1.StartInfo.UseShellExecute = true;
                    //process1.Start();
                    //System.Threading.Thread.Sleep(60000);
                    //process1.Close();

                    ////string fName = @"C:\R" + "\\" + strReportFileName + ".pdf";
                    #endregion

                    string FileName = string.Empty;
                    FileName = NewsLetterAttachment();

                    if (!string.IsNullOrEmpty(FileName))
                    {
                        if (File.Exists(FileName))
                        {
                            mm.Attachments.Add(new Attachment(FileName));
                        }
                    }
                }

                //mm.Body = strBody;

                SmtpClient smtpC = new SmtpClient();

                if (!string.IsNullOrEmpty(strFrom))
                {
                    string[] EmailElements = strFrom.Split('@');
                    string EmailId = string.Empty;
                    string Domain = string.Empty;
                    if (EmailElements.Length == 2)
                    {
                        EmailId = EmailElements[0];
                        Domain = EmailElements[1];
                        if (!string.IsNullOrEmpty(EmailId) && !string.IsNullOrEmpty(Domain))
                        {
                            mm.From = new MailAddress(strSmtpUid, EmailId);
                        }
                        else
                        {
                            mm.From = new MailAddress(strSmtpUid, strFromName);
                        }
                    }
                    else
                    {
                        mm.From = new MailAddress(strSmtpUid, strFromName);
                    }
                }
                else
                {
                    mm.From = new MailAddress(strSmtpUid, strFromName);
                }

                NetworkCredential netCred = new NetworkCredential(strSmtpUid, strSmtpPwd);
                smtpC.Host = strSmtpServer;
                smtpC.Port = Convert.ToInt16(strSmtpPort);
                if (strEnableSSL.ToUpper() == "Y")
                {
                    smtpC.EnableSsl = true;
                }
                else
                {
                    smtpC.EnableSsl = false;
                }
                smtpC.UseDefaultCredentials = false;
                smtpC.Credentials = netCred;
                smtpC.Send(mm);
                return true;
            }
            catch (Exception ex)
            {
                eventLogFileOEE.WriteEntry(ex.Message.ToString());
                LogExceptions(strLogFilePath, ex);
                return false;
            }
        }

       
        private ItemIdType CreateDraftMessage(ExchangeServiceBinding ewsServiceBinding, MessageType emMessage)
        {
            ItemIdType iiItemid = new ItemIdType();
            CreateItemType ciCreateItemRequest = new CreateItemType();
            ciCreateItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            ciCreateItemRequest.MessageDispositionSpecified = true;
            ciCreateItemRequest.SavedItemFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType dfDraftsFolder = new DistinguishedFolderIdType();
            dfDraftsFolder.Id = DistinguishedFolderIdNameType.drafts;
            ciCreateItemRequest.SavedItemFolderId.Item = dfDraftsFolder;
            ciCreateItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            ciCreateItemRequest.Items.Items = new ItemType[1];
            ciCreateItemRequest.Items.Items[0] = emMessage;
            CreateItemResponseType createItemResponse = ewsServiceBinding.CreateItem(ciCreateItemRequest);
            if (createItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
            {
                //Console.WriteLine("Error Occured");
                //Console.WriteLine(createItemResponse.ResponseMessages.Items[0].MessageText);
                eventLogFileOEE.WriteEntry(createItemResponse.ResponseMessages.Items[0].MessageText);
            }
            else
            {
                ItemInfoResponseMessageType rmResponseMessage = createItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
                //Console.WriteLine("Item was created");
                //Console.WriteLine("Item ID : " + rmResponseMessage.Items.Items[0].ItemId.Id.ToString());
                //Console.WriteLine("ChangeKey : " + rmResponseMessage.Items.Items[0].ItemId.ChangeKey.ToString());
                iiItemid.Id = rmResponseMessage.Items.Items[0].ItemId.Id.ToString();
                iiItemid.ChangeKey = rmResponseMessage.Items.Items[0].ItemId.ChangeKey.ToString();
            }

            return iiItemid;
        }

        private ItemIdType CreateAttachment(ExchangeServiceBinding ewsServiceBinding, String fnFileName, ItemIdType iiCreateItemid)
        {
            string fName = Path.GetFileName(fnFileName);
            ItemIdType iiAttachmentItemid = new ItemIdType();
            FileStream fsFileStream = new FileStream(fnFileName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            byte[] bdBinaryData = new byte[fsFileStream.Length];
            long brBytesRead = fsFileStream.Read(bdBinaryData, 0, (int)fsFileStream.Length);
            fsFileStream.Close();
            FileAttachmentType faFileAttach = new FileAttachmentType();
            faFileAttach.Content = bdBinaryData;
            //faFileAttach.Name = fnFileName; // Old
            faFileAttach.Name = fName;
            CreateAttachmentType amAttachmentMessage = new CreateAttachmentType();
            amAttachmentMessage.Attachments = new AttachmentType[1];
            amAttachmentMessage.Attachments[0] = faFileAttach;
            amAttachmentMessage.ParentItemId = iiCreateItemid;
            CreateAttachmentResponseType caCreateAttachmentResponse = ewsServiceBinding.CreateAttachment(amAttachmentMessage);
            if (caCreateAttachmentResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
            {
                //Console.WriteLine("Error Occured");
                //Console.WriteLine(caCreateAttachmentResponse.ResponseMessages.Items[0].MessageText);
                eventLogFileOEE.WriteEntry(caCreateAttachmentResponse.ResponseMessages.Items[0].MessageText);
            }
            else
            {
                AttachmentInfoResponseMessageType amAttachmentResponseMessage = caCreateAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;
                //Console.WriteLine("Attachment was created");
                //Console.WriteLine("Change Key : " + amAttachmentResponseMessage.Attachments[0].AttachmentId.RootItemChangeKey.ToString());
                iiAttachmentItemid.Id = amAttachmentResponseMessage.Attachments[0].AttachmentId.RootItemId.ToString();
                iiAttachmentItemid.ChangeKey = amAttachmentResponseMessage.Attachments[0].AttachmentId.RootItemChangeKey.ToString();
            }
            return iiAttachmentItemid;
        }

        private void SendMessage(ExchangeServiceBinding ewsServiceBinding, ItemIdType iiCreateItemid)
        {
            SendItemType siSendItem = new SendItemType();
            siSendItem.ItemIds = new BaseItemIdType[1];
            siSendItem.SavedItemFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType siSentItemsFolder = new DistinguishedFolderIdType();
            siSentItemsFolder.Id = DistinguishedFolderIdNameType.sentitems;
            siSendItem.SavedItemFolderId.Item = siSentItemsFolder;
            siSendItem.SaveItemToFolder = true; ;
            siSendItem.ItemIds[0] = (BaseItemIdType)iiCreateItemid;
            SendItemResponseType srSendItemReponseMessage = ewsServiceBinding.SendItem(siSendItem);
            if (srSendItemReponseMessage.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
            {
                //Console.WriteLine("Error Occured");
                //Console.WriteLine(srSendItemReponseMessage.ResponseMessages.Items[0].MessageText);
                eventLogFileOEE.WriteEntry(srSendItemReponseMessage.ResponseMessages.Items[0].MessageText + strSno);
            }
            else
            {
                Console.WriteLine("Message Sent");
                //eventLogFile.WriteEntry("Message Sent to : " + strSno );
            }
        }

        private string NewsLetterAttachment()
        {
            try
            {
                string shellScript = strOracleRun + "  " + strOracleReportPath + strReportName + "  " +
                                            strCondParam + "=\"" + strViewName + " " + strCondValue + "\"" + "  " +
                                            "PClient=" + "\'" + strClient + "\'" + " " +
                                            "userid=" + strUserName + "/" + strPassWord + "@" + strServerName + "  " +
                                            "destype=file desformat=pdf batch=yes mode=BITMAP desname=" + "\"" + AppPath + "\\" + strReportFileName + ".pdf" + "\"";
                //"destype=file desformat=pdf batch=yes mode=BITMAP desname=" +"\"" + @"C:\R" + "\\" + strReportFileName + ".pdf" + "\"";
                eventLogFileOEE.WriteEntry(shellScript);
                System.Diagnostics.Process process1;
                process1 = new System.Diagnostics.Process();

                string strCmdLine;
                strCmdLine = "/C " + shellScript;

                process1.StartInfo.FileName = "cmd.exe";
                process1.StartInfo.Arguments = strCmdLine;
                process1.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                process1.StartInfo.CreateNoWindow = false;
                process1.StartInfo.UseShellExecute = true;
                process1.Start();
                System.Threading.Thread.Sleep(60000);
                process1.Close();

                //string fName = @"C:\R" + "\\" + strReportFileName + ".pdf";
                fName = AppPath + "\\" + strReportFileName + ".pdf";

                return fName;
            }
            catch (Exception ex)
            {
                //fName = "";
                eventLogFileOEE.WriteEntry(ex.Message.ToString());
                LogExceptions(strLogFilePath, ex);
                return fName; 
            }
        }
    }
}
