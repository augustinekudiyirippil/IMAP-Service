using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Timers;


using System.Drawing;


using ActiveUp.Net.Mail;
//using GmailReadImapEmail;
using Message = ActiveUp.Net.Mail.Message;
using System.Data.SqlClient;
using System.Configuration;
using System.Text.RegularExpressions;

using System.Net.Mail;
using AKIMAPService;

namespace AKImapService
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();
        public Service1()
        {
            InitializeComponent();




        }



        string strSubjectFirstPart, strSubjectSecondPart;
        public string Between(string STR, string FirstString, string LastString)
        {
            string FinalString;




            try
            {
                int Pos1 = STR.IndexOf(FirstString) + FirstString.Length - 1;
                int Pos2 = STR.IndexOf(LastString);
                FinalString = STR.Substring(Pos1, Pos2 - Pos1);

            }
            catch (Exception ee)
            {

                string strerr = ee.Message.ToString();
                FinalString = "NoMatchFound";
            }

            return FinalString;




        }
        string strResult, strError;










        public void readIMAPEmail()
        {


            string strAttchmentPath = ConfigurationManager.AppSettings["IncomingAttachmentPath"];

            strSubjectFirstPart = ConfigurationManager.AppSettings["SubjectFirstPart"];

            strSubjectSecondPart = ConfigurationManager.AppSettings["SubjectSecondPart"];


            string strServerAddress, strIncomingUsername, strIncomingEmailPassword, strMailBoxID;
            int strPort = 993;
            string connectionString = ConfigurationManager.AppSettings["ConnectionString"];
            string strQuery = "select mbid, mbEmailAddress , mbIncomingServerAddress , mbIncomingUsername , mbIncomingPassword ";
            strQuery = strQuery + "    from MainBoxTable where mbIncomingEnabled =1 and mbIsDeleted=0";

            SqlConnection sqlConnection = new SqlConnection(connectionString);
            SqlCommand sqlCommand = new SqlCommand();
            sqlConnection.Open();


            sqlCommand = new SqlCommand(strQuery, sqlConnection);

            SqlDataReader reader = sqlCommand.ExecuteReader();


            while (reader.Read())
            {

                strServerAddress = reader["mbIncomingServerAddress"].ToString();
                strMailBoxID = reader["mbid"].ToString();

                strIncomingUsername = reader["mbIncomingUsername"].ToString();
                strIncomingEmailPassword = reader["mbIncomingPassword"].ToString();



                //strIncomingUsername = "research@simplisysltd.onmicrosoft.com";
                //strIncomingEmailPassword = "nzdmdrnpcgsycwcv";

                //strIncomingUsername = "simplisystesting@gmail.com";
                //strIncomingEmailPassword = "gswcirtxeulwcjyj";
                //strServerAddress = "imap.gmail.com";



                var mailRepository = new MailRepository(
                            strServerAddress,
                           strPort,
                           true,
                         strIncomingUsername,
                         strIncomingEmailPassword
                       );





                var flags = new FlagCollection { "Seen" };
                var emailList = mailRepository.GetUnreadMails("inbox");
                string content = "";
                string strFrom, strCC, strSubject, strBody, strHTMLBody, strAttachments, strMessageID, strSentDate, strReceivedDate, strDisplayName;

                //@incID uniqueidentifier,
                //@inemlID   uniqueidentifier,  



                Guid strIncID;

                Guid strIncidentEmailID;

                Guid inEmailID;





                string strDateString, strDate, strYear, strMonth, strDay, strAttachmentName = "";



                try // FIRST TRY BEGINS
                {
                    foreach (Message email in emailList) // EMAIL LIST FOR LOOP
                    {

                        content += " " + " " + email.From + " " + email.Subject + " " + " " + email.BodyText.Text + " " + Environment.NewLine + Environment.NewLine;

                        //if (email.Attachments.Count<1)
                        //{
                        //    continue;
                        //}
                        strFrom = ExtractEmails(email.From.ToString());
                        strCC = email.Cc.ToString();
                        strSubject = email.Subject.ToString();
                        strHTMLBody = email.BodyHtml.Text.ToString();
                        strBody = email.BodyText.Text.ToString();
                        //strBody = email.BodyHtml.Text.ToString();  // Added for testing.. to be removed
                        strAttachments = email.Attachments.ToString();
                        strMessageID = email.MessageId.ToString();
                        strSentDate = email.Date.ToString();
                        strReceivedDate = email.ReceivedDate.ToString();
                        strDisplayName = email.From.Name.ToString();

                        strDate = email.Date.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'");




                        strIncID = Guid.NewGuid();
                        strIncidentEmailID = Guid.NewGuid();

                        inEmailID = new Guid(strMailBoxID);
                        string strFilePath = strAttchmentPath + "\\" + strMessageID; // Your code goes here



                        try
                        {
                            System.IO.Directory.CreateDirectory(strFilePath);
                        }
                        catch (Exception ex)
                        {
                            strError = ex.Message.ToString();
                        }




                        strResult = Between(strSubject, strSubjectFirstPart, strSubjectSecondPart);


                        string var = strResult;
                        string mystr = Regex.Replace(var, @"\d", "");
                        string mynumber = Regex.Replace(var, @"\D", "");

                        strResult = mynumber;



                        string strIncidentNumber = "INCNUMBER";


                        if (strResult != "NoMatchFound")
                        {
                            try
                            {

                                SqlConnection sqlConnectionChkInc = new SqlConnection(connectionString);
                                SqlCommand sqlCommandChkInc = new SqlCommand();
                                sqlConnectionChkInc.Open();

                                string strQueryChkIncident = "select col1, col2 from ** where col='" + strResult + "'";

                                sqlCommandChkInc = new SqlCommand(strQueryChkIncident, sqlConnectionChkInc);

                                SqlDataReader readerChkInc = sqlCommandChkInc.ExecuteReader();

                                while (readerChkInc.Read())
                                {
                                    strIncID = Guid.Parse(readerChkInc["incID"].ToString());

                                    strIncidentNumber = readerChkInc["IncNumber"].ToString();
                                }
                                readerChkInc.Close();
                                sqlConnectionChkInc.Close();

                            }
                            catch (Exception exc)
                            {
                                string err = exc.Message.ToString();
                            }
                        }



                        string strSQLProcedure = "";

                        if (strIncidentNumber == strResult)
                        {
                            strSQLProcedure = "SPUpdate"; // "SPUpdateIncFrmEmail";

                        }
                        else
                        {

                            strSQLProcedure = "SPInsertTo";

                        }


                        connectionString = ConfigurationManager.AppSettings["ConnectionString"];

                        using (SqlConnection con = new SqlConnection(connectionString))
                        {
                            using (SqlCommand cmd = new SqlCommand(strSQLProcedure, con))
                            {
                                cmd.CommandType = CommandType.StoredProcedure;

                                cmd.Parameters.AddWithValue("@**", strIncID);
                                cmd.Parameters.AddWithValue("@**", strIncidentEmailID);
                                cmd.Parameters.AddWithValue("@**", inEmailID);



                                cmd.Parameters.AddWithValue("@**", strSubject);
                                cmd.Parameters.AddWithValue("@**", 1);   //@inemlIsBodyHTML bit,
                                cmd.Parameters.AddWithValue("@**", strHTMLBody); //@inemlHTMLBody nvarchar(max),  
                                cmd.Parameters.AddWithValue("@**", strBody);//@inemlTextBody nvarchar(max),  
                                strDateString = strSentDate;





                                cmd.Parameters.AddWithValue("@**", strDate); //@inemlDateAdded datetime,  
                                cmd.Parameters.AddWithValue("@**", "Normal");//@inemlImportance nvarchar(50),  
                                cmd.Parameters.AddWithValue("@**", strMessageID); //@inemlMessageID nvarchar(250),  
                                cmd.Parameters.AddWithValue("@**", strFrom);//@inemlFromAddress nvarchar(320),  
                                cmd.Parameters.AddWithValue("@**", strDisplayName); //@inemlFromDisplayName nvarchar(260),  
                                cmd.Parameters.AddWithValue("@**", 0); //@inemlBlockedBySizeLimit bit,
                                cmd.Parameters.AddWithValue("@**", 0);//@inemlBlockedByBlacklist bit,  
                                cmd.Parameters.AddWithValue("@**", 0); //@inemlBlockedBySpam bit,
                                cmd.Parameters.AddWithValue("@**", 0);//@inemlFailedDueToError bit,  
                                cmd.Parameters.AddWithValue("@**", 0);//@inemlFailedDueToErrorBeforeDeletion bit,
                                cmd.Parameters.AddWithValue("@**", 0);//@inemlIsDeleted bit
                                cmd.Parameters.AddWithValue("@**", strFilePath + "\\email.eml");


                                con.Open();
                                cmd.ExecuteNonQuery();
                                // con.Close();
                            }
                        }


                        string strfile = "";

                        string strEmbeddedFilename, strEmbeddedContentID, strEmbeddedFileFullPath = "";


                        //create a folder inside the path with messageid
                        //NORMAL ATTACHMENTS
                        if (email.Attachments.Count > 0)
                        {

                            email.Attachments.StoreToFolder(strFilePath);

                            for (int i = 0; i < email.Attachments.Count; i++)
                            {
                                // strEmbeddedContentID = email.Attachments[i].ContentId.ToString();


                                using (SqlConnection con = new SqlConnection(connectionString))
                                {
                                    using (SqlCommand cmd = new SqlCommand("SPInsertInToAttachments", con))
                                    {
                                        cmd.CommandType = CommandType.StoredProcedure;

                                        cmd.Parameters.AddWithValue("@**", strIncID);     ////// @incID uniqueidentifier,
                                        cmd.Parameters.AddWithValue("@**", strIncidentEmailID);
                                        cmd.Parameters.AddWithValue("@**", strFilePath + "\\" + email.Attachments[i].Filename.ToString());        //////@attDiskPath varchar(500),
                                        cmd.Parameters.AddWithValue("@**", email.Attachments[i].Filename.ToString());            //////@attFileName varchar(150),
                                        cmd.Parameters.AddWithValue("@**", "");
                                        cmd.Parameters.AddWithValue("@**", "File added by incoming email from " + strFrom); //////@attNotes varchar(250),
                                        cmd.Parameters.AddWithValue("@**", email.Attachments[i].Size);   //////@attSize bigint
                                        cmd.Parameters.AddWithValue("@**", 0);
                                        con.Open();
                                        cmd.ExecuteNonQuery();

                                    }
                                }


                            }
                        }




                        string strTrimmedContentID;
                        string strCommandText = "";
                        //Inline attachments
                        try
                        {
                            for (int i = 0; i < email.LeafMimeParts.Count; i++)
                            {


                                if (email.LeafMimeParts[i].IsText == false)
                                {

                                    strEmbeddedFilename = email.LeafMimeParts[i].ContentName.ToString();
                                    strEmbeddedFilename = email.LeafMimeParts[i].Filename.ToString();
                                    strEmbeddedFileFullPath = strFilePath + "\\" + strEmbeddedFilename;
                                    strEmbeddedContentID = email.LeafMimeParts[i].ContentId.ToString();

                                    strTrimmedContentID = email.LeafMimeParts[i].ContentId.Replace('<', ' ');
                                    strTrimmedContentID = strTrimmedContentID.Replace('>', ' ');

                                    strEmbeddedContentID = strTrimmedContentID.Trim();


                                    email.LeafMimeParts[i].StoreToFile(strEmbeddedFileFullPath);


                                    if (strEmbeddedContentID.Length > 0)
                                    {
                                        strCommandText = "SPInsertEmbeddedInToAttachments";
                                    }
                                    else
                                    {
                                        strCommandText = "SPInsertInToAttachments";

                                    }


                                    try
                                    {

                                        using (SqlConnection con = new SqlConnection(connectionString))
                                        {
                                            //SPInsertEmbeddedInToAttachments  
                                            using (SqlCommand cmd = new SqlCommand(strCommandText, con))
                                            // using (SqlCommand cmd = new SqlCommand("SPInsertInToAttachments", con))
                                            {
                                                cmd.CommandType = CommandType.StoredProcedure;

                                                cmd.Parameters.AddWithValue("@**", strIncID);     ////// @incID uniqueidentifier,
                                                cmd.Parameters.AddWithValue("@**", strIncidentEmailID);
                                                cmd.Parameters.AddWithValue("@**", strEmbeddedFileFullPath);        //////@attDiskPath varchar(500),
                                                cmd.Parameters.AddWithValue("@**", strEmbeddedFilename);            //////@attFileName varchar(150),
                                                cmd.Parameters.AddWithValue("@**", strEmbeddedContentID);
                                                cmd.Parameters.AddWithValue("@**", "File added by incoming email from " + strFrom); //////@attNotes varchar(250),
                                                cmd.Parameters.AddWithValue("@**", email.LeafMimeParts[i].Size);   //////@attSize bigint
                                                cmd.Parameters.AddWithValue("@**", 1);


                                                con.Open();
                                                cmd.ExecuteNonQuery();

                                            }
                                        }


                                    }
                                    catch (Exception exc)
                                    {
                                        string err = exc.Message.ToString();
                                    }


                                }

                            }


                        }
                        catch (Exception exc)
                        {
                            string err = exc.Message.ToString();
                        }


                      

                        mailRepository.Mails.RemoveFlagsSilent(email.Id, flags);    //Commented on 14th June 2022
                        mailRepository.Mails.SetFlagsSilent(email.Id, flags);


                        mailRepository.Mails.DeleteMessage(email.Id, true);


                    }//END FOR EMAIL LIST FOR LOOP
                }//END FIRST TRY
                catch (Exception ex)
                {

                    strError = ex.Message.ToString();
                }

                var allEmailList = mailRepository.GetAllMails("Inbox");
                foreach (Message email in allEmailList) // EMAIL LIST FOR LOOP
                {
                    mailRepository.Mails.DeleteMessage(email.Id, true);

                }







            }// END WHILE LOOP
            reader.Close();
            sqlConnection.Close();


        }


        protected override void OnStart(string[] args)
        {
            //WriteToFile("Service is started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 12000; //Convert.ToDouble ( ConfigurationManager.AppSettings["EmailReadingIntervals"])  ;  // 300000; //number in milisecinds  //EmailReadingIntervals
            timer.Enabled = true;
        }

        public void onDebug()
        {
            OnStart(null);
        }

        protected override void OnStop()
        {

            //WriteToFile("Service is stopped at " + DateTime.Now);
        }

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\AK_IMAP_ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }


        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            readIMAPEmail();
            //WriteToFile("Service is recall at " + DateTime.Now);
        }

        public static string ExtractEmails(string emailAddress)
        {

            string streEmailAddress = "";
            //instantiate with this pattern 
            Regex emailRegex = new Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*",
                RegexOptions.IgnoreCase);
            //find items that matches with our pattern
            MatchCollection emailMatches = emailRegex.Matches(emailAddress);

            StringBuilder sb = new StringBuilder();

            foreach (Match emailMatch in emailMatches)
            {
                //sb.AppendLine(emailMatch.Value);
                streEmailAddress = emailMatch.Value;
            }
            //store to file

            return streEmailAddress;


        }

    }
}
