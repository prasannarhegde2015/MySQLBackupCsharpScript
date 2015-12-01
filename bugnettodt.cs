using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.Sql;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using Selenium_Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support;
using OpenQA.Selenium.Support.UI;
using System.Windows.Forms;
using System.Threading;
using System.Windows.Automation;
using System.Web;

#region Comments
/* ***********************Change Histroy***********************
    * Created By : Prasanna Hegde;
    * Creation Date: 28-03-2013;
     */
# region 29-03-2013
/*    Developer : Deepankar
     *  Added additional comments and timeout in App.Config
     */
# endregion
# region 27th May 2013
/*
     Added function ReplaceHTMLTags to replace the html tags that are part of the issue description 
     Added key htmltagfile to appconfig file
     Added code to close SQl Connection object   
 */
#endregion
# region 28th May 2013
/*
     Modified function ReplaceHTMLTags with a regular expression replace instead of reading from a textfile
     
 */
#endregion
# region 29th May 2013
/*
     Modified code to close the modal window. Using Sendkeys instead of UI Automation 
     
 */
#endregion
# region 30th May 2013
/*
     Added function ConvertTicket and called from Main to make the process run continuously
     
 */
#endregion
#region #12th June 2103
//Moved the code related to command object inside IterateThroughMappingFile
#endregion 
#region #27th June 2013
/* Developer : Sneha
    Updated function SendEmails to send an email to multiple recipients
 */
#endregion
#endregion
namespace Bugnet2Devtrack
{
    class Bugnet2Devtrack
    {
        static void Main(string[] args)
        {
            int count = 1;
            while (true)
            {
                Console.WriteLine("Starting Ticket Conversion: Run {0}",count.ToString());
                ConvertTicket();
                Console.Write("Waiting for 2 Minutes..");
                System.Threading.Thread.Sleep(120000);
                SendKeys.Flush();
                SendKeys.SendWait("%{f4}");
                count++;
                Console.Clear();
            }
       
        } //end of function

        public static void ConvertTicket()
        {
            SqlConnection con = new System.Data.SqlClient.SqlConnection();

            try
            {
                string loggedInUser = ConfigurationManager.AppSettings["loginname"];
                string userpasswd = ConfigurationManager.AppSettings["loginpassword"];
                Console.WriteLine("**********************************1) Launch Browser Google Chrome****************************");
                #region Selenium_Launch_browser
                Selenium_Framework.Selenium_Framework test = new Selenium_Framework.Selenium_Framework();
                test.launchweb(Selenium_Framework.Selenium_Framework.appBrowser.Chrome,ConfigurationManager.AppSettings["devtrackurl"]);

                #endregion Selenium_Launch_browser
                #region SeleniumLogin
                Console.WriteLine("**********************************1) Launch Browser Google Chrome: Done ****************************");
                Console.WriteLine("**********************************2) Login with Super User Credentails****************************");
                test.perfaction("text", "name", "Username", loggedInUser);
                test.perfaction("text", "name", "Password", userpasswd);
                test.perfaction("img", "src", "/PTWeb/images/button_login.gif", "Login");
                Thread.Sleep(5000);
                Console.WriteLine("**********************************2) Login with Super User Credentails: Done****************************");
                Console.WriteLine("**********************************3) Establish Database Connnection ****************************");
                #endregion SeleniumLogin

                DataTable dt;
                #region Connection
                string server = "";
                con.ConnectionString = ConfigurationManager.ConnectionStrings["con"].ConnectionString;
                server = ConfigurationManager.AppSettings["server"];
                con.Open();
                Console.WriteLine(con.State);
                #endregion
             
                string strmapfile = ConfigurationManager.AppSettings["mapfile"];
                string submityesno = ConfigurationManager.AppSettings["testsubmit"];
                DataTable tableMapping = CreateDataTableFromExcel(strmapfile, "Select * from [mapping$]");
                string devtrackID = "";
                int closeStatusID = -1;
                //************** Number of Projects to be picked from Mapping Excel files *************************
                #region IterateThroughMappingFile
                for (int ij = 0; ij < tableMapping.Rows.Count; ij++)
                {

                    if (tableMapping.Rows[ij]["ExecuteSp"].ToString().ToLower() == "y")
                    {
                        #region executestoredprocedureinitialize
                        dt = new DataTable();
                        SqlCommand command = new SqlCommand();
                        command.Connection = con;
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = "MPCQA_Custom_GetDataForDevtrack";
                        #endregion executestoredprocedureinitialize
                        Console.WriteLine("Processing for: {0}", tableMapping.Rows[ij]["Application"]);

                        #region input_parameters_for_SP_from_excel


                        devtrackID = tableMapping.Rows[ij]["CustomFieldDevTrackId"].ToString();
                        closeStatusID = int.Parse(tableMapping.Rows[ij]["CloseStatusID"].ToString());
                        command.Parameters.Add(new SqlParameter("@projectID", tableMapping.Rows[ij]["Projectid"].ToString()));
                        command.Parameters.Add(new SqlParameter("@CustomFieldDevTrackID", devtrackID));
                        command.Parameters.Add(new SqlParameter("@IssueResolutionID", tableMapping.Rows[ij]["ResolutionId"].ToString()));
                        command.Parameters.Add(new SqlParameter("@CustomFieldFoundinVersionID", tableMapping.Rows[ij]["CustomFieldFoundInVersionId"].ToString()));
                        command.Parameters.Add(new SqlParameter("@CustomFieldSeverityId", tableMapping.Rows[ij]["CustomFieldSeverityId"].ToString()));
                        command.Parameters.Add(new SqlParameter("@CustomFieldComponentId", tableMapping.Rows[ij]["CustomFieldComponentId"].ToString()));
                        #endregion input_parameters_for_SP_from_excel
                        SqlDataAdapter adapter = new SqlDataAdapter(command);
                        adapter.Fill(dt);
                        Console.WriteLine(dt.Rows.Count.ToString());
                        command.Parameters.Clear();
                        string strbugnetissueid = "";
                        string strIssueDesc = "";
                        string issueTitle = "";
                        //****************** Number of Bugnet Observations to be logged as Devtrack Tickets ****************
                        Console.WriteLine("**********************************3) Establish Database Connnection : Done Records count" + dt.Rows.Count + "****************************");
                        #region IterateThroughTickets
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            Console.WriteLine("**********************************4) Iteration for Record " + i + "****************************");

                            strIssueDesc = "";
                            strbugnetissueid = dt.Rows[i]["IssueId"].ToString();
                            Console.WriteLine("Issue Title:" + dt.Rows[i]["issueTitle"].ToString());
                            Console.WriteLine("Priority:" + dt.Rows[i]["PriorityName"].ToString());
                            Console.WriteLine("Found in Verrsion:" + dt.Rows[i]["FoundInVersion"].ToString());
                            Console.WriteLine("Severity:" + dt.Rows[i]["Severity"].ToString());
                            Console.WriteLine("IssueDescription:" + dt.Rows[i]["IssueDescription"].ToString());
                            #region Selenium_actions_on_UI
                            //********************* 1. Click on 'New' issue button **********************************
                            Thread.Sleep(5000);
                            test.goToFirstFrame();
                            test.perfaction("frame", "name", "spider1", "spider1");
                            test.perfaction("Link", "onClick", "return issuenew()", "New");
                            Thread.Sleep(5000);
                            test.goToFirstFrame();
                            test.perfaction("frame", "name", "spider2", "spider2");
                            test.perfaction("frame", "name", "buglist2", "buglist2");
                            test.perfaction("iframe", "name", "newbugIFrame", "newbugIFrame");
                            test.perfaction("iframe", "name", "page", "page");
                            // 1.a *************** Handling Special  Characters in string retreived from database.
                            issueTitle = dt.Rows[i]["issueTitle"].ToString();
                            issueTitle = issueTitle.Replace("&#39;", "'");
                            issueTitle = issueTitle.Replace("&nbsp;", "");
                            issueTitle = issueTitle.Replace("&rsquo;", "`");
                            issueTitle = issueTitle.Replace("&lsquo;", "`");
                            //********************* 2. Enter Values for all fields **********************************
                            test.perfaction("text", "name", "devcustfield101", issueTitle); // Issue Title
                            test.perfaction("dropdown", "name", "devcustfield103", "Issue"); //Issue Type
                            test.perfaction("dropdown", "name", "devcustfield106", dt.Rows[i]["FoundInVersion"].ToString()); //Found in Version
                            test.perfaction("dropdown", "name", "devcustfield104", dt.Rows[i]["PriorityName"].ToString()); //Priority
                            test.perfaction("dropdown", "name", "devcustfield107", dt.Rows[i]["Severity"].ToString());//Severity
                            test.perfaction("dropdown", "name", "devcustfield105", tableMapping.Rows[ij]["Application"].ToString()); //Application
                            test.perfaction("dropdown", "name", "devcustfield13", dt.Rows[i]["FoundInVersion"].ToString()); //Traget Version
                            test.perfaction("dropdown", "name", "devcustfield14", dt.Rows[i]["Component"].ToString()); //Component
                            test.perfaction("dropdown", "name", "devcustfield15", "QA"); // Originating Org
                            // 2.a *************** Handling Special  Characters in string retreived from database. ********************************************
                            //  strIssueDesc = strIssueDesc + "Bugnet Issue ID:  " + dt.Rows[i]["IssueID"].ToString() + Environment.NewLine;
                            strIssueDesc = strIssueDesc + Environment.NewLine + dt.Rows[i]["IssueDescription"].ToString();

                            strIssueDesc = strIssueDesc.Replace("\r\n", "");
                            strIssueDesc = strIssueDesc.Replace("\t", "");
                            strIssueDesc = strIssueDesc.Replace("<p>", "");
                            strIssueDesc = strIssueDesc.Replace("&nbsp;", "");
                            strIssueDesc = strIssueDesc.Replace("&gt;", "");
                            strIssueDesc = strIssueDesc.Replace("&lt;", "");
                            strIssueDesc = strIssueDesc.Replace("&#39;", "");
                            strIssueDesc = strIssueDesc.Replace("&rsquo;", "`");
                            strIssueDesc = strIssueDesc.Replace("&lsquo;", "`");
                            strIssueDesc = strIssueDesc.Replace("</p>", Environment.NewLine);
                            strIssueDesc = strIssueDesc.Replace("<div>", "");
                            strIssueDesc = strIssueDesc.Replace("</div>", Environment.NewLine);
                            strIssueDesc = strIssueDesc.Replace("<strong>", "");
                            strIssueDesc = strIssueDesc.Replace("</strong>", "");
                            strIssueDesc = strIssueDesc.Replace("\r\n\r\n", Environment.NewLine);
                            strIssueDesc = ReplaceHTMLTags(strIssueDesc);
                            strIssueDesc = "Bugnet Issue ID:  " + dt.Rows[i]["IssueID"].ToString() + Environment.NewLine + strIssueDesc;
                            // 2. a ****************************************************************************************************************************
                            test.perfaction("textarea", "name", "devcustfield102", strIssueDesc); //description 
                            loggedInUser = ConfigurationManager.AppSettings["assignedto"];
                            test.perfaction("dropdown", "name", "devcustfield108", loggedInUser); //Assigned to (seems like initially it can be assigned to originator only)
                            Thread.Sleep(2000);
                            #region UploadingAttchmentsinDevTrack
                            Console.WriteLine("Trying to upload attachments");
                            int attachcount = 0;
                            attachcount = GetIssueAttachmentsCount(strbugnetissueid);

                            string returnPath = AddAttachments(strbugnetissueid).ToString();
                            char[] celldellim = new char[] { ';' };
                            string[] arrcelladd = returnPath.Split(celldellim);
                            for (int ik = 0; ik < attachcount; ik++)
                            {
                                if (ik == 0)
                                {
                                    Thread.Sleep(2000);
                                    test.perfaction("button", "name", "devcustfield132", "Choose File");
                                    Thread.Sleep(2000);
                                    SendKeys.Flush();
                                    StringBuilder temp = new StringBuilder(@arrcelladd[0]);
                                    temp.Replace(@"\\\\", @"\\");
                                    Console.WriteLine("upload path : " + temp);
                                    Thread.Sleep(2000);
                                    //Console.ReadLine();
                                    Console.WriteLine(temp.ToString());
                                    SendKeys.SendWait(temp.ToString());
                                    //    Thread.Sleep(2000);
                                    SendKeys.SendWait("{Enter}");
                                    Thread.Sleep(2000);

                                }
                                else if (ik > 0)
                                {

                                    test.perfaction("Link", "Linktext", "Attach another file", "Attach another file");
                                    test.perfaction("button", "name", "devcustfieldfile100000" + ik.ToString(), "Choose File");
                                    Thread.Sleep(2000);
                                    SendKeys.Flush();
                                    StringBuilder temp1 = new StringBuilder(@arrcelladd[ik]);
                                    temp1.Replace(@"\\\\", @"\\");
                                    Console.WriteLine("upload path : " + temp1);
                                    Thread.Sleep(2000);
                                    //  Console.ReadLine();
                                    Console.WriteLine(temp1.ToString());
                                    SendKeys.SendWait(temp1.ToString());
                                    Thread.Sleep(2000);

                                    // Console.ReadLine();
                                    //  SendKeys.SendWait("test");
                                    // SendKeys.Flush();
                                    // Console.ReadLine();
                                    SendKeys.SendWait("{Enter}");
                                    Thread.Sleep(2000);

                                }
                            }
                            #endregion
                            test.goToFirstFrame();
                            test.perfaction("frame", "name", "spider2", "spider2");
                            test.perfaction("frame", "name", "buglist2", "buglist2");
                            test.perfaction("iframe", "name", "newbugIFrame", "newbugIFrame");
                            test.perfaction("tabledata", "Linktext", "^Project Accounting$", "Project Accounting", _regexp: "y");
                            test.perfaction("iframe", "name", "page", "page");
                            test.perfaction("dropdown", "name", "devcustfield201", "Weatherford");
                            #endregion Selenium_actions_on_UI
                            #region clicksubmitbutton
                            //        Console.ReadLine();
                            if (submityesno == "no")
                            {
                                return;
                            }
                            test.perfaction("img", "src", "/PTWeb/images/button_submit.gif", "submit");
                            Thread.Sleep(2000);
                            #endregion
                            #region getdevtrackTicketId
                            test.goToFirstFrame();
                            test.perfaction("frame", "name", "spider2", "spider2");
                            test.perfaction("frame", "name", "buglist2", "buglist2");
                            string newval = test.getdata("dropdown", "name", "buglist", "dd");
                            Console.WriteLine("devtrack etxt: " + newval);
                            string newdt = newval.Substring(0, newval.IndexOf(":"));
                            Console.WriteLine("devtrack id: " + newdt);
                            #endregion
                            #region updatequeryfire
                            SqlCommand cmdUpdateDevtrack = new SqlCommand();
                            cmdUpdateDevtrack.Connection = con;
                            cmdUpdateDevtrack.CommandType = CommandType.StoredProcedure;
                            cmdUpdateDevtrack.CommandText = "MPCQA_Custom_Update";

                            cmdUpdateDevtrack.Parameters.Add(new SqlParameter("@IssueID", dt.Rows[i]["IssueId"].ToString()));
                            cmdUpdateDevtrack.Parameters.Add(new SqlParameter("@CustomFieldID", devtrackID));
                            cmdUpdateDevtrack.Parameters.Add(new SqlParameter("@CustomDevtrackValue", newdt));
                            cmdUpdateDevtrack.Parameters.Add(new SqlParameter("@CloseStatusID", closeStatusID));
                            cmdUpdateDevtrack.ExecuteNonQuery();
                            cmdUpdateDevtrack.Parameters.Clear();
                            #endregion
                            #region assignSave
                            test.goToFirstFrame();
                            test.perfaction("frame", "name", "spider2", "spider2");
                            test.perfaction("frame", "name", "buglist1", "buglist1");
                            test.perfaction("text", "name", "srchkeyword", newdt);
                            test.perfaction("Link", "onClick", "return gotomode()", "Go");
                            Thread.Sleep(5000);
                            test.goToFirstFrame();
                            test.perfaction("frame", "name", "spider2", "spider2");
                            test.perfaction("frame", "name", "buglist2", "buglist2");
                            test.perfaction("iframe", "name", "IssueList", "IssueList");
                            test.perfaction("img", "src", "/PTWeb/images/icon_newissue.gif", "openbug");
                            Thread.Sleep(5000);
                            test.goToFirstFrame();
                            test.perfaction("frame", "name", "spider2", "spider2");
                            test.perfaction("frame", "name", "buglist2", "buglist2");
                            test.perfaction("iframe", "name", "buginfo2", "buginfo2");
                            Console.WriteLine("Got frame: Buginfo2");
                            test.perfaction("tabledata", "Linktext", "^Current Status$", "Current Status", _regexp: "y");
                            Console.WriteLine("Clicked Current Status Tab");
                            Thread.Sleep(1000);
                            test.perfaction("iframe", "name", "page", "page");
                            Console.WriteLine("Navigated to frame page");
                            Thread.Sleep(1000);
                            //save
                            test.perfaction("dropdown", "name", "devcustfield108", tableMapping.Rows[ij]["AssignedUser"].ToString());
                            Console.WriteLine("Selected assigned user value");
                            // test.perfaction("button", "onClick", "if(buginfo2.page && buginfo2.page.dosubmit) buginfo2.page.dosubmit(); else if(buginfo2.dosubmit) buginfo2.dosubmit();", "save");
                            test.goToFirstFrame();
                            test.perfaction("frame", "name", "spider2", "spider2");
                            test.perfaction("frame", "name", "buglist2", "buglist2");
                            test.perfaction("special_button", "value", "Save", "Save", 1);
                            Thread.Sleep(5000);
                            Console.WriteLine("Savd the assigned user");
                            #endregion Assign Save
                            Thread.Sleep(2000);
                            SendKeys.Flush();
                            SendKeys.SendWait("%{f4}");
                            SendEmails(tableMapping.Rows[ij]["AssignedUserEmailAdress"].ToString(), newdt, dt.Rows[i]["IssueID"].ToString());
                            Console.WriteLine("Done!!!:Email has been sent !!!");

                            #region CloseChildBrowserWindow - Commented Code replaced by ALT+F4 key stroke
                            //AutomationElement ae = AutomationElement.RootElement;
                            //AutomationElement childwindow = null;
                            //AutomationElementCollection allwindows = ae.FindAll(TreeScope.Children,
                            //           new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
                            //for (int iw = 0; iw < allwindows.Count; iw++)
                            //{
                            //    if (allwindows[iw].Current.Name == "Time Track Triggering - Google Chrome")
                            //    {
                            //        childwindow = allwindows[iw];
                            //        break;
                            //    }
                            //}
                            //WindowPattern winptn = (WindowPattern)childwindow.GetCurrentPattern(WindowPattern.Pattern);
                            //winptn.Close();

                            //Console.WriteLine("Done:Window is closed!!!");
                            //Thread.Sleep(5000);
                            //Console.WriteLine("Done:Window is closed!!!--Waiting for Devtrack to Breathe");
                            //#endregion
                            #endregion
                            //  Console.ReadLine();
                        } //end of Tickets count fo loop
                        #endregion Iterate through ticket
                        //#endregion
                    } //end for If condition to be checked in mapping excel


                }//end of Mapping Excel datatable rows 
                #endregion  IterateThroughMappingFile
            }
            finally
            {
                con.Close();
                con.Dispose();
            }
        }
        static string AddAttachments(string strIssueID)
        {
            string addAttachmentspath = "";
            SqlConnection con2 = new System.Data.SqlClient.SqlConnection();
            DataTable dtat = new DataTable();
            try
            {
                #region ConnectionForAttachments
                string server = "";
                Console.WriteLine("Trying to open connections string ");
                server = ConfigurationManager.AppSettings["server"];
                con2.ConnectionString = ConfigurationManager.ConnectionStrings["con"].ConnectionString;
                con2.Open();
                Console.WriteLine(con2.State);
                SqlCommand commandat = new SqlCommand();
                commandat.Connection = con2;
                commandat.CommandType = CommandType.StoredProcedure;
                commandat.CommandText = "MPCQA_Custom_GetDataForAttachment";
                commandat.Parameters.Add(new SqlParameter("@IssueID", strIssueID));
                #endregion
                SqlDataAdapter adapter2 = new SqlDataAdapter(commandat);
                adapter2.Fill(dtat);
                Console.WriteLine("Total number of attachments " + dtat.Rows.Count.ToString());
                DataTable dt4 = new DataTable();
                for (int iat = 0; iat < dtat.Rows.Count; iat++)
                {
                    byte[] byteArrayStoredImage = (byte[])dtat.Rows[iat]["Attachment"];
                    string issueattachid = dtat.Rows[iat]["IssueAttachmentID"].ToString();
                    #region DetermineAttachmenttype
                    string extname = dtat.Rows[iat]["FileName"].ToString();

                    // extname = extname.Substring(extname.LastIndexOf(".") + 1);
                    //mextname= extname.LastIndexOf(".")
                    string issueFolderPath = @"C:\Bugnet\" + strIssueID;
                    Console.WriteLine("Directory path : " + issueFolderPath);
                    if (System.IO.Directory.Exists(issueFolderPath) == false)
                    {
                        System.IO.Directory.CreateDirectory(issueFolderPath);
                    }
                    FileStream fs = new FileStream(issueFolderPath + @"\" + extname, FileMode.Create);
                    fs.Write(byteArrayStoredImage, 0, System.Convert.ToInt32(byteArrayStoredImage.Length));
                    fs.Seek(0, SeekOrigin.Begin);
                    fs.Close();
                    Console.WriteLine("Done for Extension Attachment: ." + dtat.Rows[iat]["ContentType"].ToString());

                    addAttachmentspath = addAttachmentspath + issueFolderPath + @"\" + extname + ";";
                    Console.WriteLine("the  reaturn path  ." + addAttachmentspath);


                    //        }

                    //  default : 
                    //        {
                    //        Console.WriteLine("Unable to determinde Attachment Type .Please add " + dtat.Rows[iat]["ContentType"].ToString() + "To your code select case !!" );
                    //        break;
                    //        }

                    //}
                    #endregion
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Execption shown " + ex.Message.ToString());
            }
            finally
            {
                con2.Close();
                con2.Dispose();
            }


            return addAttachmentspath;
        }
        static int GetIssueAttachmentsCount(string strIssueID)
        {
            SqlConnection con2 = new System.Data.SqlClient.SqlConnection();
            try
            {

                DataTable dtat = new DataTable();

                #region ConnectionForAttachments
                string server = "";
                Console.WriteLine("Trying to open connections string ");
                server = ConfigurationManager.AppSettings["server"];
                con2.ConnectionString = ConfigurationManager.ConnectionStrings["con"].ConnectionString;
                //  con2.ConnectionString = @"data source=" + server + ";database=Bugnet;uid=sa;pwd=test@123";
                con2.Open();
                Console.WriteLine(con2.State);
                SqlCommand commandat = new SqlCommand();
                commandat.Connection = con2;
                commandat.CommandType = CommandType.StoredProcedure;
                commandat.CommandText = "MPCQA_Custom_GetDataForAttachment";
                commandat.Parameters.Add(new SqlParameter("@IssueID", strIssueID));
                #endregion
                SqlDataAdapter adapter2 = new SqlDataAdapter(commandat);
                adapter2.Fill(dtat);
                Console.WriteLine("Total number of attachments " + dtat.Rows.Count.ToString());
                return dtat.Rows.Count;
            }
            finally
            {
                con2.Close();
                con2.Dispose();
            }
        }
        static DataTable CreateDataTableFromExcel(string excelmapperFile, string queryexec)
        {
            OdbcConnection conn2 = new OdbcConnection();
            try
            {
                DataTable dt2 = new DataTable();

                conn2.ConnectionString = @"Driver={Microsoft Excel Driver (*.xls)};DriverId=790;ReadOnly=0;Dbq=" + excelmapperFile;
                conn2.Open();
                string strcmdText = queryexec;
                OdbcCommand cmd = new OdbcCommand(strcmdText);
                cmd.Connection = conn2;
                //OdbcDataReader reder = cmd.ExecuteReader();
                OdbcDataAdapter da = new OdbcDataAdapter(cmd);
                da.Fill(dt2);
                return dt2;
            }
            finally
            {
                conn2.Close();
                conn2.Dispose();
            }
        }
        public static string MakeValidFileName(string name)
        {
            var builder = new StringBuilder();
            var invalid = System.IO.Path.GetInvalidFileNameChars();
            foreach (var cur in name)
            {
                if (!invalid.Contains(cur))
                {
                    builder.Append(cur);
                }
            }
            return builder.ToString();
        }
        public static void SendEmails(string strTo, string strDevtracktk, string strBugnetTkt)
        {
            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            string[] recipients = strTo.Split(';');
            foreach (string recipient in recipients)
            {
                message.To.Add(recipient);
            }
            message.CC.Add("Deepankar.Bandopadhyay@me.weatherford.com");
            message.Subject = "DevTrack Ticket # " + strDevtracktk + " created for Bugnet: " + strBugnetTkt;
            message.From = new System.Net.Mail.MailAddress("noreply@bugnet-vm1.com");
            message.Body = "This is to inform you that a Devtrack Ticket has been Created for Bugnet  " + strBugnetTkt + ". Devtrack Id Is: " + strDevtracktk;
            System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("mail2.weatherford.com");
            smtp.Port = 25;

            smtp.Send(message);

            Console.WriteLine("Done!!!:Exit from Email function !!!");
        }

        public static string ReplaceHTMLTags(string issueDesc)
        {
            Console.WriteLine("Inside ReplaceHTMLTags");

            var newValue = System.Text.RegularExpressions.Regex.Replace(issueDesc, "<(.|\n)*?>", "");
            return newValue.Replace("  ", "");

        }
    }
}
