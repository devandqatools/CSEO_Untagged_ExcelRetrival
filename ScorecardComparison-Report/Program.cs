using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using DataTable = System.Data.DataTable;

namespace ScorecardComparison_Report
{
    //Authors : Chidvilash Vakada, Krishna Reddy
    //Date Written : 18th May 2020
    /// <summary>
    /// Automatic Email for Untagged MAS bugs.
    /// </summary>
    class Program
    {
        public static double xCount = 0;

        public static Dictionary<string, int> diffSheerCount = new Dictionary<string, int>();
        public static List<string> datesInFilePathList = new List<string>();
        public static string ApplicationNameWiseData_Untagged = "ApplicationNameWiseData_Untagged";
        static void Main(string[] args)
        {
            string applicationType = ConfigurationManager.AppSettings["ApplicationType"];
            string excelPath = ConfigurationManager.AppSettings["ExcelLocPath"];
            Console.WriteLine(" Results Data table added to Data Table .....!");
            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", excelPath);
            OleDbConnection Econ = new OleDbConnection(constr);
            string Query = string.Format("select RecID,AppName from [{0}] where AppName <> null and RecID <> null group by RecID,AppName", "UntaggedBugs$");
            OleDbCommand Ecom = new OleDbCommand(Query, Econ);
            Econ.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
            Econ.Close();
            oda.Fill(ds);
            DataTable table = ds.Tables[0];
            Console.WriteLine("Group Application total Count ::" + table.Rows.Count);
            List<string> finalDataList = new List<string>();
            if (table.Rows != null && table.Rows.Count > 0)
            {
                foreach (DataRow dr in table.Rows)
                {
                    try
                    {
                        string recordID = dr["RecId"].ToString();
                        string appName = dr["AppName"].ToString();
                        int recID = Int32.Parse(recordID);
                        string Query2 = string.Format("Select ID,RecID,AppName,Severity,Bug_Identified,Tagging,Bug_Aging,Created_By,Created_Date,Area_Path,Iteration_Path,Link,AssignedTo_Email,AssignedTo_Name,vdash,AIRT_LatestLead_Alies,PMOwner,EngOnwer,SubGrp,Priority,OpsStatus from [{0}] where AppName='" + appName + "' and RecID="+recID, "UntaggedBugs$");
                        OleDbConnection Econ2 = new OleDbConnection(constr);
                        OleDbCommand Ecom2 = new OleDbCommand(Query2, Econ2);
                        Econ2.Open();
                        DataSet ds2 = new DataSet();
                        OleDbDataAdapter oda2 = new OleDbDataAdapter(Query2, Econ2);
                        Econ2.Close();
                        oda2.Fill(ds2);
                        DataTable table2 = ds2.Tables[0];
                        Console.WriteLine("ApplicationWise Row Count ::" + table2.Rows.Count);
                        Dictionary<String, List<String>> dataMapWithMails = SendMasBugsDataToIndividaulByTitle(applicationType, table2, recordID, appName);
                        sendMail(dataMapWithMails, appName, recordID);
                    }catch (Exception e)
                    {
                        Console.WriteLine("Error in fetching Excel Sheet ::: "+e.InnerException);
                    }
                }
            }   
        }

        /**
         * 
         * This method is used to fetch data from DataTable and pushing into required Dictionary format.
         * 
         * */
        public static Dictionary<String, List<String>> SendMasBugsDataToIndividaulByTitle(String applicationType,DataTable table, String recID,String currentAppName)
        {
            Dictionary<String, List<String>> applicationData = new Dictionary<String, List<String>>();
            
                        
                        if (table.Rows != null && table.Rows.Count > 0)
                        {
                            foreach (DataRow dr in table.Rows)
                            {
                                string id = dr["ID"].ToString();
                                string recordID = dr["RecID"].ToString();
                                string appName = dr["AppName"].ToString();
                                string severity = dr["Severity"].ToString();
                                string bugIdentified = dr["Bug_Identified"].ToString();
                                string tagging = dr["Tagging"].ToString();
                                string bugAging = dr["Bug_Aging"].ToString();
                                string createdBy = dr["Created_By"].ToString();
                                string createdDate = dr["Created_Date"].ToString();
                                string areaPath = dr["Area_Path"].ToString();
                                string iterationPath = dr["Iteration_Path"].ToString();
                                string link = dr["Link"].ToString();
                                string assignedToEmail = dr["AssignedTo_Email"].ToString();
                                string assignedToName = dr["AssignedTo_Name"].ToString();
                                string vdash = dr["vdash"].ToString();
                                string aIRT_LatestLead_Alies = dr["AIRT_LatestLead_Alies"].ToString();
                                string pmOwner = dr["PMOwner"].ToString();
                                string engOwner = dr["EngOnwer"].ToString();
                                string subGrp = dr["SubGrp"].ToString();
                                string priority = dr["Priority"].ToString();
                                string opsStatus = dr["OpsStatus"].ToString();
                            
                            //Code for RecID wise mail data Integration
                                if (applicationData.Count == 0 || !applicationData.ContainsKey(recordID))
                                {
                                    List<String> applicationList = new List<String>();
                                    applicationList.Add(id + "#$%" + recordID + "#$%" + appName + "#$%" + severity + "#$%" + bugIdentified + "#$%" + tagging + "#$%" + bugAging + "#$%" + createdBy + "#$%" + createdDate + "#$%" + areaPath + "#$%" + iterationPath + "#$%" + link + "#$%" + assignedToEmail + "#$%" + assignedToName + "#$%" + vdash + "#$%" + aIRT_LatestLead_Alies + "#$%" + pmOwner + "#$%" + engOwner + "#$%" + subGrp + "#$%" + priority + "#$%" + opsStatus);
                                    applicationData.Add(recordID, applicationList);
                                }
                                else if (applicationData.ContainsKey(recordID))
                                {
                                    List<String> newList = applicationData[recordID];
                                    newList.Add(id + "#$%" + recordID + "#$%" + appName + "#$%" + severity + "#$%" + bugIdentified + "#$%" + tagging + "#$%" + bugAging + "#$%" + createdBy + "#$%" + createdDate + "#$%" + areaPath + "#$%" + iterationPath + "#$%" + link + "#$%" + assignedToEmail + "#$%" + assignedToName + "#$%" + vdash + "#$%" + aIRT_LatestLead_Alies + "#$%" + pmOwner + "#$%" + engOwner + "#$%" + subGrp + "#$%" + priority + "#$%" + opsStatus);
                                    applicationData.Remove(recordID);
                                    applicationData.Add(recordID, newList);
                                }

                            }
                        }
            return applicationData;
        }

        /*
         *  This method is used to sendMails to respective mail recievers
         * 
         * */
        private static void sendMail(Dictionary<String, List<String>> dataMapWithMails, String currentAppName, String recID)
        {            
            List<String> emailIdList = dataMapWithMails.Keys.ToList();
            foreach(String emailId in emailIdList) {
                StringBuilder builder = new StringBuilder();
                StringBuilder mainBuiler = new StringBuilder();
                List<String> emailIdData = dataMapWithMails[emailId];
                if (emailIdData != null && emailIdData.Count > 0)
                {
                    builder.Append("<span style=\"font-size:11.0pt;mso-fareast-font-family:&quot;Times New Roman&quot;\"><br>We are seeing Untagged Bugs( <b>" + emailIdData.Count + "</b> ) assigned to you for the application (details below) and we need your immediate attention here!</span><br><br>");
                    builder.Append("<span style=\"font-size:11.0pt;mso-fareast-font-family:&quot;Times New Roman&quot;\"><b>Details of the application : </b></span>");
                    builder.Append("<table class=\"imagetable\">");
                    String row = emailIdData[0];
                    string[] dataArray = row.Split(new String[] { "#$%" }, StringSplitOptions.None);
                    builder.Append("<tr bgcolor='#0F2BDE'>");
                    if (currentAppName != null && !"".Equals(currentAppName))
                        builder.Append("<th style=\"width:30%;\"><font color=\"#FFFFFF\">Application Name</font></th>");
                    if (dataArray[16] != null && !"".Equals(dataArray[16]))
                        builder.Append("<th style=\"width:15%;\"><font color=\"#FFFFFF\">PM Owner</font></th>");
                    if (dataArray[17] != null && !"".Equals(dataArray[17]))
                        builder.Append("<th style=\"width:10%;\"><font color=\"#FFFFFF\">Engineering Owner</font></th>");
                    if (dataArray[18] != null && !"".Equals(dataArray[18]))
                        builder.Append("<th style=\"width:30%;\"><font color=\"#FFFFFF\">Sub Group</font></th>");
                    if (dataArray[19] != null && !"".Equals(dataArray[19]))
                        builder.Append("<th style=\"width:10%;\"><font color=\"#FFFFFF\">Priority</font></th>");
                    if (dataArray[20] != null && !"".Equals(dataArray[20]))
                        builder.Append("<th style=\"width:15%;\"><font color=\"#FFFFFF\">Application Status</font></th>");
                    builder.Append("</tr>");

                    builder.Append("<tr>");
                    if (currentAppName != null && !"".Equals(currentAppName))
                        builder.Append("<td>" + currentAppName + "</td>");
                    if (dataArray[16] != null && !"".Equals(dataArray[16]))
                        builder.Append("<td>" + dataArray[16] + "</td>");
                    if (dataArray[17] != null && !"".Equals(dataArray[17]))
                        builder.Append("<td>" + dataArray[17] + "</td>");
                    if (dataArray[18] != null && !"".Equals(dataArray[18]))
                        builder.Append("<td>" + dataArray[18] + "</td>");
                    if (dataArray[19] != null && !"".Equals(dataArray[19]))
                        builder.Append("<td>" + dataArray[19] + "</td>");
                    if (dataArray[20] != null && !"".Equals(dataArray[20]))
                        builder.Append("<td>" + dataArray[20] + "</td>");
                    builder.Append("</tr>");
                    builder.Append("</table><br>");
                }
                else
                {
                    builder.Append("<span style=\"font-size:11.0pt;mso-fareast-font-family:&quot;Times New Roman&quot;\"><br>We are seeing Untagged Bugs( " + emailIdData.Count + " ) assigned to you");
                    builder.Append("<br>");
                }

                builder.Append("<span style=\"font-size:11.0pt;mso-fareast-font-family:&quot;Times New Roman&quot;\">It's important to tag our bugs related to Accessibility so we can understand the Accessibility landscape and help us identify and engage with engineering teams in remediation. Refer: <a href='https://apc01.safelinks.protection.outlook.com/?url=https%3A%2F%2Faka.ms%2Fbugtagging&amp;data=02%7C01%7Cvenkatakrishn.sabbe%40hcl.com%7Cfe5ee50a7c2d4e83974008d74664b122%7C189de737c93a4f5a8b686f4ca9941912%7C0%7C0%7C637055270587327118&amp;sdata=olS9Mm41kdLsE0%2FC%2FNwrFWzQzYVEq1Ga63Wd6DthvCY%3D&amp;reserved=0' target='_blank' rel='noopener noreferrer' data-auth='Verified' originalsrc='https://aka.ms/bugtagging' shash='j6HH9xt0jdUu8gNTsiAVydHiuh4mOU68p45bGpyCl9nie/b67pf6RuFDY5aT/ZothN/zcJvHC8k/1D89Ryj6S062+rZZAri6ooCxh0bErAQN89+zlYPwy8w/4Qj/hyVBVSnIhsdXYiUPelOT8TWwEHMVsQDkOiaDLZRhHxzw7h0=' title='Original URL: https://aka.ms/bugtagging" +
                "Click or tap if you trust this link.'>https://aka.ms/bugtagging</a>  (link to overall CSEO Accessibility Tagging Guidelines). Below are also those tags for quick reference. Feel free to touch base with your groups <a href='https://apc01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fmicrosoft.sharepoint.com%2Fteams%2Fmeconnection%2FSitePages%2FAccessibility-Contacts.aspx&amp;data=02%7C01%7Cvenkatakrishn.sabbe%40hcl.com%7Cfe5ee50a7c2d4e83974008d74664b122%7C189de737c93a4f5a8b686f4ca9941912%7C0%7C0%7C637055270587337112&amp;sdata=y5BDL052uj%2BQlgAweOG2E16vNEkpU6xM9fsldqE8bsI%3D&amp;reserved=0' target='_blank' rel='noopener noreferrer' data-auth='Verified' originalsrc='https://microsoft.sharepoint.com/teams/meconnection/SitePages/Accessibility-Contacts.aspx' shash='miBHLxMGT4zD1hmQe2pnph0JhDM2V50SIjFk1wSyGjgwAZ+FY67PRZ4rSxTRyTmKVe6Subu5sWmsjrLUu76cWvRTLtj/dOCoknas5ZX5Tx1KbaYMqm793+9p6UjQj/svaEsADeoZqQ6Hud4HTcURoA1RWomDtvKVmYJLEnzo7cs=' title='Original URL: https://microsoft.sharepoint.com/teams/meconnection/SitePages/Accessibility-Contacts.aspx" +
                "Click or tap if you trust this link.'>Accessibility Track Lead </a> or us for any queries or concerns and we would help you here.<br><br></span>");
                builder.Append("<table class='imagetable'>" +
                "<tr bgcolor='#0F2BDE'>" +
                "<th style=\"width: 15%;\"><font color=\"#FFFFFF\">Tag</font></th>" +
                "<th style=\"width: 20%;\"><font color=\"#FFFFFF\">Description</font></th>" +
                "<th style=\"width: 30%;\"><font color=\"#FFFFFF\">Example/Use</font></th>" +
                "</tr>" +
                "<tr>" +
                "<td><span>A11y-COREDEV</span></td>" +
                "<td><span>For any Core Development Team bug</span></td>" +
                "<td>&nbsp;</td>" +
                "</tr>" +
                "<tr>" +
                "<td><span>A11y-1STPARTY</span></td>" +
                "<td><span>For any 1</span><sup>st</sup><span> party bug (MS Product Group)</span></td>" +
                "<td><span>MS Product Groups include: SharePoint, Dynamics, Teams</span></td>" +
                "</tr>" +
                "<tr>" +
                "<td>" +
                "<span>FirstParty_&lt;DependencyName&gt;</span>" +
                "</td>" +
                "<td>Additional 1st party tag, naming the actual 1st party</td>" +
                "<td>" +
                "<span>FirstParty_SPO (for Share Point Online)</span>" +
                "<span>FirstParty_PowerBI (for Power BI)</span>" +
                "<span>FirstParty_PowerApps (for Power apps)</span>" +
                "<span>For a complete list, consult the <a href='https://microsoft.sharepoint.com/teams/meconnection/SitePages/First-Party-Points-of-Contact.aspx' data-interception='on'>full list of 1st party products</a></span>" +
                "</td>" +
                "</tr>" +
                "<tr>" +
                "<td><span>A11y-3RDPARTY</span></td>" +
                "<td>For any 3rd party bug</td>" +
                "<td><span>3</span><sup>rd</sup><span> party examples include: SAP, Fidelity, YouTube</span></td>" +
                "</tr>"+
                "<tr>"+
                "<td>"+
                "<span>ThirdParty_&lt;DependencyName&gt;</span>" +
                "</td>" +
                "<td>Additional 3rd party tag, naming the actual 3rd party</td>" +
                "<td>" +
                "<span>ThirdParty_SAP (for SAP)</span>" +
                "<span>ThirdParty_Adobe (for Adobe)</span>" +
                "<span>ThirdParty_ServiceNow (for Service Now)</span>" +
                "</td>" +
                "</tr>" +
                "</tbody>" +
                "</table>");
                builder.Append("<span style=\"mso-fareast-font-family:&quot;Times New Roman&quot;\"><u5:p></u5:p><o:p></o:p></span><br>");
                builder.Append("<u5:p></u5:p>");
                builder.Append("<body>");
                List<String> leadMailDetails = new List<String>();
                List<String> mailToDetails = new List<String>();
                if (emailIdData.Count != 0)
                {
                    builder.Append("<br>");
                    builder.Append("<span style=\"font-size:11.0pt;mso-fareast-font-family:&quot;Times New Roman&quot;\"><b>List of All Untagged Bugs Assigned to you : </b></span>");
                    builder.Append("<table class=\"imagetable\">");
                    builder.Append("<tr bgcolor=\"#0F2BDE\">");
                    builder.Append("<th><font color=\"#FFFFFF\">Bug Id</font></th>" +
                    "<th><font color=\"#FFFFFF\">Severity</font></th>" +
                    "<th><font color=\"#FFFFFF\">Bug Aging</font></th>" +
                    "<th><font color=\"#FFFFFF\">Created By</font></th>" +
                    "<th><font color=\"#FFFFFF\">Created Date</font></th>" +
                    "<th><font color=\"#FFFFFF\">Area Path</font></th>" +
                    "<th><font color=\"#FFFFFF\">Assigned To</font></th>"
                    );
                    builder.Append("</tr>");
                    foreach (String appData in emailIdData)
                    {
                        string[] appDataArray = appData.Split(new String[] { "#$%" }, StringSplitOptions.None);
                        builder.Append("<tr>");
                        builder.Append("<td><a href='https://microsoftit.visualstudio.com/DefaultCollection/OneITVSO/_workitems/edit/" + appDataArray[0] + "'>" + appDataArray[0] + "</a></td>");
                        builder.Append("<td>" + appDataArray[3] + "</td>");
                        builder.Append("<td>" + appDataArray[6] + "</td>");
                        builder.Append("<td>" + appDataArray[7] + "</td>");
                        builder.Append("<td>" + appDataArray[8] + "</td>");
                        builder.Append("<td>" + appDataArray[9] + "</td>");
                        builder.Append("<td>" + appDataArray[12] + "</td>");
                        leadMailDetails.Add(appDataArray[15]);
                        mailToDetails.Add(appDataArray[12]);
                        builder.Append("</tr>");
                    }
                    builder.Append("</table>");
                }
                builder.Append("<br>");
                builder.Append("<p style=\"color:blue;\">Thanks,<br>CSEO Assessment Service </p>");
                builder.Append("</body>");

                StringBuilder builder2 = new StringBuilder();
                builder2.Append("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 3.2//EN\">");
                builder2.Append("<html xmlns='http://www.w3.org/1999/xhtml'>");
                builder2.Append("<head>");
                builder2.Append("<title>");
                builder2.Append("Need Attention: RecID: " + recID + " " + currentAppName + "All Untagged Bugs Report :" + System.DateTime.Now.ToShortDateString() + " " + System.DateTime.Now.ToShortTimeString());
                builder2.Append("</title>");
                builder2.Append("<style type=\"text/css\">");
                builder2.Append("table.MsoNormalTable{font-size:10.0pt;font-family:\"Calibri\",serif;}" +
                    "p.MsoNormal{margin-bottom:.0001pt;font-size:12.0pt;font-family:\"Calibri\",serif;margin-left: 0in;margin-right: 0in;margin-top: 0in;}" +
                    "h1{margin-right:0in;margin-left:0in;font-size:24.0pt;font-family:\"Calibri\",serif;font-weight:bold;}" +
                    "a:link{color:#0563C1;text-decoration:underline;text-underline:single;}p{margin-right:0in;margin-left:0in;font-size:12.0pt;font-family:\"Calibri\",serif;}");
                builder2.Append("table.imagetable{font-family: verdana,Calibri,sans-serif;font-size:11px;" +
                    "color:#333333;border-width: 1px;border-color: #999999;border-collapse: collapse;}");
                builder2.Append("table.imagetable th {background:#0070C0 url('cell-blue.jpg');border-width: 1px;padding: 8px;border-style: solid;border-color: #999999;}");
                builder2.Append("table.imagetable td {background:#FFFFFF url('cell-grey.jpg');border-width: 1px;" +
                    "padding: 8px;border-style: solid;border-color: #999999;width:1px;white-space:nowrap;}</style>");
                builder2.Append("</head>");
                List<String> distinctLeadMails = leadMailDetails.Distinct().ToList();
                List<String> distinctMailToDetails = mailToDetails.Distinct().ToList();
                String leadMailString = String.Join(";", distinctLeadMails);
                String mailToString = String.Join(";", distinctMailToDetails);
                //String[] emailArray =  emailId.Split(':');
                builder2.Append("<span style=\"font-size:11.0pt;mso-fareast-font-family:&quot;Times New Roman&quot;\">Hi Team,</span><br>");
                mainBuiler.Append(builder2);
                mainBuiler.Append(builder);
                string HtmlFile = mainBuiler.ToString();
                Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                //Live
                //oMsg.To = mailToString;
                //QA
                oMsg.To = "v-chvak@microsoft.com;v-vesabb@microsoft.com;";
                oMsg.Subject = "Need Attention: RecID: " + recID+" - " + currentAppName + " - All Untagged Bugs Report :" + System.DateTime.Now.ToShortDateString() + " " + System.DateTime.Now.ToShortTimeString();
                oMsg.HTMLBody = HtmlFile;
                String localEmailList = ConfigurationManager.AppSettings["emailsList"].ToString();
                List<String> localEmailListArray = localEmailList.Split(',').ToList();
                bool containsValues = localEmailListArray.Intersect(distinctMailToDetails).Any();
                if (containsValues)
                {
                    //Live Only if QA need to comment
                    // oMsg.CC = ConfigurationManager.AppSettings["DftSendMailTo"];
                }
                else
                {
                    //Live Only if QA need to comment
                   // oMsg.CC = ConfigurationManager.AppSettings["clientMailsCC"].ToString()+ leadMailString + ";"+ ConfigurationManager.AppSettings["DftSendMailTo"].ToString();                   
                    //oMsg.CC = "v-vesabb@microsoft.com";
                }
                oMsg.Send();
            }
        }
    }
}