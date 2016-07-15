using System;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using System.DirectoryServices;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;
using System.ComponentModel;



// Utility class with methods for retrieving values from AD and sending notifications

namespace Mosaicco.UpdateTivoliFlag.Console
{
    public class MosaicUtility
    {
        public struct ADReportProperties
        {
            public string LegalHoldAttributeName;
            public string EmployeeId;
            public string EmailAddress;
            public string LegalHoldUserEmployeeStatus;
            public string LegalHoldUserTermedDate;

        }



        public void LogToFile(String msg)
        {
            Boolean blAppend;
            string FileName = ConfigurationManager.AppSettings["logFile"];
            try
            {
                if (File.Exists(FileName))
                {
                    blAppend = true;
                }
                else
                {
                    blAppend = false;
                }
                StreamWriter sw = new StreamWriter(FileName, blAppend);
                sw.WriteLine(msg);
                sw.Close();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
        }

        public ADReportProperties ReportProperties;

        public ADReportProperties GetLegalHoldReportProperties(string username)
        {
            const string legalHoldAttributeName = "extensionAttribute4";
            const string employeeId = "employeenumber";
            const string legalHoldUserEmployeeStatus = "employeestatus";
            const string emailAddress = "mail";
            const string legalHoldUserTermedDate = "termination_date";


            ADReportProperties ReportProperties = new ADReportProperties();
            try
            {
                DirectoryEntry myLdapConnection = createDirectoryEntry();

                DirectorySearcher search = new DirectorySearcher(myLdapConnection);
                search.Filter = "(displayName=" + username + ")";
                search.PropertiesToLoad.Add(legalHoldAttributeName);
                search.PropertiesToLoad.Add(employeeId);
                search.PropertiesToLoad.Add(legalHoldUserEmployeeStatus);
                search.PropertiesToLoad.Add(emailAddress);
                search.PropertiesToLoad.Add(legalHoldUserTermedDate);

                SearchResult result = search.FindOne();

                if (result != null)
                {
                    DirectoryEntry reportEntries = result.GetDirectoryEntry();
                    // retrieve attributes from AD     

                    if (reportEntries.Properties[employeeId].Value != null)
                        ReportProperties.EmployeeId = reportEntries.Properties[employeeId].Value.ToString();
                    else
                        ReportProperties.EmployeeId = "Null";

                    if (reportEntries.Properties[emailAddress].Value != null)
                        ReportProperties.EmailAddress = reportEntries.Properties[emailAddress].Value.ToString();
                    else
                        ReportProperties.EmailAddress = "Null";

                    if (reportEntries.Properties[legalHoldAttributeName].Value != null)
                        ReportProperties.LegalHoldAttributeName = reportEntries.Properties[legalHoldAttributeName].Value.ToString();
                    else
                        ReportProperties.LegalHoldAttributeName = "Null";

                    if (reportEntries.Properties[legalHoldUserEmployeeStatus].Value != null)
                        ReportProperties.LegalHoldUserEmployeeStatus = reportEntries.Properties[legalHoldUserEmployeeStatus].Value.ToString();
                    else
                        ReportProperties.LegalHoldUserEmployeeStatus = "Null";

                    if (reportEntries.Properties[legalHoldUserTermedDate].Value != null)
                        ReportProperties.LegalHoldUserTermedDate = reportEntries.Properties[legalHoldUserTermedDate].Value.ToString();
                    else
                        ReportProperties.LegalHoldUserTermedDate = "Null";

                    return ReportProperties;
                }

                else throw new Exception("user not found");
            }

            catch (Exception e)
            {
                throw new Exception(e.ToString());
            }
        }



        public bool UpdateLegalHoldStatus(string username, bool legalholdstatus)
        {
            const string LegalHoldAttributeName = "extensionAttribute4";
            try
            {
                DirectoryEntry myLdapConnection = createDirectoryEntry();

                DirectorySearcher search = new DirectorySearcher(myLdapConnection);
                search.Filter = "(displayName=" + username + ")";
                search.PropertiesToLoad.Add(LegalHoldAttributeName);

                SearchResult result = search.FindOne();

                if (result != null)
                {
                    DirectoryEntry entryToUpdate = result.GetDirectoryEntry();
                    // update legal hold status attribute in AD                    
                    entryToUpdate.Properties[LegalHoldAttributeName].Value = legalholdstatus;
                    entryToUpdate.CommitChanges();
                    return true;
                }

                else throw new Exception("user not found");
            }

            catch (Exception e)
            {
                throw new Exception(e.ToString());
            }
        }

        static DirectoryEntry createDirectoryEntry()
        {
            try
            {
                string LdapServer = System.Configuration.ConfigurationManager.AppSettings["LdapServer"];
                string user = ConfigurationManager.AppSettings["ImpersonationUser"];
                string password = ConfigurationManager.AppSettings["ImpersonationPassword"];
                string LdapConnectionString = System.Configuration.ConfigurationManager.AppSettings["LdapConnectionString"];
                DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://az-dc-adm1.corp.mosaicco.com:3268/DC=corp,DC=mosaicco,DC=com", user, password);

                return directoryEntry;
            }

            catch (Exception e)
            {
                throw new Exception(e.ToString());
            }

        }


        public bool SendEmailNotification(string emailaddress, string caseSiteUrl, bool legalholdstatus)
        {

            string SmtpHostName = System.Configuration.ConfigurationManager.AppSettings["SmtpHostName"];
            // Command line argument must the the SMTP host.
            SmtpClient client = new SmtpClient(SmtpHostName);
            // Specify the e-mail sender.
            // Create a mailing address that includes a UTF8 character
            // in the display name.
            MailAddress from = new MailAddress("ediscovery@mosaicco.com",
               "eDiscovery -" + (char)0xD8 + " Mosaic",
            System.Text.Encoding.UTF8);
            // Set destinations for the e-mail message.
            MailAddress to = new MailAddress(emailaddress);
            // Specify the message content.
            MailMessage message = new MailMessage(from, to);

            message.Body = "We are notifying you that you have been placed on Legal Hold. Please click this Case link to manage your legal hold status: " + caseSiteUrl;

            // Include some non-ASCII characters in body and subject.
            string someArrows = new string(new char[] { '\u2190', '\u2191', '\u2192', '\u2193' });
            message.Body += Environment.NewLine + someArrows;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.Subject = "test message 1" + someArrows;
            message.SubjectEncoding = System.Text.Encoding.UTF8;
            // Set the method that is called back when the send operation ends.
            client.UseDefaultCredentials = true;

            try
            {
                client.Send(message);
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
        }

    }
}