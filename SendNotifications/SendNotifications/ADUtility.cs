using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.DirectoryServices;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;
using System.ComponentModel;

namespace Mosaicco.LegalHoldReport.Console
{
    public struct ADReportProperties
    {
        public string LegalHoldAttributeName ;
        public string EmployeeId;
        public string LegalHoldUserEmployeeStatus;
        public string LegalHoldUserTermedDate;
        
    }
  
    public class ADUtility
    {

        public ADReportProperties ReportProperties;
        
        public ADReportProperties RetrieveReportPropertiesFromAD(string username)
        {
            const string LegalHoldAttributeName = "extensionAttribute4";
            const string EmployeeId = "employeeid";
            const string LegalHoldUserEmployeeStatus = "legalholduseremployeestatus";
            const string LegalHoldUserTermedDate = "legalholdusertermeddate";

            ADReportProperties ReportProperties = new ADReportProperties();
            try
            {
                DirectoryEntry myLdapConnection = createDirectoryEntry();

                DirectorySearcher search = new DirectorySearcher(myLdapConnection);
                search.Filter = "(cn=" + username + ")";
                search.PropertiesToLoad.Add(LegalHoldAttributeName);

                SearchResult result = search.FindOne();

                if (result != null)
                {
                    DirectoryEntry reportEntries = result.GetDirectoryEntry();
                    // update legal hold status attribute in AD                    
                    ReportProperties.EmployeeId =  reportEntries.Properties[EmployeeId].Value.ToString();
                    ReportProperties.LegalHoldAttributeName = reportEntries.Properties[LegalHoldAttributeName].Value.ToString();
                    ReportProperties.LegalHoldUserEmployeeStatus = reportEntries.Properties[LegalHoldUserEmployeeStatus].Value.ToString();
                    ReportProperties.LegalHoldUserTermedDate = reportEntries.Properties[LegalHoldUserTermedDate].Value.ToString();
                    
                    return ReportProperties;
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
            string LdapServer = System.Configuration.ConfigurationManager.AppSettings["LdapServer"];
            string LdapConnectionString = System.Configuration.ConfigurationManager.AppSettings["LdapConnectionString"];

            DirectoryEntry ldapConnection = new DirectoryEntry(LdapServer); //"ldap://az-dc-adm1.corp.mosaicco.com"
            ldapConnection.Path = LdapConnectionString; //ldap://az-dc-adm1.corp.mosaicco.com:3268/DC=corp,DC=mosaicco,DC=com";
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
            return ldapConnection;
        }

        public bool SendEmailNotification(string emailaddress, bool legalholdstatus)
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

            message.Body = "We are notifying you that you have been placed on Legal Hold. ";

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
