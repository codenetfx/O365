using System;
using System.Net;
using System.Collections;
using System.Collections.Specialized;
using System.Configuration;
using System.Collections.Generic;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Xml.Linq;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Mosaicco.LegalHoldReport.Console.SPListWS;
using Mosaicco.LegalHoldReport.Console.SPSiteWS;
using Mosaicco;

//logon impersonation
using System.Security;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices; // DllImport
using System.Runtime.ConstrainedExecution;
using System.Security.Principal; // WindowsImpersonationContext
using System.Security.Permissions; // PermissionSetAttribute

namespace Mosaicco.LegalHoldReport.Console {
    class Program {

        // obtains user token
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool RevertToSelf();

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LogonUser(string pszUsername, string pszDomain, string pszPassword,
            int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

        // closes open handes returned by LogonUser
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public extern static bool CloseHandle(IntPtr handle);


        static void Main(string[] args) {
            if (args.Count() != 3) {
                //System.Console.WriteLine("Syntax: UpdateTivoliFlag.exe url username password");
            }

            //Authenticate against the SharePoint Portal using the portal user account
            MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper("https://mosaicco.sharepoint.com/sites/ediscvrcntr/HOLD-044", "samir.dobric@mosaicco.com", "Auftvrd902");
            //MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper("https://mosaicco.sharepoint.com/sites/ediscvrcntr/HOLD-044", "legalholdsvc@mosaicco.com", "Mosaic123456");
            HttpRequestMessageProperty p = new HttpRequestMessageProperty();
            var cookie = claimsHelper.CookieContainer;
            string cookieHeader = cookie.GetCookieHeader(new Uri("https://mosaicco.sharepoint.com/sites/edicvrcntr/HOLD-044"));
            p.Headers.Add("Cookie", cookieHeader);
            
            var WebCollection = new List<Tuple<string, string>>();
            using (SPWebsWS.WebsSoapClient proxy = new SPWebsWS.WebsSoapClient())
            {
                proxy.Endpoint.Address = new EndpointAddress("https://mosaicco.sharepoint.com/sites/ediscvrcntr/_vti_bin/Webs.asmx");

                using (new OperationContextScope(proxy.InnerChannel))
                    {
                            OperationContext.Current.OutgoingMessageProperties[HttpRequestMessageProperty.Name] = p;
                            //Get list of all subsites (case subsites)
                            XElement subWebsItems = proxy.GetWebCollection();
                            
                            //parse the response message
                            System.Xml.XmlDocument xd = new System.Xml.XmlDocument();
                            xd.LoadXml(subWebsItems.ToString());
                            System.Xml.XmlNamespaceManager nm = new System.Xml.XmlNamespaceManager(xd.NameTable);
                            nm.AddNamespace("rootNS", "http://schemas.microsoft.com/sharepoint/soap/");
                            System.Xml.XmlNodeList nl = xd.SelectNodes("/rootNS:Webs/Web", nm);

                           
                            foreach (System.Xml.XmlNode websnode in xd)
                            {
                                foreach (System.Xml.XmlNode webnode in websnode)
                                {
  
                                    WebCollection.Add(new Tuple<string, string>(webnode.Attributes["Url"].Value.ToString(), webnode.Attributes["Title"].Value.ToString()));
                                    
                                }
                            }
                      
                    }
               
            }
            
            //Build content for the Legal Hold report file

            var csvExportFileContent = new StringBuilder();
            var newCSVHeaderLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}", "Case Name", "Case Number", "Case Status", "Attorney Lead For Case", "Legal Hold User Name", "EmployeeID", "Legal Hold User Employee Status", "Legal Hold User Termed Date", "Email Address/UserId", "In-Place Hold", "Date Modified");
            csvExportFileContent.AppendLine(newCSVHeaderLine);

            MosaicUtility mosaicUtility = new MosaicUtility();
                    

            foreach (Tuple<string, string>WebItem in WebCollection)
            using (ListsSoapClient proxyList = new ListsSoapClient())
            {
                //This endpoint must be site-specific
                //proxyList.Endpoint.Address = new EndpointAddress("https://mosaicco.sharepoint.com/sites/ediscvrcntr/HOLD-044/_vti_bin/Lists.asmx");
                proxyList.Endpoint.Address = new EndpointAddress(WebItem.Item1 + "/_vti_bin/Lists.asmx");

                string caseSiteUrl = WebItem.Item1;
                string caseName = WebItem.Item2;
                string caseNumber = string.Empty;
                string attorneyLeadForCase = string.Empty;

                

                System.Console.WriteLine("==================================================================");
                System.Console.WriteLine(" Processing Case Site: " + WebItem.Item1);
                

                
                System.Console.WriteLine("Processing Case Name: " + caseName);
                System.Console.WriteLine("Case Site Url: " + caseSiteUrl);

                using (new OperationContextScope(proxyList.InnerChannel))
                {
                    // Retrieve list items from the Sources list which contains legal hold data
                    OperationContext.Current.OutgoingMessageProperties[HttpRequestMessageProperty.Name] = p;

                    XElement nodeListItems;

                    try
                    {                     
                        nodeListItems = proxyList.GetListItems("Sources", null, null, null, "65535", null, null);
                    }
                    catch (Exception ex)
                    {
                        nodeListItems = null;
                        mosaicUtility.LogToFile(ex.Message + " : Error retrieving list items from 'Source' in case site " + caseSiteUrl);
                    }

 

                    // Retrieve and parse values for Attorney Lead for Case and Case Number from the "Additional Information" list
                    // on the current site

                    try
                    {
                        XElement spListAdditionalInfo = proxyList.GetList("Additional Information");
                        XElement nodeAdditionalInfoListItems = proxyList.GetListItems("Additional Information", null, null, null, "65535", null, null);

                        System.Xml.XmlDocument xdadditionalinfo = new System.Xml.XmlDocument();
                        xdadditionalinfo.LoadXml(nodeAdditionalInfoListItems.ToString());
                        System.Xml.XmlNamespaceManager nmadditionalinfo = new System.Xml.XmlNamespaceManager(xdadditionalinfo.NameTable);
                        nmadditionalinfo.AddNamespace("rs", "urn:schemas-microsoft-com:rowset");
                        nmadditionalinfo.AddNamespace("z", "#RowsetSchema");
                        nmadditionalinfo.AddNamespace("rootNS", "http://schemas.microsoft.com/sharepoint/soap/");
                        System.Xml.XmlNodeList nladditionalinfo = xdadditionalinfo.SelectNodes("/rootNS:listitems/rs:data/z:row", nmadditionalinfo);
                        //System.Console.WriteLine("Attorney Lead For Case, Case Number");

                        foreach (System.Xml.XmlNode listItem in nladditionalinfo)
                        {
                            attorneyLeadForCase = listItem.Attributes["ows_Attorney_x0020_Lead_x0020_For_x0"].Value.ToString();
                            attorneyLeadForCase = attorneyLeadForCase.Replace(",", " ");
                            attorneyLeadForCase = attorneyLeadForCase.Replace("#", "");
                            attorneyLeadForCase = attorneyLeadForCase.Replace(";", "");
                            caseNumber = listItem.Attributes["ows_Case_x0020_Number"].Value.ToString();

                        }
                    }
                    catch (Exception ex)
                    {
                        attorneyLeadForCase = "Not available";
                        caseNumber = "Not available";
                        mosaicUtility.LogToFile(ex.Message + ": List 'Additional Information' was not found in case site" + caseSiteUrl);
                    }


                    // Parse all values from the Sources list in SharePoint
                    // This list contains  data about the Legal Holds

                    System.Xml.XmlDocument xd = new System.Xml.XmlDocument();
                    if (nodeListItems != null)
                        xd.LoadXml(nodeListItems.ToString());
                    System.Xml.XmlNamespaceManager nm = new System.Xml.XmlNamespaceManager(xd.NameTable);
                    nm.AddNamespace("rs", "urn:schemas-microsoft-com:rowset");
                    nm.AddNamespace("z", "#RowsetSchema");
                    nm.AddNamespace("rootNS", "http://schemas.microsoft.com/sharepoint/soap/");
                    System.Xml.XmlNodeList nl = xd.SelectNodes("/rootNS:listitems/rs:data/z:row", nm);
                    //System.Console.WriteLine("Case Name, Case Number, Case Status, Attorney Lead For Case, Legal Hold User Name, EmployeeID, Legal Hold User Employee Status, legal Hold User Termed Date, Email Address/userId, Date On Hold, Date Released"); //ows_LinkTitle, ows_DiscoverySourceType, ows_DiscoveryPreservationStatus, ows_Title, ows_Created, ows_Modified");                                              

                    foreach (System.Xml.XmlNode listItem in nl)
                    {

                        string userAccount = "\u0022" + listItem.Attributes["ows_LinkTitle"].Value.ToString() + "\u0022";
                        string displayName = listItem.Attributes["ows_LinkTitle"].Value.ToString();
                        userAccount = userAccount.Replace(",", " ");
                        
                        string sourceType = listItem.Attributes["ows_DiscoverySourceType"].Value.ToString();
                        string legalHoldStatus = listItem.Attributes["ows_DiscoveryPreservationStatus"].Value.ToString().Trim();
                        DateTime createdDate = Convert.ToDateTime(listItem.Attributes["ows_Created"].Value.ToString());
                        DateTime modifiedDate = Convert.ToDateTime(listItem.Attributes["ows_Modified"].Value.ToString());
                        //System.Console.Write("   " + caseName);             //Case Name
                        //System.Console.Write("   " + caseNumber);           //Case Number
                        //System.Console.Write("   " + legalHoldStatus);      //Case Status
                        //System.Console.Write("   " + attorneyLeadForCase);  //Attorney Lead For Case                       
                        //System.Console.Write("   " + userAccount);          //Legal Hold User Name
                        //System.Console.Write("   " + userAccount);          //EmployeeID
                        //System.Console.Write("   " + legalHoldStatus);      //Legal Hold User Employee Status
                        //System.Console.Write("   " + createdDate);          //Legal Hold User Termed Date
                        //System.Console.Write("   " + userAccount);          //Email Address/UserId
                        ////System.Console.Write("   " + sourceType);
                        //System.Console.Write("   " + createdDate);          //Date On Hold  (clarify)
                        //System.Console.Write("   " + modifiedDate);         //Date Released (clarify)
                        ////System.Console.Write("   " + listItem.Attributes["ows_Title"].Value.ToString());
                    
                        //Get properties from AD for Legal Hold Report
                        string adEmployeeId = string.Empty;
                        string adEmployeeLegalHoldStatus = string.Empty;
                        string adEmailAddress = string.Empty;
                        string adLegalHoldUserTermedDate = string.Empty;
                        try
                        {
                            //Call the utility method to retrieve values from AD which need to be added to the Legal Hold Report
                            mosaicUtility.ReportProperties = mosaicUtility.GetLegalHoldReportProperties(displayName);

                            adEmployeeId = mosaicUtility.ReportProperties.EmployeeId;
                            adEmployeeLegalHoldStatus = mosaicUtility.ReportProperties.LegalHoldUserEmployeeStatus;
                            adEmailAddress = mosaicUtility.ReportProperties.EmailAddress;
                            adLegalHoldUserTermedDate = mosaicUtility.ReportProperties.LegalHoldUserTermedDate;

                        }
                        catch (Exception ex)
                        {
                            adEmployeeId = "Error retrieving employeeId";
                            adEmployeeLegalHoldStatus = "Error retrieving employeeId employee legal hold status";
                            adEmailAddress = "Error retrieving email address";
                            adLegalHoldUserTermedDate = "Error retrieving legal hold user termed date";

                            mosaicUtility.LogToFile(ex.Message + " : Error connecting to AD. Could not retrieve values from AD.");
                        }


                        var newExportLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}", caseName, caseNumber, legalHoldStatus, attorneyLeadForCase, userAccount, adEmployeeId, adEmployeeLegalHoldStatus, adLegalHoldUserTermedDate, adEmailAddress, createdDate, modifiedDate);
                        csvExportFileContent.AppendLine(newExportLine);
                                                
                    }

                    //Write report data to a file
                    try
                    {
                        System.IO.File.WriteAllText(@".\\dataexport.csv", csvExportFileContent.ToString());
                    }
                    catch (Exception ex)
                    {
                        mosaicUtility.LogToFile(ex.Message + " : Cannot write Legal Hold Report data to file.");
                    }

                }
                
            }

            System.Console.WriteLine("LegalHoldReport Console Application - Done.");                
            
        }
    }
}
