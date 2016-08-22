// Copyright � Microsoft Corporation.  All Rights Reserved.
// This code released under the terms of the 
// Microsoft Public License (MS-PL, http://opensource.org/licenses/ms-pl.html.)

/*
Sample Code is provided for the purpose of illustration only and is not intended 
to be used in a production environment. THIS SAMPLE CODE AND ANY RELATED INFORMATION 
ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS 
FOR A PARTICULAR PURPOSE. We grant You a nonexclusive, royalty-free right to 
use and modify the Sample Code and to reproduce and distribute the object code 
form of the Sample Code, provided that. You agree: (i) to not use Our name, logo, 
or trademarks to market Your software product in which the Sample Code is embedded; 
(ii) to include a valid copyright notice on Your software product in which the 
Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us 
and Our suppliers from and against any claims or lawsuits, including attorneys� 
fees, that arise or result from the use or distribution of the Sample Code
*/

using System;
using System.Text;
using System.Collections.Generic;
using System.Configuration;
using System.Threading;
using o365stepbystep.Services;
using o365stepbystep.Models;
using System.Text.RegularExpressions;
using Newtonsoft.Json;


namespace o365stepbystep
{
    class Program
    {       
        internal static O365ServiceApi o365 = new O365ServiceApi();
        internal volatile List<Content> _content = new List<Content>();
        internal static ReportingWebServiceApi reporting = 
            new ReportingWebServiceApi(
                "javier@javier.nyc", 
                "password123",
                "MailboxUsageDetail");
        private const string pattern = @"\r\n";        
        private const string _header = "CreationTime,RecordType,UserType,PolicyName,FileName,FileOwner"
                + ",FilePathUrl,From";
        private static Regex rgx = new Regex(pattern);

        
        
        public static void Main()
        {
            List<CsvRow> rowCollection = new List<CsvRow>();

            List<Subscription> subscriptions = o365.ListSubscriptions();
            
            dynamic temp = o365.GetServices();
            temp = o365.GetCurrentStatus();

            temp = o365.GetHistoricalStatus();

            temp = o365.GetMessages();

            var que = reporting._stream1;
                        
            List<ContentMetaData> sp = o365.ListAvailableContent("Audit.Sharepoint");            

            int i;
            int x;

            //Sharepoint
            for (i = 0; i< sp.Count; i++)
            {
                dynamic spc = o365.GetContent(sp[i].contentId);
                for(x =0; x < spc.Count; x++)
                {
                    if(spc[x].Operation== "DLPRuleMatch" && spc[x].UserId == "DLPAgent")
                    {                        
                        string dlprule = rgx.Replace(spc[x].ToString(), "");
                        SharepointDlp tempdlp = JsonConvert.DeserializeObject<SharepointDlp>(dlprule);
                        
                        CsvRow Sharepointcsvrow = new CsvRow();
                        Sharepointcsvrow.Add(String.Format("{0},{1},{2},{3},{4},{5},{6},{7}", 
                           tempdlp.CreationTime,                            
                            Enum.GetName(typeof(RecordType),tempdlp.RecordType),
                            Enum.GetName(typeof(UserType), tempdlp.UserType),
                            tempdlp.PolicyDetails[0].PolicyName,
                            tempdlp.SharePointMetaData.FileName,
                            tempdlp.SharePointMetaData.FileOwner,
                            tempdlp.SharePointMetaData.FilePathUrl,
                            tempdlp.SharePointMetaData.From
                            ));
                        rowCollection.Add(Sharepointcsvrow);                    
                    }
                    
                }
            }

            foreach(List<ContentMetaData> npc in o365._nextPageContent)
            {
                foreach(ContentMetaData cmd in npc)
                {
                    dynamic spc = o365.GetContent(cmd.contentId);
                    for (x = 0; x < spc.Count; x++)
                    {
                        if (spc[x].Operation == "DLPRuleMatch" && spc[x].UserId == "DLPAgent")
                        {
                            string dlprule = rgx.Replace(spc[x].ToString(), "");
                            SharepointDlp tempdlp = JsonConvert.DeserializeObject<SharepointDlp>(dlprule);

                            CsvRow Sharepointcsvrow = new CsvRow();
                            Sharepointcsvrow.Add(String.Format("{0},{1},{2},{3},{4},{5},{6},{7}",
                               tempdlp.CreationTime,
                                Enum.GetName(typeof(RecordType), tempdlp.RecordType),
                                Enum.GetName(typeof(UserType), tempdlp.UserType),
                                tempdlp.PolicyDetails[0].PolicyName,
                                tempdlp.SharePointMetaData.FileName,
                                tempdlp.SharePointMetaData.FileOwner,
                                tempdlp.SharePointMetaData.FilePathUrl,
                                tempdlp.SharePointMetaData.From
                                ));
                            rowCollection.Add(Sharepointcsvrow);
                        }

                    }
                }
            }

            WriteToFile(rowCollection,_header);
           


        }
        
        public static void WriteToFile(List<CsvRow> rowCollection, string header)
        {
            // Write sample data to CSV file
            using (CsvFileWriter writer = new CsvFileWriter(@"C:\Temp\WriteTest.csv",header))
            {
                foreach(CsvRow row in rowCollection)
                {
                    writer.WriteRow(row);
                }
                
                
            }
        }

    }
}

