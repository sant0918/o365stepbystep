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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace o365stepbystep
{
    
    public partial class Webhook
    {
        public string address { get; set; }
        public string authID { get; set; }
        public string expiration { get; set; }
    }

    public partial class StartInput
    {
        public Webhook webhook { get; set; }
    }

    public enum RecordType {
        ExchangeAdmin = 1,
        ExchangeItem = 2,
        ExchangeItemGroup = 3,
        SharePoint = 4,
        SharePointFileOperation = 6,
        AzureActiveDirectory = 8,
        AzureActiveDirectoryAccountLogon = 9,
        DataCenterSecurityCmdlet = 10,
        ComplianceDLPSharePoint = 11,
        Sway = 12,
        SharePointSharingOperation = 14,
        AzureActiveDirectoryStsLogon = 15
    }

    public enum UserType
    {
        Regular,
        Reserved,
        Admin,
        DcAdmin,
        System,
        Application,
        ServicePrincipal
    }




    }

