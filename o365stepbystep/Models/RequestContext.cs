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
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Text;


namespace o365stepbystep
{
    public class RequestContext
    {
        public string TenantId;
        public string ContentId;
        public string ContentType;
        public string ContentUrl;
        public string WebhookAddress;
        public string WebhookAuthId;
        public string WebhookExpiration;
        public string StartTime;
        public string EndTime;
        public string Token;

        public RequestContext(Dictionary<string, string> parameters)
        {
            // default values from configuration
            TenantId = ConfigurationManager.AppSettings["TenantId"];
            WebhookAddress = ConfigurationManager.AppSettings["WebhookAddress"];
            WebhookAuthId = ConfigurationManager.AppSettings["WebhookAuthId"];
            WebhookExpiration = ConfigurationManager.AppSettings["WebhookExpiration"];

            // parsed parameters
            TenantId = parameters.ContainsKey("tenantid") ? parameters["tenantid"] : TenantId;
            ContentId = parameters.ContainsKey("contentid") ? parameters["contentid"] : "";
            ContentType = parameters.ContainsKey("contenttype") ? parameters["contenttype"] : "";
            ContentUrl = parameters.ContainsKey("contenturl") ? parameters["contenturl"] : "";
            WebhookAddress = parameters.ContainsKey("webhookaddress") ? parameters["webhookaddress"] : WebhookAddress;
            WebhookAuthId = parameters.ContainsKey("webhookauthid") ? parameters["webhookauthid"] : WebhookAuthId;
            WebhookExpiration = parameters.ContainsKey("webhookexpiration") ? parameters["webhookexpiration"] : WebhookExpiration;
            StartTime = parameters.ContainsKey("starttime") ? parameters["starttime"] : "";
            EndTime = parameters.ContainsKey("endtime") ? parameters["endtime"] : "";
            Token = parameters.ContainsKey("token") ? parameters["token"] : "";
        }
    }
}

  