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
using System.Configuration;
using System.Text;
using Microsoft.IdentityModel.Clients.ActiveDirectory;


namespace o365stepbystep.Services
{
    class AADHelper
    {
        private static string AuthorityUrl = ConfigurationManager.AppSettings["AAD:AuthorityUrl"];
        private static string ClientId = ConfigurationManager.AppSettings["AAD:ClientId"];
        private static string ClientKey = ConfigurationManager.AppSettings["AAD:ClientKey"];
        private static Uri ClientRedirectUri = new Uri(ConfigurationManager.AppSettings["AAD:ClientRedirectUri"]);
        private static string ResourceID = ConfigurationManager.AppSettings["AAD:ResourceID"];

        /// <summary>
        /// Retrieves an access token from AAD using the client app credentials.
        /// </summary>
        /// <param name="tenantId"></param>
        /// <returns></returns>
        public static string AuthenticateAndGetToken(string tenantId)
        {
            AuthenticationContext authContext = new AuthenticationContext(String.Format(AuthorityUrl, tenantId));
            ClientCredential clientCredentials = new ClientCredential(AADHelper.ClientId, AADHelper.ClientKey);

            AuthenticationResult result = null;
            string accessToken;

            try
            {
                result = authContext.AcquireTokenAsync(AADHelper.ResourceID, clientCredentials).Result;
                accessToken = result.AccessToken;

                if (String.IsNullOrEmpty(accessToken))
                {
                    Console.WriteLine("A token was not received from AAD.");
                }
            }
            catch (AdalException ex)
            {
                string message = ex.Message;
                if (ex.InnerException != null)
                {
                    message += "Inner Exception : " + ex.InnerException.Message;
                }
                Console.WriteLine("There was an exception when requesting a token from AAD.");
                Console.WriteLine(message);
            }

            return result.AccessToken;
        }
    }


    

}


