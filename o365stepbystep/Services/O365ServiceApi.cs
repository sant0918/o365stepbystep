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
using System.Threading.Tasks;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization.Json;
using o365stepbystep.Models;
using Newtonsoft.Json;
using System.Configuration;
using System.Text.RegularExpressions;


namespace o365stepbystep.Services
{
    public class O365ServiceApi
    {
        internal RequestContext context;
        internal List<List<ContentMetaData>> _nextPageContent = new List<List<ContentMetaData>>();
        private string ActivityApiBaseUrl;
        private string _serviceCommunicationsApiBaseUrl;
        private string AccessToken;
        

        public O365ServiceApi()
        {
            this.context = new RequestContext(new Dictionary<string, string>());

            ActivityApiBaseUrl = String.Format(ConfigurationManager.AppSettings["ActivityApiBaseUrl"], this.context.TenantId);
            _serviceCommunicationsApiBaseUrl = String.Format(ConfigurationManager.AppSettings["ServiceCommunicationsApiBaseUrl"], this.context.TenantId);
            AccessToken = (string.IsNullOrEmpty(context.Token)) ? AADHelper.AuthenticateAndGetToken(context.TenantId) : context.Token;

            if (String.IsNullOrEmpty(AccessToken))
            {
                throw new ArgumentNullException(AccessToken);
            }
            
        }

        public void StartSubscription(string contentType)
        {
            StartInput startInput = new StartInput();
            Webhook webhook = new Webhook();
            webhook.address = context.WebhookAddress;
            webhook.authID = context.WebhookAuthId;
            webhook.expiration = context.WebhookExpiration;

            if (string.IsNullOrEmpty(webhook.address + webhook.authID + webhook.expiration))
            {
                startInput.webhook = null;
            }
            else
            {
                startInput.webhook = webhook;
            }

            string url = ActivityApiBaseUrl + "subscriptions/start?contentType=" + contentType;

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);
            request.Content = SerializeRequestBody(startInput);
            HttpResponseMessage response = client.SendAsync(request).Result;
            string output = response.Content.ReadAsStringAsync().Result;

            if (response.IsSuccessStatusCode == true)
            {
                Console.WriteLine("\n{0}\n", JSONHelper.FormatJson(output));
            }
            else
            {
                Console.WriteLine("\nError:  {0}\n More Details: {1}", response.ReasonPhrase, output);
            }
        }

        /// <summary>
        /// Stop a subscription.
        /// </summary>
        public async Task StopSubscription(string contentType)
        {
            string url = ActivityApiBaseUrl + "subscriptions/stop?contentType=" + contentType;

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);
            HttpResponseMessage response = await client.SendAsync(request);
            string output = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode == true)
            {
                Console.WriteLine("\n{0}\n", JSONHelper.FormatJson(output));
            }
            else
            {
                Console.WriteLine("\nError:  {0}\n More Details: {1}", response.ReasonPhrase, output);
            }
        }

        /// <summary>
        /// List existing subscriptions.
        /// </summary>
        public List<Subscription> ListSubscriptions()
        {
            string url = ActivityApiBaseUrl + "subscriptions/list";

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);

            HttpResponseMessage response =  client.SendAsync(request).Result;
            string output =  response.Content.ReadAsStringAsync().Result;

            if (response.IsSuccessStatusCode == true)
            {
                Console.WriteLine("\n{0}", JSONHelper.FormatJson(output));
                return JsonConvert.DeserializeObject<List<Subscription>>(output);
            }
            else
            {
                Console.WriteLine("\nError:  {0}\n More Details: {1}", response.ReasonPhrase, output);
                return null;
            }
        }

        /// <summary>
        /// List available content for a subscription.
        /// </summary>
        public List<ContentMetaData> ListAvailableContent(string contentType, string startTime = null,string endTime = null)
        {
            string url = ActivityApiBaseUrl + String.Format("subscriptions/content?contentType={0}&endtime={1}&starttime={2}", contentType, endTime, startTime);

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);

            HttpResponseMessage response = client.SendAsync(request).Result;
            
            string output = response.Content.ReadAsStringAsync().Result;
            
            if (response.IsSuccessStatusCode == true)
            {
                try
                {
                    while(true)
                    {
                        string nextPageUrl = Regex.Match(response.Headers.ToString(), @"NextPageUri:\s(?<value>.+)\r").Groups["value"].Value;
                        if (String.IsNullOrEmpty(nextPageUrl))
                            return JsonConvert.DeserializeObject<List<ContentMetaData>>(output);
                        
                        GetNextPage(nextPageUrl);
                    }

                    
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
            else
            {
                return null;
            }
        }

        private void GetNextPage(string url)
        {
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);

            HttpResponseMessage response = client.SendAsync(request).Result;
            this._nextPageContent.Add(JsonConvert.DeserializeObject <List<ContentMetaData>>(response.Content.ReadAsStringAsync().Result));

            string nextPageUrl = Regex.Match(response.Headers.ToString(), @"NextPageUri:\s(?<value>.+)\r").Groups["value"].Value;
            if (!String.IsNullOrEmpty(nextPageUrl))
                this.GetNextPage(nextPageUrl);
            
        }
        
        /// <summary>
        /// List notifications the service has sent for a subscription.
        /// </summary>
        public async Task GetNotifications(string contentType, string startTime = null, string endTime = null)
        {
            string url = ActivityApiBaseUrl + String.Format("subscriptions/notifications?contentType={0}&endtime={1}&starttime={2}", contentType, endTime, startTime);
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);

            HttpResponseMessage response = await client.SendAsync(request);
            string output = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode == true)
            {
                Console.WriteLine("\n{0}\n", JSONHelper.FormatJson(output));
            }
            else
            {
                Console.WriteLine("\nError:  {0}\n More Details: {1}", response.ReasonPhrase, output);
            }
        }

        /// <summary>
        /// Retrieve content.
        /// </summary>
        public dynamic GetContent(string contentId)
        {
            string url = context.ContentUrl;
            if (string.IsNullOrEmpty(url))
            {
                url = ActivityApiBaseUrl + "audit/" + contentId;
            }
            url.Replace("$", "%24");

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);

            HttpResponseMessage response = client.SendAsync(request).Result;
            string output = response.Content.ReadAsStringAsync().Result;

            if (response.IsSuccessStatusCode == true)
            {
                Console.WriteLine("\n{0}\n", JSONHelper.FormatJson(output));                
                return JsonConvert.DeserializeObject(JSONHelper.FormatJson(output));
                
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// Helper method for JSON serialization.
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public StringContent SerializeRequestBody(object input)
        {
            DataContractJsonSerializer jsonSer = new DataContractJsonSerializer(input.GetType());
            MemoryStream ms = new MemoryStream();
            jsonSer.WriteObject(ms, input);
            ms.Position = 0;
            StreamReader sr = new StreamReader(ms);
            StringContent content = new StringContent(sr.ReadToEnd(), System.Text.Encoding.UTF8, "application/json");

            return content;
        }

        /// <summary>
        /// Returns the list of subscribed services.
        /// </summary>
        /// <returns></returns>
        public dynamic GetServices() {

            string url = _serviceCommunicationsApiBaseUrl + "Services";

            url.Replace("$", "%24");

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);

            HttpResponseMessage response = client.SendAsync(request).Result;
            string output = response.Content.ReadAsStringAsync().Result;

            if (response.IsSuccessStatusCode == true)
            {
                Console.WriteLine("\n{0}\n", JSONHelper.FormatJson(output));
                return JsonConvert.DeserializeObject(JSONHelper.FormatJson(output));

            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Returns the current status of the service.
        /// </summary>
        /// <param name="workload">Filter by workload (String, default: all).</param>
        /// <returns></returns>
        public dynamic GetCurrentStatus(string workload = null) {
            string url = _serviceCommunicationsApiBaseUrl + "CurrentStatus";

            url.Replace("$", "%24");

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);

            HttpResponseMessage response = client.SendAsync(request).Result;
            string output = response.Content.ReadAsStringAsync().Result;

            if (response.IsSuccessStatusCode == true)
            {
                Console.WriteLine("\n{0}\n", JSONHelper.FormatJson(output));
                return JsonConvert.DeserializeObject(JSONHelper.FormatJson(output));

            }
            else
            {
                return null;
            }
             }

        /// <summary>
        /// Returns the historical status of the service, by day, over a certain time range
        /// </summary>
        /// <param name="workload">Filter by workload (String, default: all).</param>
        /// <param name="statusTime">Filter by days greater than StatusTime (DateTimeOffset, default: ge CurrentTime – 7 days).</param>
        /// <returns></returns>
        public dynamic GetHistoricalStatus(string workload = null, string statusTime = null) {
            string url = _serviceCommunicationsApiBaseUrl + "HistoricalStatus";

            url.Replace("$", "%24");

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);

            HttpResponseMessage response = client.SendAsync(request).Result;
            string output = response.Content.ReadAsStringAsync().Result;

            if (response.IsSuccessStatusCode == true)
            {
                Console.WriteLine("\n{0}\n", JSONHelper.FormatJson(output));
                return JsonConvert.DeserializeObject(JSONHelper.FormatJson(output));

            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Returns the messages about the service over a certain time range. 
        /// Use the type filter to filter for "Service Incident", 
        /// "Planned Maintenance", or "Message Center" messages.
        /// </summary>
        /// <param name="workload">Filter by workload (String, default: all).</param>
        /// <param name="startTime">Filter by Start Time (DateTimeOffset, default: ge CurrentTime – 7 days).</param>
        /// <param name="endTime">Filter by End Time (DateTimeOffset, default: le CurrentTime).</param>
        /// <param name="messageType">Filter by MessageType (String, default: all).</param>
        /// <param name="id">Filter by ID (String, default: all).</param>
        /// <returns></returns>
        public dynamic GetMessages(string workload = null,
            string startTime = null, 
            string endTime = null, 
            string messageType = null,
            string id = null) {


            string url = _serviceCommunicationsApiBaseUrl + "Messages";

            url.Replace("$", "%24");

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);

            HttpResponseMessage response = client.SendAsync(request).Result;
            string output = response.Content.ReadAsStringAsync().Result;

            if (response.IsSuccessStatusCode == true)
            {
                Console.WriteLine("\n{0}\n", JSONHelper.FormatJson(output));
                return JsonConvert.DeserializeObject(JSONHelper.FormatJson(output));

            }
            else
            {
                return null;
            }
        }
    }
}
