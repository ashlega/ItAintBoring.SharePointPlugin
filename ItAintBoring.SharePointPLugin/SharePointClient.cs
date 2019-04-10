using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;

namespace ItAintBoring.SharePointPlugin
{

    public class AccessToken
    {
        public string Token { get; set; }
        public DateTime ExpirationDate { get; set; }
    }

    public class SharePointClient
    {
        private AccessToken Token = null;

        string clientId = null;
        string tenantId = null;
        string clientKey = null;
        string siteRoot = null;

        public SharePointClient(string clientId, string clientKey, string tenantId, string siteRoot)
        {
            this.clientId = clientId;
            this.tenantId = tenantId;
            this.clientKey = clientKey;
            this.siteRoot = siteRoot;
            
        }

        /// <summary>
        /// Get the token from Sharepoint
        /// </summary>
        /// <returns></returns>
        public async Task GetToken(bool forceRefresh = false)
        {
            if (!forceRefresh && Token != null && Token.ExpirationDate >= DateTime.Now)
            {
                return;
            }

            string responseBody =  null;
            using (var httpClient = new HttpClient())
            {
                try
                {
                    FormUrlEncodedContent content = new FormUrlEncodedContent(
                        new[]
                        {
                            new KeyValuePair<string, string>("grant_type", "client_credentials"),
                            new KeyValuePair<string, string>("client_id", $"{clientId}@{tenantId}"),
                            new KeyValuePair<string, string>("client_secret", clientKey),
                            new KeyValuePair<string, string>("resource", $"00000003-0000-0ff1-ce00-000000000000/{siteRoot}@{tenantId}")
                        });
                    string url = $"https://accounts.accesscontrol.windows.net/{tenantId}/tokens/OAuth/2";
                    HttpResponseMessage response = await httpClient.PostAsync(new Uri(url), content);
                    //response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                    responseBody = await response.Content.ReadAsStringAsync();
                    response.EnsureSuccessStatusCode();
                    //Dictionary<string, string> tokenAttributes = new Dictionary<string, string>();
                    responseBody = responseBody.Replace("{", "").Replace("}", "");
                    string[] tokenData = responseBody.Split(',');
                    Token = new AccessToken();
                    foreach(var s in tokenData)
                    {
                        string[] pair = s.Replace("\"", "").Split(':');
                        if(pair[0] == "expires_in")
                        {
                            Token.ExpirationDate = DateTime.Now.AddSeconds(int.Parse(pair[1]) - 600);
                        }
                        if (pair[0] == "access_token")
                        {
                            Token.Token = pair[1];
                        }
                        //tokenAttributes.Add(pair[0], pair[1]);
                    }
                }
                catch(Exception ex)
                {
                    throw new Exception(responseBody, ex);
                }
            }
        }

        /// <summary>
        /// Get the token from Sharepoint
        /// </summary>
        /// <returns></returns>
        private async Task<string> RunQueryInternal(string api, string json)
        {
            string responseBody = null;
            using (var httpClient = new HttpClient())
            {
                try
                {
                    await GetToken();//Get/refesh the token

                    string url = $"https://{siteRoot}/_api/web/{api}";
                    StringContent content = new StringContent(json);
                    content.Headers.Clear();
                    content.Headers.Add("Content-Type", "application/json;odata=verbose");

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);
                    request.Content = content;
                    request.Headers.Add("Authorization", $"Bearer {Token.Token}");
                    request.Headers.Add("Accept", "application/json;odata=verbose");

                    HttpResponseMessage response = await httpClient.SendAsync(request);

                    responseBody = await response.Content.ReadAsStringAsync();
                    response.EnsureSuccessStatusCode();
                }
                catch (Exception ex)
                {
                    throw new Exception(responseBody, ex);
                }
            }
            return responseBody;
        }

        public string RunQuery(string api, string json)
        {
            var t = Task.Run<string>(async () => await RunQueryInternal(api, json));
            return t.Result;
        }
    }
}
