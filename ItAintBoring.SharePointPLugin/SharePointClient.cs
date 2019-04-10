using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;

namespace ItAintBoring.SharePointPLugin
{
    public class SharePointClient
    {
        private string token = null;

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

        public async Task<string> GetToken()
        {
            using (var httpClient = new HttpClient())
            {
                
                
                try
                {
                    //StringContent queryString = new StringContent("");
                    
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
                    string responseBody = await response.Content.ReadAsStringAsync();


                    response.EnsureSuccessStatusCode();
                    Dictionary<string, string> tokenAttributes = new Dictionary<string, string>();
                    string[] tokenData = responseBody.Split(',');
                    foreach(var s in tokenData)
                    {
                        string[] pair = s.Split(':');
                        tokenAttributes.Add(pair[0], pair[1]);
                    }

                    //StringContent queryString = new StringContent(data);
                    //var response = httpClient.PostStringAsync().Result;
                    //"access_token"
                    response = response;
                }
                catch(Exception ex)
                {
                    ex = ex;
                }
            }
            return "";
        }
    }
}
