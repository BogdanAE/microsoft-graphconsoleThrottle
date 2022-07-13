using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using System.Text.Json;
using Helpers;

namespace graphconsoleapp
{
    public class Program
    {
        private static object? _deltaLink = null;
        private static IUserDeltaCollectionPage? _previousPage = null;

        public static void Main(string[] args)
        {
            var config = LoadAppSettings();
            if (config == null)
            {
                Console.WriteLine("Invalid appsettings.json file.");
                return;
            }
            Console.WriteLine("All users in tenant:");
            CheckForUpdates(config);
            Console.WriteLine();
            while (true)
            {
                Console.WriteLine("... sleeping for 10s - press CTRL+C to terminate");
                System.Threading.Thread.Sleep(10 * 1000);
                Console.WriteLine("> Checking for new/updated users since last query...");
                CheckForUpdates(config);
            }
            //Bogdan - commented for pooling example -after HTTP erquest and SDK
            // //Bogdan - used in HTTP Request - removed for SDK
            // // var client = GetAuthenticatedHTTPClient(config);

            // //Bogdan - used in HTTP Request - removed for SDK
            // // var profileResponse = client.GetAsync("https://graph.microsoft.com/v1.0/me").Result;
            // // var profileJson = profileResponse.Content.ReadAsStringAsync().Result;
            // // var profileObject = JsonDocument.Parse(profileJson);
            // // var displayName = profileObject.RootElement.GetProperty("displayName").GetString();
            // // Console.WriteLine("Hello " + displayName);

            // var client = GetAuthenticatedGraphClient(config);

            // var profileResponse = client.Me.Request()
            //                                 .GetAsync()
            //                                 .Result;
            // var stopwatch = new System.Diagnostics.Stopwatch();
            // stopwatch.Start();

            // //Bogdan - used in HTTP Request - removed for SDK
            // // var clientResponse = client.GetAsync("https://graph.microsoft.com/v1.0/me/messages?$select=id&$top=100").Result;
            // var clientResponse = client.Me.Messages.Request()
            //                                 .Select(m => new { m.Id })
            //                                 .Top(100)
            //                                 .GetAsync()
            //                                 .Result;
            // var items = clientResponse.CurrentPage;

            // //Bogdan - used in HTTP Request - removed for SDK
            // // // enumerate through the list of messages
            // // var httpResponseTask = clientResponse.Content.ReadAsStringAsync();
            // // httpResponseTask.Wait();
            // // var graphMessages = JsonSerializer.Deserialize<Messages>(httpResponseTask.Result);
            // // var items = graphMessages == null ? Array.Empty<Message>() : graphMessages.Items;

            // var tasks = new List<Task>();
            // foreach (var graphMessage in items)
            // {
            //     tasks.Add(Task.Run(() =>
            //     {

            //         Console.WriteLine("...retrieving message: {0}", graphMessage.Id);

            //         var messageDetail = GetMessageDetail(client, graphMessage.Id);

            //         if (messageDetail != null)
            //         {
            //             Console.WriteLine("SUBJECT: {0}", messageDetail.Subject);
            //         }

            //     }));
            // }

            // // do all work in parallel & wait for it to complete
            // var allWork = Task.WhenAll(tasks);
            // try
            // {
            //     allWork.Wait();
            // }
            // catch { }

            // stopwatch.Stop();
            // Console.WriteLine();
            // Console.WriteLine("Elapsed time: {0} seconds", stopwatch.Elapsed.Seconds);
            //BOGDAN  - repvious example with reqeust done in parallel
            // var totalRequests = 100;
            // var successRequests = 0;
            // var tasks = new List<Task>();
            // var failResponseCode = HttpStatusCode.OK;
            // HttpResponseHeaders failedHeaders = null!;

            // for (int i = 0; i < totalRequests; i++)
            // {
            //     tasks.Add(Task.Run(() =>
            //     {
            //         var response = client.GetAsync("https://graph.microsoft.com/v1.0/me/messages").Result;
            //         Console.Write(".");
            //         if (response.StatusCode == HttpStatusCode.OK)
            //         {
            //             successRequests++;
            //         }
            //         else
            //         {
            //             Console.Write('X');
            //             failResponseCode = response.StatusCode;
            //             failedHeaders = response.Headers;
            //         }
            //     }));
            // }

            // var allWork = Task.WhenAll(tasks);
            // try
            // {
            //     allWork.Wait();
            // }
            // catch { }
            // Console.WriteLine();
            // Console.WriteLine("{0}/{1} requests succeeded.", successRequests, totalRequests);
            // if (successRequests != totalRequests)
            // {
            //     Console.WriteLine("Failed response code: {0}", failResponseCode.ToString());
            //     Console.WriteLine("Failed response headers: {0}", failedHeaders);
            // }
        }

        private static void OutputUsers(IUserDeltaCollectionPage users)
        {
            foreach (var user in users)
            {
                Console.WriteLine($"User: {user.Id}, {user.GivenName} {user.Surname}");
            }
        }

        private static IUserDeltaCollectionPage GetUsers(GraphServiceClient graphClient, object? deltaLink)
        {
            IUserDeltaCollectionPage page;

            // IF this is the first request, then request all users
            //    and include Delta() to request a delta link to be included in the
            //    last page of data
            if (_previousPage == null || deltaLink == null)
            {
                page = graphClient.Users
                                  .Delta()
                                  .Request()
                                  .Select("Id,GivenName,Surname")
                                  .GetAsync()
                                  .Result;
            }
            // ELSE, not the first page so get the next page of users
            else
            {
                _previousPage.InitializeNextPageRequest(graphClient, deltaLink.ToString());
                page = _previousPage.NextPageRequest.GetAsync().Result;
            }

            _previousPage = page;
            return page;
        }

        private static void CheckForUpdates(IConfigurationRoot config)
        {
            var graphClient = GetAuthenticatedGraphClient(config);

            // get a page of users
            var users = GetUsers(graphClient, _deltaLink);

            OutputUsers(users);

            // go through all of the pages so that we can get the delta link on the last page.
            while (users.NextPageRequest != null)
            {
                users = users.NextPageRequest.GetAsync().Result;
                OutputUsers(users);
            }
            object? deltaLink;

            if (users.AdditionalData.TryGetValue("@odata.deltaLink", out deltaLink))
            {
                _deltaLink = deltaLink;
            }
        }

        private static IConfigurationRoot? LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                                  .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                                  .AddJsonFile("appsettings.json", false, true)
                                  .Build();

                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["tenantId"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = PublicClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithDefaultRedirectUri()
                                                    .Build();
            return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray());
        }

        private static Microsoft.Graph.Message GetMessageDetail(GraphServiceClient client, string messageId, int defaultDelay = 2)
        {
            return client.Me.Messages[messageId].Request().GetAsync().Result;
        }

        //Bogdan - used in HTTP Request - removed for SDK
        // private static Message? GetMessageDetail(HttpClient client, string messageId, int defaultDelay = 2)
        // {
        //     Message? messageDetail = null;

        //     string endpoint = "https://graph.microsoft.com/v1.0/me/messages/" + messageId;

        //     // add code here
        //     // submit request to Microsoft Graph & wait to process response
        //     var clientResponse = client.GetAsync(endpoint).Result;
        //     var httpResponseTask = clientResponse.Content.ReadAsStringAsync();
        //     httpResponseTask.Wait();

        //     Console.WriteLine("...Response status code: {0}  ", clientResponse.StatusCode);

        //     // IF request successful (not throttled), set message to retrieved message
        //     if (clientResponse.StatusCode == HttpStatusCode.OK)
        //     {
        //         messageDetail = JsonSerializer.Deserialize<Message>(httpResponseTask.Result);
        //     } // ELSE IF request was throttled (429, aka: TooManyRequests)...
        //     else if (clientResponse.StatusCode == HttpStatusCode.TooManyRequests)
        //     {
        //         // get retry-after if provided; if not provided default to 2s
        //         var retryAfterDelay = defaultDelay;
        //         var retryAfter = clientResponse.Headers.RetryAfter;
        //         if (retryAfter != null && retryAfter.Delta.HasValue && (retryAfter.Delta.Value.Seconds > 0))
        //         {
        //             retryAfterDelay = retryAfter.Delta.Value.Seconds;
        //         }

        //         // wait for specified time as instructed by Microsoft Graph's Retry-After header,
        //         //    or fall back to default
        //         Console.WriteLine(">>>>>>>>>>>>> sleeping for {0} seconds...", retryAfterDelay);
        //         System.Threading.Thread.Sleep(retryAfterDelay * 1000);

        //         // call method again after waiting
        //         messageDetail = GetMessageDetail(client, messageId);
        //     }

        //     return messageDetail;
        // }

        //Bogdan - used for HTTP request - removed for SDK example
        // private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config)
        // {
        //     var authenticationProvider = CreateAuthorizationProvider(config);
        //     var httpClient = new HttpClient(new AuthHandler(authenticationProvider, new HttpClientHandler()));
        //     return httpClient;
        // }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvide = CreateAuthorizationProvider(config);
            var graphClient = new GraphServiceClient(authenticationProvide);
            return graphClient;
        }
    }
}