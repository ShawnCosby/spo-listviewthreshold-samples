using Microsoft.SharePoint.Client;
using System;
using System.Net;
using System.Threading;

namespace SpoLvtTest
{
    /// <summary>
    /// Best practices for handling throttling in SharePoint Online. Every instance of a <see cref="ClientContext"/> must decorate traffic and incrementally retry any queries that are HTTP-429 throttled .
    /// </summary>
    /// <see cref="https://docs.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online"/>
    public static class ClientContextExtensions
    {
        private static readonly int RETRY_ATTEMPTS = 5;
        private static readonly int RETRY_DELAY_IN_SECONDS = 10;
        private static readonly string WEBREQUEST_USERAGENT = "ISV|Gimmal|SPO-LVT-Test/1.0.0.0";
        private static readonly int WEBREQUEST_TIMEOUT = (int)TimeSpan.FromMinutes(5D).TotalMilliseconds;
        private static readonly HttpStatusCode HTTPSTATUSCODE_TOOMANYREQUESTS = (HttpStatusCode)429; //This status code is documented in IETF RFC 6585

        /// <summary>
        /// Set the web request's user agent in order to decorate the HTTP traffic
        /// </summary>
        /// <see cref="https://docs.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online#how-to-decorate-your-http-traffic-to-avoid-throttling"/>
        public static void SetTrafficDecorator(this ClientContext clientContext)
        {
            if (clientContext != null)
            {
                clientContext.ExecutingWebRequest += (sender, args) =>
                {
                    var request = args?.WebRequestExecutor?.WebRequest;

                    if (request != null)
                    {
                        request.UserAgent = WEBREQUEST_USERAGENT;
                        request.Timeout = WEBREQUEST_TIMEOUT;
                    }
                };
            }
        }

        /// <summary>
        /// Calls <see cref="ExecuteQueryWithIncrementalRetry(ClientContext, Action{string})"> with default values for retryCount and delay
        /// </summary>
        /// <see cref="https://docs.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online#csom-code-sample-executequerywithincrementalretry-extension-method"/>
        public static void ExecuteQueryWithIncrementalRetry(this ClientContext clientContext)
        {
            ExecuteQueryWithIncrementalRetry(clientContext, null);
        }

        /// <summary>
        /// Calls <see cref="InternalExecuteQueryWithIncrementalRetry(ClientContext, int, int, Action{string})"> with default values for retryCount and delay and provides a delegate for logging HTTP throttling
        /// </summary>
        public static void ExecuteQueryWithIncrementalRetry(this ClientContext clientContext, Action<string> throttlingLogAction)
        {
            int retryCount = RETRY_ATTEMPTS;
            int delay = RETRY_DELAY_IN_SECONDS;

            InternalExecuteQueryWithIncrementalRetry(clientContext, retryCount, delay, throttlingLogAction);
        }

        private static void InternalExecuteQueryWithIncrementalRetry(this ClientContext clientContext,
                                                                     int retryCount,
                                                                     int delay,
                                                                     Action<string> throttlingLogAction)
        {
            int retryAttempts = 0;
            int backoffInterval = delay;
            int retryAfterInterval = 0;
            bool retry = false;
            ClientRequestWrapper wrapper = null;

            if (retryCount <= 0)
            {
                throw new ArgumentException("Provide a retry count greater than zero.");
            }

            if (delay <= 0)
            {
                throw new ArgumentException("Provide a delay greater than zero.");
            }

            // Do while retry attempt is less than retry count
            while (retryAttempts < retryCount)
            {
                try
                {
                    if (!retry)
                    {
                        clientContext.ExecuteQuery();

                        return;
                    }
                    else
                    {
                        // retry the previous request
                        if (wrapper != null && wrapper.Value != null)
                        {
                            clientContext.RetryQuery(wrapper.Value);

                            return;
                        }
                    }
                }
                catch (WebException ex)
                {
                    var response = ex.Response as HttpWebResponse;

                    // Check if request was throttled - http status code 429
                    // Check is request failed due to server unavailable - http status code 503
                    if (response != null && (response.StatusCode == HTTPSTATUSCODE_TOOMANYREQUESTS ||
                                             response.StatusCode == HttpStatusCode.ServiceUnavailable))
                    {
                        string throttlingLogMsg = $"Throttling occurred (HTTP {(int)response.StatusCode})! Retry attempt: {retryAttempts + 1}.";

                        wrapper = (ClientRequestWrapper)ex.Data["ClientRequest"];
                        retry = true;

                        // Determine the retry after value - use the retry-after header when available
                        string retryAfterHeader = response.GetResponseHeader("Retry-After");
                        if (!string.IsNullOrEmpty(retryAfterHeader))
                        {
                            throttlingLogMsg += $" Retry-After header: {retryAfterHeader}.";

                            if (!int.TryParse(retryAfterHeader, out retryAfterInterval))
                            {
                                retryAfterInterval = backoffInterval;
                            }
                        }
                        else
                        {
                            throttlingLogMsg += $" This response did not issue a Retry-After header, setting retry interval to {backoffInterval} seconds.";
                            retryAfterInterval = backoffInterval;
                        }

                        if (throttlingLogAction != null)
                        {
                            throttlingLogAction.Invoke(throttlingLogMsg);
                        }

                        // Wait until retry-after interval has elapsed...
                        Thread.Sleep(TimeSpan.FromSeconds(retryAfterInterval));

                        // Increase counters
                        retryAttempts++;
                        backoffInterval = backoffInterval * 2;
                    }
                    else
                    {
                        throw;
                    }
                }
            }

            throw new InvalidOperationException($"Maximum number of retries ({retryCount}) have been attempted.");
        }
    }
}
