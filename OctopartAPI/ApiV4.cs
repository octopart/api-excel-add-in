using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.Net;
using System.Threading;
using Newtonsoft.Json;
using RestSharp;
using System.IO;
using System.Xml.Linq;

namespace OctopartApi
{
    public static class ApiV4
    {
        #region Variables
        /// <summary>
        /// Log (for debugging or otherwise)
        /// </summary>
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        #endregion

        #region Classes
        /// <summary>
        /// Encapsulates search results
        /// </summary>
        public class SearchResponse
        {
            /// <summary>
            /// Gets the data returned by the search query
            /// </summary>
            public object Data { get; internal set; }

            /// <summary>
            /// Gets the error message from search query (if available)
            /// </summary>
            public string ErrorMessage { get; internal set; }
        }
        #endregion

        #region Constants
        /// <summary>
        /// The standard limit per Octopart web request (set to 20 as indicated by the API)
        /// Note: Returning 20 results in up to 20 MPNs. In testing, this means data unrelated to the intended part is being pulled. Commenting out and setting to 1 pending future side effects.
        /// </summary>
        /// public const int RECORD_LIMIT_PER_QUERY = 20;
	    public const int RECORD_LIMIT_PER_QUERY = 1;

        /// <summary>
        /// The start limit for an Octopart web request 
        /// </summary>
        public const int RECORD_START_MAX = 80;
        #endregion

        #region Constants-Private
        /// <summary>
        /// Base URL for all queries
        /// </summary>
        private const string OCTOPART_URL_BASE = "https://octopart.com/api/v4/rest";

        /// <summary>
        /// Url request for the part match query
        /// </summary>
        private const string OCTOPART_URL_PART_MATCH_ENDPOINT = "parts/match";

        /// <summary>
        /// Url request to upload a BOM
        /// </summary>
        private const string OCTOPART_URL_BOM_UPLOAD = "https://octopart.com/bom-lookup/upload";

        /// <summary>
        /// Number of retries when an error occurs
        /// </summary>
        private const int QUERY_RETRIES = 5;
        #endregion

        #region Methods-Search
        /// <summary>
        /// Provides a delegate method for the PartMatch search query
        /// </summary>
        /// <param name="pn">Search String: Part Number based</param>
        /// <param name="startRecord">Start record number</param>
        /// <param name="numRecords">Record limit</param>
        /// <param name="apiKey">Octopart API Key (http://octopart.com/api/home)</param>
        /// <returns>
        /// Returns a list of parts that were found from the provided search string.
        /// NULL indicates that no parts were found, or there was an error with the request
        /// </returns>
        public delegate SearchResponse PartMatchKeyedDelegate(string pn, int startRecord, int numRecords, string apiKey);

        /// <summary>
        /// Execute a part/match endpoint search
        /// </summary>
        /// <param name="pnList">Search String: Part Number based</param>
        /// <param name="apiKey">Octopart API Key (http://octopart.com/api/home)</param>
        /// <param name="httpTimeout">The desired timeout for the http request (in ms)</param>
        /// <returns>
        /// Returns a list of parts that were found from the provided search string.
        /// NULL indicates that no parts were found, or there was an error with the request
        /// </returns>
        /// <notes>
        /// Server timeout is defaulted to 5000ms
        /// </notes>
        public static SearchResponse PartsMatch(List<ApiV4Schema.PartsMatchQuery> pnList, string apiKey, int httpTimeout = 5000)
        {
            // Check arguments
            if ((pnList == null) || (pnList.Count == 0))
            {
                return null;
            }

            // Create the search request
            var query = new List<Dictionary<string, string>>();
            foreach (ApiV4Schema.PartsMatchQuery pn in pnList)
            {
                query.Add(
                    new Dictionary<string, string>()
                    {
                        { "mpn", pn.q },
                        { "limit", pn.limit.ToString() },
                        { "start", pn.start.ToString() }
                    }
                  );
            }
            string queryString = JsonConvert.SerializeObject(query);

            // Query octopart.com
            var client = new RestClient(OCTOPART_URL_BASE);
            var req = new RestRequest(OCTOPART_URL_PART_MATCH_ENDPOINT, Method.GET)
                .AddParameter("apikey", apiKey)
                .AddParameter("queries", queryString);

            req.Timeout = httpTimeout;
            req.UseDefaultCredentials = true;
            req.RequestFormat = DataFormat.Json;
            req.OnBeforeDeserialization = r => { r.ContentType = "application/json"; };
            client.Proxy = WebRequest.DefaultWebProxy;

            IRestResponse<ApiV4Schema.PartsMatchResponse> resp = null;

            var ret = new SearchResponse();
            for (int i = 0; i < QUERY_RETRIES; i++)
            {
                resp = client.Execute<ApiV4Schema.PartsMatchResponse>(req);

                if (resp == null)
                {
                    ret.ErrorMessage = "Server did not provide a response";
                    Log.Error(string.Format("Unexpected Error (resp == null) '{0}'", ret.ErrorMessage));
                    break;
                }
                else if ((int)resp.StatusCode == 0)
                {
                    ret.ErrorMessage = "Server did not provide a response (" + resp.ErrorMessage + ")";
                    Log.Error(string.Format("Response error '{0}'", ret.ErrorMessage));
                    break;
                }
                else if (resp.StatusCode == HttpStatusCode.BadRequest)
                {
                    ret.ErrorMessage = "Bad Request. Please email contact@octopart.com for assistance.";
                    Log.Debug(string.Format("Response error '{0}'", ret.ErrorMessage));
                    break;
                }
                else if (resp.StatusCode == HttpStatusCode.Unauthorized)
                {
                    ret.ErrorMessage = "Invalid API Key. Please email contact@octopart.com for assistance.";
                    Log.Debug(string.Format("Response error '{0}'", ret.ErrorMessage));
                    break;
                }
                else if (resp.StatusCode == HttpStatusCode.Forbidden)
                {
                    ret.ErrorMessage = "Unauthorized Access. Please email contact@octopart.com for assistance.";
                    Log.Debug(string.Format("Response error '{0}'", ret.ErrorMessage));
                    break;
                }
                else if (resp.StatusCode == HttpStatusCode.ProxyAuthenticationRequired)
                {
                    // Attempt the request using default proxy information
                    if (i == 0)
                    {
                        if (client.Proxy != null)
                        {
                            client.Proxy.Credentials = WebRequest.DefaultWebProxy.Credentials;
                        }
                    }

                    // If still not working, prompt the user for credentials
                    if (i >= 2 && i < 4)
                    {
                        if (client.Proxy != null)
                        {
                            bool? result = false;
                            string user = string.Empty;
                            string pass = string.Empty;
                            var thread = new Thread(() =>
                            {
                                var tempDlg = new ProxyAuthPrompt(client.Proxy.GetProxy(new Uri("https://www.octopart.com")).AbsoluteUri);
                                result = tempDlg.ShowDialog();
                                if (result == true)
                                {
                                    user = tempDlg.User;
                                    pass = tempDlg.Pass;
                                }
                            });
                            thread.SetApartmentState(ApartmentState.STA);
                            thread.Start();
                            while (thread.IsAlive)
                            {
                                Thread.Sleep(1);
                            }

                            if (result == false)
                                break;

                            if (!string.IsNullOrEmpty(user) && !string.IsNullOrEmpty(pass))
                            {
                                client.Proxy.Credentials = new NetworkCredential(user, pass);
                            }
                        }
                    }

                    ret.ErrorMessage = "Proxy Authentication Required. Please verify if proxy is configured correctly.";
                }
                else if ((int)resp.StatusCode == 429)
                {
                    ret.ErrorMessage = "Server is overloaded (" + resp.ErrorMessage + ")";
                    Log.Debug(string.Format("Response error '{0}'", ret.ErrorMessage));
                }
                else if (resp.StatusCode != HttpStatusCode.OK)
                {
                    ret.ErrorMessage = String.Format("Invalid HTTP response ({0} - {1})", resp.StatusCode, resp.StatusCode);
                    Log.Debug(string.Format("Response error '{0}'", ret.ErrorMessage));
                    break;
                }
                else
                {
                    ret.Data = resp.Data;
                    ret.ErrorMessage = resp.ErrorMessage;
                    break;
                }

                Log.Debug(string.Format("Wait and try again {0} of {1}...", i, QUERY_RETRIES));
                Thread.Sleep(400);
            }

            return ret;
        }
        #endregion

        #region Methods-Upload
        /// <summary>
        /// Uploads a BOM to octopart.com, and displays the column chooser in the default browser
        /// </summary>
        /// <param name="email">The email authentication of the user</param>
        /// <param name="file">The file to upload to the BOM tool</param>
        /// <returns>An error string, if an error has occured</returns>
        public static string UploadBom(string email, string file)
        {
            string url = OCTOPART_URL_BOM_UPLOAD + "?user=" + email;
            const string paramName = "datafile";
            const string contentType = "application/vnd.ms-excel";

            while (true)
            {
                var wr = (HttpWebRequest)WebRequest.Create(url);
                WebResponse wresp = null;
                try
                {
                    string boundary = "---------------------------" + DateTime.Now.Ticks.ToString("x");
                    byte[] boundarybytes = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");

                    wr.Method = "POST";
                    wr.KeepAlive = true;
                    wr.Credentials = CredentialCache.DefaultCredentials;
                    wr.Accept = "text/html, application/xhtml+xml, */*";
                    wr.Referer = "https://octopart.com/bom-lookup/manage";
                    wr.Headers.Add(HttpRequestHeader.AcceptLanguage, "en-CA,en-US;q=0.7,en;q=0.3");
                    wr.UserAgent = "Octopart-Excel-Plugin/1.0 (+http://octopart.com/excel)";
                    wr.ContentType = "multipart/form-data; boundary=" + boundary;
                    wr.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip, deflate");
                    wr.Headers.Add(HttpRequestHeader.CacheControl, "no-cache");
                    wr.Expect = null;
                    wr.AllowAutoRedirect = false;

                    Stream rs = wr.GetRequestStream();

                    // START OF DATA
                    rs.Write(boundarybytes, 0, boundarybytes.Length);

                    // Data information
                    var headerTemplate =
                        "Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"\r\nContent-Type: {2}\r\n\r\n";
                    var header = string.Format(headerTemplate, paramName, Path.GetFileName(file), contentType);
                    var headerbytes = System.Text.Encoding.UTF8.GetBytes(header);
                    rs.Write(headerbytes, 0, headerbytes.Length);

                    // Include the file data
                    var fileBytes = File.ReadAllBytes(file);
                    rs.Write(fileBytes, 0, fileBytes.Length);

                    // END OF DATA
                    var trailer = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "--\r\n");
                    rs.Write(trailer, 0, trailer.Length);
                    rs.Close();

                    wresp = wr.GetResponse();
                    var resp = wresp.GetResponseStream();
                    if (resp != null)
                    {
                        var redirect = (new StreamReader(resp)).ReadToEnd();
                        var doc = XDocument.Parse(redirect);
                        var items = doc.Descendants("a");
                        var href = items.FirstOrDefault();
                        if (href != null)
                        {
                            // Launch the browser with the redirect url
                            System.Diagnostics.Process.Start(href.Value);
                            break;
                        }
                    }
                    else
                    {
                        Log.Debug("Invalid response");
                        return "Error Uploading BOM. Please ensure that you are connected to the internet.";
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex.Message);

                    if (ex.Message.ToLower().Contains("proxy"))
                    {
                        bool? result = false;
                        string user = string.Empty;
                        string pass = string.Empty;
                        var thread = new Thread(() =>
                        {
                            var tempDlg =
                                new ProxyAuthPrompt(wr.Proxy.GetProxy(new Uri("https://www.octopart.com")).AbsoluteUri);
                            result = tempDlg.ShowDialog();
                            if (result == true)
                            {
                                user = tempDlg.User;
                                pass = tempDlg.Pass;
                            }
                        });
                        thread.SetApartmentState(ApartmentState.STA);
                        thread.Start();
                        while (thread.IsAlive)
                        {
                            Thread.Sleep(1);
                        }

                        if (result == false)
                        {
                            return "Error Uploading BOM. Please ensure that you are connected to the internet, and that the correct proxy authentication is provided.";
                        }

                        if (!string.IsNullOrEmpty(user) && !string.IsNullOrEmpty(pass))
                        {
                            wr.Proxy.Credentials = new NetworkCredential(user, pass);
                        }
                    }
                }
                finally
                {
                    if (wresp != null)
                    {
                        wresp.Close();
                        wresp = null;
                    }

                    wr = null;
                }
            }

            return null;
        }
        #endregion

        #region Methods-Helper
        /// <summary>
        /// Find the best price given the price break
        /// </summary>
        /// <param name="desiredCurrency">The 3 character currency code</param>
        /// <param name="offer">The offer to search within</param>
        /// <param name="qty">The quantity to search for (i.e. QTY required)</param>
        /// <param name="ignoreMoq">Indicates if the MOQ should be considered when looking up pricing</param>
        /// <returns>The minimum price available for the specified QTY, in Current-Culture string format</returns>
        public static string OffersMinPrice(string desiredCurrency, ApiV4Schema.PartOffer offer, int qty, bool ignoreMoq)
        {
            if (string.IsNullOrEmpty(desiredCurrency) || (offer == null))
                return string.Empty;

            double priceMin = double.MaxValue;
            if (offer.prices.Keys.Contains(desiredCurrency))
            {
                string prices = offer.prices[desiredCurrency];
                if (!string.IsNullOrEmpty(prices))
                {
                    // Get price info
                    var priceInfo = JsonConvert.DeserializeObject<dynamic>(prices);

                    bool mouserSpecialCase = ((offer.seller.name == "Mouser") && (priceInfo.Count == 1) && (priceInfo[0][0] == 1));

                    for (int i = 0; (priceInfo != null) && (i < priceInfo.Count); i++)
                    {
                        try
                        {
                            int priceTupleQty = Convert.ToInt32(priceInfo[i][0], CultureInfo.CreateSpecificCulture("en-US"));
                            double priceTuplePrice = Convert.ToDouble(priceInfo[i][1], CultureInfo.CreateSpecificCulture("en-US"));

                            if (mouserSpecialCase)
                            {
                                if (qty >= priceTupleQty)
                                    priceMin = priceTuplePrice;
                            }
                            else if (qty >= priceTupleQty)
                            {
                                if (ignoreMoq || (string.IsNullOrEmpty(offer.moq) || (qty >= Convert.ToInt32(offer.moq, CultureInfo.CreateSpecificCulture("en-US")))))
                                {
                                    if (priceTuplePrice < priceMin)
                                        priceMin = priceTuplePrice;
                                }
                            }
                        }
                        catch (FormatException) { /* Do nothing */ }
                        catch (OverflowException) { /* Do nothing */ }
                    }
                }
            }

            if (priceMin == double.MaxValue)
                return string.Empty;
            else
                return priceMin.ToString("F5", CultureInfo.CurrentCulture);
        }

        /// <summary>
        /// Find the preferred currency, or whatever else
        /// </summary>
        /// <param name="desiredCurrency">The desired currency to get the price breaks from</param>
        /// <param name="offer">The offer to search within</param>
        /// <returns>Currency string</returns>
        public static string FindPreferredCurrency(string desiredCurrency, ApiV4Schema.PartOffer offer)
        {
            string ret = string.Empty;
            foreach (string currency in offer.prices.Keys)
            {
                string prices = offer.prices[currency];
                if (!string.IsNullOrEmpty(prices))
                {
                    if (currency == desiredCurrency)
                    {
                        // We found at least one price with the specified currency, so return it.
                        return desiredCurrency;
                    }
                    else
                    {
                        // Well, we didn't find what we were looking for, but at least give something.
                        ret = currency;
                    }
                }
            }

            return ret;
        }
        #endregion
    }
}
