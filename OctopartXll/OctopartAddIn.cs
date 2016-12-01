using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using ExcelDna.Integration;
using ExtensionMethods;
using OctopartApi;

namespace OctopartXll
{
    public class OctopartAddIn : IExcelAddIn
    {
        #region Variables
        /// <summary>
        /// A collection of Options that configure the Octopart Add-In
        /// </summary>
        private static readonly Dictionary<string, dynamic> Options = new Dictionary<string, dynamic>()
        {
            {"log", false}
        };

        /// <summary>
        /// Handles all queries to the octopart website, and caches the results for the Excel session
        /// </summary>
        public static readonly OctopartQueryManager QueryManager = new OctopartQueryManager();

        /// <summary>
        /// This 'refresh' hack is used to trick excel to refresh the information from the QueryManager
        /// </summary>
        private static string _refreshhack = string.Empty;

        /// <summary>
        /// Log (for debugging or otherwise)
        /// </summary>
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        #endregion

        #region Constructors
        /// <summary>
        /// This function is called when the XLL is installed into an Excel spreadsheet
        /// </summary>
        public void AutoOpen()
        {
            try
            {
                ExcelIntegration.RegisterUnhandledExceptionHandler(
                    ex => "!!! EXCEPTION: " + ex.ToString());
            }
            catch (Exception ex)
            {
                log.Fatal(ex.ToString());
            }
        }

        /// <summary>
        /// Gracefully detaches the XLL
        /// </summary>
        public void AutoClose()
        { }
        #endregion

        #region ExcelUdfs
        [ExcelFunction(Category = "Octopart Queries", Description = "Gets the Octopart Details URL", HelpTopic = "OctopartAddIn.chm!1001")]
        public static object OCTOPART_DETAIL_URL(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part (optional)", Name = "Manufacturer")] string manuf = "")
        {
            // Check to see if a cached version is available (only checks non-error'd queries)
            ApiV4Schema.Part part = GetManuf(mpn_or_sku, manuf);
            if (part != null)
                return part.octopart_url;

            // Excel's recalculation engine is based on INPUTS. The main function will be called if:
            // - Inputs are changed
            // - Input cells location are changed (i.e., moving the cell, or deleting a row that impacts this cell)
            // However, the async function will ONLY be run if the inputs are DIFFERENT.
            //   The impact of this is that if last time the function was run it returned an error that was unrelated to the inputs 
            //   (i.e., invalid ApiKey, network was down, etc), then the function would not run again.
            // To fix that issue, whitespace padding is added to the mpn_or_sku. This whitespace is removed anyway, so it has no 
            // real impact other than to generate a refresh.
            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
#if !TEST
            object asyncResult = ExcelAsyncUtil.Run("OCTOPART_DETAIL_URL", new object[] { mpn_or_sku, manuf }, delegate
            {
#endif
            try
            {
                part = SearchAndWaitPart(mpn_or_sku, manuf);
                if (part == null)
                {
                    string err = QueryManager.GetLastError(mpn_or_sku);
                    if (string.IsNullOrEmpty(err))
                        err = "Query did not provide a result. Please widen your search criteria.";

                    return "ERROR: " + err;
                }

                return part.octopart_url;
            }
            catch (Exception ex)
            {
                log.Fatal(ex.ToString());
                return "ERROR: " + OctopartQueryManager.FATAL_ERROR;
            }
#if !TEST
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                // Still processing...
                return "!!! Processing !!!";
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (part == null))
            {
                // Regenerate the hack value if an error was received
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }
            
            // Done processing!
            return asyncResult;
#endif
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Gets the Octopart Datasheet URL", HelpTopic = "OctopartAddIn.chm!1002")]
        public static object OCTOPART_DATASHEET_URL(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "")
        {
            ApiV4Schema.Part part = GetManuf(mpn_or_sku, manuf);
            if (part != null)
            {
                // ---- BEGIN Function Specific Information ----
                return part.GetDatasheetUrl();
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("OCTOPART_DATASHEET_URL", new object[] { mpn_or_sku, manuf }, delegate
            {
                try
                {
                    part = SearchAndWaitPart(mpn_or_sku, manuf);
                    if (part == null)
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return part.GetDatasheetUrl();
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + OctopartQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return "!!! Processing !!!";
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (part == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Gets the distributor price from Octopart", HelpTopic = "OctopartAddIn.chm!1003")]
        public static object OCTOPART_DISTRIBUTOR_PRICE(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributors for lookup (optional)", Name = "Distributor(s)")] Object[] distributors = null,
            [ExcelArgument(Description = "Quantity for lookup (optional, default = 1)", Name = "Quantity")] int qty = 1,
            [ExcelArgument(Description = "Currency for lookup (optional, default = USD). Standard currency codes apply (http://en.wikipedia.org/wiki/ISO_4217)", Name = "Currency")] string currency = "USD")
        {
            List<ApiV4Schema.PartOffer> offers = GetOffers(mpn_or_sku, manuf, GetDistributors(distributors));
            if ((offers != null) && (offers.Count > 0))
            {
                // ---- BEGIN Function Specific Information ----
                double minprice = offers.Min(offer => offer.MinPrice(currency, qty));
                if (minprice < double.MaxValue)
                    return minprice;
                else
                    return "ERROR: Query did not provide a result. Please widen your search criteria.";
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
#if !TEST
            object asyncResult = ExcelAsyncUtil.Run("OCTOPART_DISTRIBUTOR_PRICE", new object[] { mpn_or_sku, manuf, distributors, qty, currency }, delegate
            {
#endif
            try
            {
                offers = SearchAndWaitOffers(mpn_or_sku, manuf, GetDistributors(distributors));
                if ((offers == null) || (offers.Count == 0))
                {
                    string err = QueryManager.GetLastError(mpn_or_sku);
                    if (string.IsNullOrEmpty(err))
                        err = "Query did not provide a result. Please widen your search criteria.";

                    return "ERROR: " + err;
                }

                // ---- BEGIN Function Specific Information ----
                double minprice = offers.Min(offer => offer.MinPrice(currency, qty));
                if (minprice < double.MaxValue)
                    return minprice;
                else
                    return "ERROR: Query did not provide a result. Please widen your search criteria.";
                // ---- END Function Specific Information ----
            }
            catch (Exception ex)
            {
                log.Fatal(ex.ToString());
                return "ERROR: " + OctopartQueryManager.FATAL_ERROR;
            }
#if !TEST
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return "!!! Processing !!!";
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
#endif
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Gets the average price from Octopart.com", HelpTopic = "OctopartAddIn.chm!1004")]
        public static object OCTOPART_AVERAGE_PRICE(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Quantity for lookup (optional, default = 1)", Name = "Quantity")] int qty = 1,
            [ExcelArgument(Description = "Currency for lookup (optional, default = USD). Standard currency codes apply (http://en.wikipedia.org/wiki/ISO_4217)", Name = "Currency")] string currency = "USD")
        {
            List<ApiV4Schema.PartOffer> offers = GetOffers(mpn_or_sku, manuf);
            if ((offers != null) && (offers.Count > 0))
            {
                // ---- BEGIN Function Specific Information ----
                offers = offers.Where(offer => offer.MinPrice(currency, qty) < double.MaxValue).ToList();
                if ((offers != null) && (offers.Count > 0))
                {
                    double price = offers.Average(offer => offer.MinPrice(currency, qty));
                    if (price < double.MaxValue)
                        return price;
                    else
                        return "ERROR: Query did not provide a result. Please widen your search criteria.";
                }
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
#if !TEST
            object asyncResult = ExcelAsyncUtil.Run("OCTOPART_AVERAGE_PRICE", new object[] { mpn_or_sku, manuf, qty, currency }, delegate
            {
#endif
            try
            {
                offers = SearchAndWaitOffers(mpn_or_sku, manuf);
                if ((offers == null) || (offers.Count == 0))
                {
                    string err = QueryManager.GetLastError(mpn_or_sku);
                    if (string.IsNullOrEmpty(err))
                        err = "Query did not provide a result. Please widen your search criteria.";

                    return "ERROR: " + err;
                }

                // ---- BEGIN Function Specific Information ----
                offers = offers.Where(offer => offer.MinPrice(currency, qty) < double.MaxValue).ToList();
                if ((offers != null) && (offers.Count > 0))
                {
                    double price = offers.Average(offer => offer.MinPrice(currency, qty));
                    if (price < double.MaxValue)
                        return price;
                }

                return "ERROR: Query did not provide a result. Please widen your search criteria.";
                // ---- END Function Specific Information ----
            }
            catch (Exception ex)
            {
                log.Fatal(ex.ToString());
                return "ERROR: " + OctopartQueryManager.FATAL_ERROR;
            }
#if !TEST
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return "!!! Processing !!!";
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
#endif
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Gets the distributor stock quantity from Octopart.com", HelpTopic = "OctopartAddIn.chm!1005")]
        public static object OCTOPART_DISTRIBUTOR_STOCK(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributors for lookup (optional)", Name = "Distributor(s)")] Object[] distributors = null)
        {
            List<ApiV4Schema.PartOffer> offers = GetOffers(mpn_or_sku, manuf, GetDistributors(distributors));
            if ((offers != null) && (offers.Count > 0))
            {
                // ---- BEGIN Function Specific Information ----
                int stock = offers.Max(offer => offer.in_stock_quantity);
                switch (stock)
                {
                    case -1: return "Non-stocked";
                    case -2: return "Yes";
                    case -3: return "Unknown";
                    case -4: return "RFQ";
                    default: return stock;
                }
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("OCTOPART_DISTRIBUTOR_STOCK", new object[] { mpn_or_sku, manuf, distributors }, delegate
            {
                try
                {
                    offers = SearchAndWaitOffers(mpn_or_sku, manuf, GetDistributors(distributors));
                    if ((offers == null) || (offers.Count == 0))
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    int stock = offers.Max(offer => offer.in_stock_quantity);
                    switch (stock)
                    {
                        case -1: return "Non-stocked";
                        case -2: return "Yes";
                        case -3: return "Unknown";
                        case -4: return "RFQ";
                        default: return stock;
                    }
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + OctopartQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return "!!! Processing !!!";
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Gets the distributor MOQ from Octopart.com", HelpTopic = "OctopartAddIn.chm!1006")]
        public static object OCTOPART_DISTRIBUTOR_MOQ(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributors for lookup (optional)", Name = "Distributor(s)")] Object[] distributors = null)
        {
            List<ApiV4Schema.PartOffer> offers = GetOffers(mpn_or_sku, manuf, GetDistributors(distributors));
            if ((offers != null) && (offers.Count > 0))
            {
                // ---- BEGIN Function Specific Information ----
                offers = offers.Where(offer => offer.GetRealMoq() > 0).ToList();
                if ((offers != null) && (offers.Count > 0))
                {
                    return offers.Min(offer => offer.GetRealMoq());
                }
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("OCTOPART_DISTRIBUTOR_MOQ", new object[] { mpn_or_sku, manuf, distributors }, delegate
            {
                try
                {
                    offers = SearchAndWaitOffers(mpn_or_sku, manuf, GetDistributors(distributors));
                    if ((offers == null) || (offers.Count == 0))
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    offers = offers.Where(offer => offer.GetRealMoq() > 0).ToList();
                    if ((offers != null) && (offers.Count > 0))
                    {
                        return offers.Min(offer => offer.GetRealMoq());
                    }
                    else
                    {
                        string err = "Query did not provide a result. Please widen your search criteria.";
                        return "ERROR: " + err;
                    }
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + OctopartQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return "!!! Processing !!!";
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Gets the distributor order multiple from Octopart.com", HelpTopic = "OctopartAddIn.chm!1007")]
        public static object OCTOPART_DISTRIBUTOR_ORDER_MULTIPLE(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributors for lookup (optional)", Name = "Distributor(s)")] Object[] distributors = null)
        {
            List<ApiV4Schema.PartOffer> offers = GetOffers(mpn_or_sku, manuf, GetDistributors(distributors));
            if ((offers != null) && (offers.Count > 0))
            {
                // ---- BEGIN Function Specific Information ----
                return offers.Min(offer => offer.GetRealOrderMultiple());
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("OCTOPART_DISTRIBUTOR_ORDER_MULTIPLE", new object[] { mpn_or_sku, manuf, distributors }, delegate
            {
                try
                {
                    offers = SearchAndWaitOffers(mpn_or_sku, manuf, GetDistributors(distributors));
                    if ((offers == null) || (offers.Count == 0))
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return offers.Min(offer => offer.GetRealOrderMultiple());
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + OctopartQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return "!!! Processing !!!";
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Gets the distributor lead time from Octopart.com", HelpTopic = "OctopartAddIn.chm!1008")]
        public static object OCTOPART_DISTRIBUTOR_LEAD_TIME(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributors for lookup (optional)", Name = "Distributor(s)")] Object[] distributors = null)
        {
            List<ApiV4Schema.PartOffer> offers = GetOffers(mpn_or_sku, manuf, GetDistributors(distributors));
            if ((offers != null) && (offers.Count > 0))
            {
                // ---- BEGIN Function Specific Information ----
                return offers.Min(offer => offer.GetRealFactoryLeadDays());
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("OCTOPART_DISTRIBUTOR_LEAD_TIME", new object[] { mpn_or_sku, manuf, distributors }, delegate
            {
                try
                {
                    offers = SearchAndWaitOffers(mpn_or_sku, manuf, GetDistributors(distributors));
                    if ((offers == null) || (offers.Count == 0))
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return offers.Min(offer => offer.GetRealFactoryLeadDays());
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + OctopartQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return "!!! Processing !!!";
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Gets the distributor packaging style from Octopart.com", HelpTopic = "OctopartAddIn.chm!1009")]
        public static object OCTOPART_DISTRIBUTOR_PACKAGING(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributor for lookup (optional)", Name = "Distributor")] string distributor = "")
        {
            ApiV4Schema.PartOffer offer = GetOffer(mpn_or_sku, manuf, distributor);
            if (offer != null)
            {
                // ---- BEGIN Function Specific Information ----
                return offer.packaging ?? string.Empty;
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("OCTOPART_DISTRIBUTOR_PACKAGING", new object[] { mpn_or_sku, manuf, distributor }, delegate
            {
                try
                {
                    offer = SearchAndWaitOffer(mpn_or_sku, manuf, distributor);
                    if (offer == null)
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return offer.packaging ?? string.Empty;
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + OctopartQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return "!!! Processing !!!";
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offer == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Gets the distributor url from Octopart.com", HelpTopic = "OctopartAddIn.chm!1010")]
        public static object OCTOPART_DISTRIBUTOR_URL(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributor for lookup (optional)", Name = "Distributor")] string distributor = "")
        {
            ApiV4Schema.PartOffer offer = GetOffer(mpn_or_sku, manuf, distributor);
            if (offer != null)
            {
                // ---- BEGIN Function Specific Information ----
                return offer.product_url;
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("OCTOPART_DISTRIBUTOR_URL", new object[] { mpn_or_sku, manuf, distributor }, delegate
            {
                try
                {
                    offer = SearchAndWaitOffer(mpn_or_sku, manuf, distributor);
                    if (offer == null)
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return offer.product_url;
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + OctopartQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return "!!! Processing !!!";
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offer == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Gets the distributor SKU from Octopart.com", HelpTopic = "OctopartAddIn.chm!1011")]
        public static object OCTOPART_DISTRIBUTOR_SKU(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributor for lookup (optional)", Name = "Distributor")] string distributor = "")
        {
            ApiV4Schema.PartOffer offer = GetOffer(mpn_or_sku, manuf, distributor);
            if (offer != null)
            {
                // ---- BEGIN Function Specific Information ----
                return offer.sku;
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("OCTOPART_DISTRIBUTOR_SKU", new object[] { mpn_or_sku, manuf, distributor }, delegate
            {
                try
                {
                    offer = SearchAndWaitOffer(mpn_or_sku, manuf, distributor);
                    if (offer == null)
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return offer.sku;
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + OctopartQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return "!!! Processing !!!";
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offer == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Sets Options for Octopart Queries", HelpTopic = "OctopartAddIn.chm!1101", IsVolatile = true)]
        public static object OCTOPART_SET_OPTIONS(
            [ExcelArgument(Description = "Options (separated by ',')")] string optionStr)
        {
            var optionList = optionStr.ToLower().Split(',');
            string result = "The following options were set:";

            if (optionList.Contains("enable log"))
                Options["log"] = true;
            if (optionList.Contains("disable log"))
                Options["log"] = false;

            var temp = optionList.FirstOrDefault(i => i.Contains("querytimeout"));
            if (temp != null)
            {
                var param = temp.Split('=');
                if (param.Count() > 1)
                {
                    try
                    {
                        int timeout = Convert.ToInt32(param[1]);
                        QueryManager.HttpTimeout = timeout;
                        Options["querytimeout"] = timeout;
                        result += " QueryTimeout(" + timeout + ")";
                    }
                    catch
                    {
                        result += " QueryTimeout(Format Error)";
                        log.Info(string.Format("Invalid value for QueryTimeout ({0})", param[1]));
                    }
                }
            }

            temp = optionList.FirstOrDefault(i => i.Contains("apikey"));
            if (temp != null)
            {
                var param = temp.Split('=');
                if ((param.Count() > 1) && !string.IsNullOrEmpty(param[1]))
                {
                    string key = param[1];
                    Options["apikey"] = key;
                    QueryManager.ApiKey = key;
                    result += " ApiKey(" + key + ")";
                }
            }

            return result;
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Get Information about the Octopart Add-In", HelpTopic = "OctopartAddIn.chm!1102", IsVolatile = true)]
        public static object OCTOPART_GET_INFO()
        {
            return "Version: " + Assembly.GetExecutingAssembly().GetName().Version;
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "[Obsolete] Sets the user's email address for Octopart Queries", HelpTopic = "OctopartAddIn.chm!1103", IsVolatile = true)]
        public static object OCTOPART_SET_USER(
            [ExcelArgument(Description = "User email address", Name = "Email Address")] string email)
        {
            return "[Obsolete] Please use =OCTOPART_SET_APIKEY(...)";
        }

        [ExcelFunction(Category = "Octopart Queries", Description = "Sets the user's API Key for Octopart Queries", HelpTopic = "OctopartAddIn.chm!1103", IsVolatile = true)]
        public static object OCTOPART_SET_APIKEY(
            [ExcelArgument(Description = "User API Key", Name = "API Key")] string api_key)
        {
            if (string.IsNullOrEmpty(api_key))
            {
                return "Please sign-in to Octopart and visit https://www.octopart.com/api/home to register for an API key or https://www.octopart.com/api/dashboard to view your current API key.";
            }

            if (QueryManager.ApiKey != api_key)
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
                QueryManager.ApiKey = api_key;
            }

            return "Octopart Add-In is ready!";
        }
        #endregion

        #region Methods
        private static ApiV4Schema.Part SearchAndWaitPart(string mpnOrSku, string manuf)
        {
            // Check for errors. If one exists, force a refresh
            if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpnOrSku)))
                QueryManager.QueryNext(mpnOrSku);

            List<ApiV4Schema.Part> parts = QueryManager.GetParts(mpnOrSku);
            for (int i = 0; i < 10000 && !QueryManager.IsQueryLimitMaxed(mpnOrSku) && (parts.Count == 0); i++, Thread.Sleep(1))
            {
                parts = QueryManager.GetParts(mpnOrSku);
                if (parts.Count(item => string.IsNullOrEmpty(manuf) || item.manufacturer.name.Sanitize().Contains(manuf.Sanitize()) || item.brand.name.Sanitize().Contains(manuf.Sanitize())) == 0)
                {
                    if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpnOrSku)))
                        break;
                    QueryManager.QueryNext(mpnOrSku);
                }
            }

            return GetManuf(mpnOrSku, manuf);
        }

        private static ApiV4Schema.PartOffer SearchAndWaitOffer(string mpn_or_sku, string manuf, string distributor = "")
        {
            // Check for errors. If one exists, force a refresh
            if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                QueryManager.QueryNext(mpn_or_sku);

            for (int i = 0; i < 1000 && !QueryManager.IsQueryLimitMaxed(mpn_or_sku); i++, Thread.Sleep(10))
            {
                List<ApiV4Schema.Part> parts = QueryManager.GetParts(mpn_or_sku);
                if (parts.Count(item => string.IsNullOrEmpty(manuf) || item.manufacturer.name.Sanitize().Contains(manuf.Sanitize()) || item.brand.name.Sanitize().Contains(manuf.Sanitize())) == 0)
                {
                    if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                        break;
                    QueryManager.QueryNext(mpn_or_sku);
                }
                else if (string.IsNullOrEmpty(distributor))
                {
                    break;
                }
                else
                {
                    // Search for specified distributor
                    var offers = parts.SelectMany(offer => offer.offers).ToList();
                    if (offers.Count(offer => offer.seller.name.Sanitize().Contains(distributor.Sanitize())) == 0)
                    {
                        if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                            break;
                        QueryManager.QueryNext(mpn_or_sku);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return GetOffer(mpn_or_sku, manuf, distributor);
        }

        private static List<ApiV4Schema.PartOffer> SearchAndWaitOffers(string mpn_or_sku, string manuf, string distributor = "")
        {
            // Check for errors. If one exists, force a refresh
            if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                QueryManager.QueryNext(mpn_or_sku);

            for (int i = 0; i < 1000 && !QueryManager.IsQueryLimitMaxed(mpn_or_sku); i++, Thread.Sleep(10))
            {
                List<ApiV4Schema.Part> parts = QueryManager.GetParts(mpn_or_sku);
                if (parts.Count(item => string.IsNullOrEmpty(manuf) || item.manufacturer.name.Sanitize().Contains(manuf.Sanitize()) || item.brand.name.Sanitize().Contains(manuf.Sanitize())) == 0)
                {
                    if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                        break;
                    QueryManager.QueryNext(mpn_or_sku);
                }
                else if (string.IsNullOrEmpty(distributor))
                {
                    break;
                }
                else
                {
                    // Search for specified distributor
                    var offers = parts.SelectMany(offer => offer.offers).ToList();
                    if (offers.Count(offer => offer.seller.name.Sanitize().Contains(distributor.Sanitize())) == 0)
                    {
                        if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                            break;
                        QueryManager.QueryNext(mpn_or_sku);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return GetOffers(mpn_or_sku, manuf, distributor);
        }

        private static List<ApiV4Schema.PartOffer> SearchAndWaitOffers(string mpn_or_sku, string manuf, List<string> distributor)
        {
            // Check for errors. If one exists, force a refresh
            if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                QueryManager.QueryNext(mpn_or_sku);

            for (int i = 0; i < 1000 && !QueryManager.IsQueryLimitMaxed(mpn_or_sku); i++, Thread.Sleep(10))
            {
                List<ApiV4Schema.Part> parts = QueryManager.GetParts(mpn_or_sku);
                if (parts.Count(item => string.IsNullOrEmpty(manuf) || item.manufacturer.name.Sanitize().Contains(manuf.Sanitize()) || item.brand.name.Sanitize().Contains(manuf.Sanitize())) == 0)
                {
                    if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                        break;
                    QueryManager.QueryNext(mpn_or_sku);
                }
                else if (distributor == null || distributor.Count == 0)
                {
                    break;
                }
                else
                {
                    // Search for specified distributor
                    var offers = parts.SelectMany(offer => offer.offers).ToList();
                    if (offers.Count(offer => distributor.Any(d => offer.seller.name.Sanitize().Contains(d.Sanitize()))) == 0)
                    {
                        if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                            break;
                        QueryManager.QueryNext(mpn_or_sku);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return GetOffers(mpn_or_sku, manuf, distributor);
        }

        private static ApiV4Schema.Part GetManuf(string mpnOrSku, string manuf)
        {
            List<ApiV4Schema.Part> parts = QueryManager.GetParts(mpnOrSku);
            return parts.FirstOrDefault(item => string.IsNullOrEmpty(manuf) || item.manufacturer.name.Sanitize().Contains(manuf.Sanitize()) || item.brand.name.Sanitize().Contains(manuf.Sanitize()));
        }

        private static ApiV4Schema.PartOffer GetOffer(string mpnOrSku, string manuf, string distributor = "")
        {
            List<ApiV4Schema.Part> parts = QueryManager.GetParts(mpnOrSku);
            List<ApiV4Schema.PartOffer> offers = parts.SelectMany(offer => offer.offers).ToList();
            return offers.FirstOrDefault(offer => string.IsNullOrEmpty(distributor) || offer.seller.name.Sanitize().Contains(distributor.Sanitize()));
        }

        private static List<ApiV4Schema.PartOffer> GetOffers(string mpnOrSku, string manuf, string distributor = "")
        {
            List<ApiV4Schema.Part> parts = QueryManager.GetParts(mpnOrSku);
            List<ApiV4Schema.PartOffer> offers = parts.SelectMany(offer => offer.offers).ToList();
            return offers.Where(offer => string.IsNullOrEmpty(distributor) || offer.seller.name.Sanitize().Contains(distributor.Sanitize())).ToList();
        }

        private static List<ApiV4Schema.PartOffer> GetOffers(string mpnOrSku, string manuf, List<string> distributors)
        {
            List<ApiV4Schema.Part> parts = QueryManager.GetParts(mpnOrSku);
            List<ApiV4Schema.PartOffer> offers = parts.SelectMany(offer => offer.offers).ToList();
            return offers.Where(
                offer => distributors == null || distributors.Count == 0 || distributors.Any(d => offer.seller.name.Sanitize().Contains(d.Sanitize()))
            ).ToList();
        }

        private static List<string> GetDistributors(Object[] distributors)
        {
            List<string> cleanDistributors = distributors.ToList()
                .Where(x => x != ExcelMissing.Value && x != ExcelEmpty.Value)
                .Select(x => x.ToString())
                .ToList();
            return cleanDistributors;
        }
        #endregion
    }
}