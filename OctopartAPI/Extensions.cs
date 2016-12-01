using System;
using System.Globalization;
using System.Linq;
using OctopartApi;

namespace ExtensionMethods
{
    public static class Extensions
    {
        /// <summary>
        /// Gets the url of the datasheet; returns first option if available
        /// </summary>
        /// <param name="part">The part as returned by the search</param>
        /// <returns>The url for the 'best' datasheet</returns>
        public static string GetDatasheetUrl(this ApiV4Schema.Part part)
        {
            if ((part != null) && (part.datasheets != null))
            {
                var datasheet = part.datasheets.FirstOrDefault(i => !string.IsNullOrEmpty(i.url));
                if (datasheet != null)
                {
                    // Success!
                    return datasheet.url;
                }
            }

            return "ERROR: Datasheet url not found. Please try expanding your search";
        }

        /// <summary>
        /// Find the best price given the price break
        /// </summary>
        /// <param name="offer">The offer to search within</param>
        /// <param name="currency">The 3 character currency code</param>
        /// <param name="qty">The quantity to search for (i.e. QTY required)</param>
        /// <returns>The minimum price available for the specified QTY</returns>
        public static double MinPrice(this ApiV4Schema.PartOffer offer, string currency, int qty)
        {
            // Force format optional arguments
            if (currency == string.Empty) currency = "USD";
            if (qty == 0) qty = 1;

            double minprice = double.MaxValue;

            try
            {
                string minpricestr = ApiV4.OffersMinPrice(currency, offer, qty, true);
                if (!string.IsNullOrEmpty(minpricestr))
                    minprice = Convert.ToDouble(minpricestr, CultureInfo.CurrentCulture);
            }
            catch (FormatException) { /* Do nothing */ }
            catch (OverflowException) { /* Do nothing */ }

            return minprice;
        }

        /// <summary>
        /// Calculate offer MOQ as maximum of moq and order_multiple fields in offer dict.
        /// </summary>
        /// <param name="offer">The offer to find the MOQ from</param>
        /// <returns>The MOQ, or -1 if not found</returns>
        public static int GetRealMoq(this ApiV4Schema.PartOffer offer)
        {
            int moq = -1;
            int order_multiple = -1;

            try
            {
                if (!string.IsNullOrEmpty(offer.moq)) moq = Convert.ToInt32(offer.moq, CultureInfo.CreateSpecificCulture("en-US"));
                if (!string.IsNullOrEmpty(offer.order_multiple)) Convert.ToInt32(offer.order_multiple, CultureInfo.CreateSpecificCulture("en-US"));
            }
            catch (FormatException) { /* Do nothing */ }
            catch (OverflowException) { /* Do nothing */ }

            if (order_multiple != -1)
            {
                if ((moq == -1) || (moq < order_multiple))
                {
                    moq = order_multiple;
                }
            }

            return moq;
        }

        /// <summary>
        /// Convert a string based factory lead days to an int. 
        /// </summary>
        /// <param name="offer">The offer to find the MOQ from</param>
        /// <returns>The factory lead days, or Int.Max if not found</returns>
        public static int GetRealFactoryLeadDays(this ApiV4Schema.PartOffer offer)
        {
            int factory_lead_days = int.MaxValue;

            try
            {
                if (!string.IsNullOrEmpty(offer.factory_lead_days)) factory_lead_days = Convert.ToInt32(offer.factory_lead_days, CultureInfo.CreateSpecificCulture("en-US"));
            }
            catch (FormatException) { /* Do nothing */ }
            catch (OverflowException) { /* Do nothing */ }

            return factory_lead_days;
        }

        public static int GetRealOrderMultiple(this ApiV4Schema.PartOffer offer)
        {
            int order_multiple = 1;

            try
            {
                if (!string.IsNullOrEmpty(offer.order_multiple)) order_multiple = Convert.ToInt32(offer.order_multiple, CultureInfo.CreateSpecificCulture("en-US"));
            }
            catch (FormatException) { /* Do nothing */ }
            catch (OverflowException) { /* Do nothing */ }

            return order_multiple;
        }
    }
}
