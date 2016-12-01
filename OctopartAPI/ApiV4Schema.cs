using System.Collections.Generic;

namespace OctopartApi
{
    /// <summary>
    /// Provides an interface to the REST V4 API
    /// </summary>
    public class ApiV4Schema
    {
        public class Asset
        {
            public string url { get; set; }
            public string mimetype { get; set; }
        }

        public class Attribution
        {
            public List<Source> sources { get; set; }
            public string first_acquired { get; set; }
        }

        public class Brand
        {
            public string uid { get; set; }
            public string name { get; set; }
            public string homepage_url { get; set; }
        }

        public class BrokerListing
        {
            public Seller seller { get; set; }
            public string listing_url { get; set; }
            public string octopart_rfq_url { get; set; }
        }

        public class CADModel : Asset
        {
            public Attribution attribution { get; set; }
        }

        public class ComplianceDocument : Asset
        {
            // TODO: subtypes
            public Attribution attribution { get; set; }
        }

        public class Datasheet : Asset
        {
            public Attribution attribution { get; set; }
        }

        public class Description
        {
            public string value { get; set; }
            public Attribution attribution { get; set; }
        }

        public class ExternalLinks
        {
            public string product_url { get; set; }
            public string freesample_url { get; set; }
            public string evalkit_url { get; set; }
        }
        
        public class Imageset
        {
            public Asset swatch_image { get; set; }
            public Asset small_image { get; set; }
            public Asset medium_image { get; set; }
            public Asset large_image { get; set; }
            public Attribution attribution { get; set; }
        }

        public class Manufacturer
        {
            public string uid { get; set; }
            public string name { get; set; }
            public string homepage_url { get; set; }
        }

        public class Part
        {
            public string uid { get; set; }
            public string mpn { get; set; }
            public Manufacturer manufacturer { get; set; }
            public Brand brand { get; set; }
            public string octopart_url { get; set; }
            public ExternalLinks external_links { get; set; }
            public List<PartOffer> offers { get; set; }
            public BrokerListing broker_listings { get; set; }
            public string short_description { get; set; }
            public List<Description> descriptions { get; set; }
            public List<Imageset> imagesets { get; set; }
            public List<Datasheet> datasheets { get; set; }
            public List<ComplianceDocument> compliance_documents { get; set; }
            public List<ReferenceDesign> reference_designs { get; set; }
            public List<CADModel> cad_models { get; set; }
            public Document best_datasheet { get; set; }
        }

        public class Document
        {
            public string name { get; set; }
            public string url { get; set; }
        }

        public class PartOffer
        {
            public string sku { get; set; }
            public Seller seller { get; set; }
            public string product_url { get; set; }
            public string octopart_rfq_url { get; set; }
            public Dictionary<string, string> prices { get; set; }
            public int in_stock_quantity { get; set; }
            public string on_order_quantity { get; set; }
            public string on_order_eta { get; set; }
            public string factory_lead_days { get; set; }
            public string factory_order_multiple { get; set; }
            public string order_multiple { get; set; }
            public string moq { get; set; }
            public string packaging { get; set; }
            public bool is_authorized { get; set; }
            public string last_updated { get; set; }
        }

        public class PartsMatchQuery
        {
            public string brand { get; set; }
            public int limit { get; set; }
            public string mpn { get; set; }
            public string mpn_or_sku { get; set; }
            public string q { get; set; }
            public string reference { get; set; }
            public string seller { get; set; }
            public string sku { get; set; }
            public int start { get; set; }
        }

        public interface IPartsMatchQuery
        {
            string brand { get; set; }
            int limit { get; set; }
            string mpn { get; set; }
            string mpn_or_sku { get; set; }
            string q { get; set; }
            string reference { get; set; }
            string seller { get; set; }
            string sku { get; set; }
            int start { get; set; }
        }

        public class PartsMatchRequest
        {
            public bool exact_only { get; set; }
            public List<PartsMatchQuery> queries { get; set; }
        }

        public class PartsMatchResponse
        {
            public int msec { get; set; }
            public PartsMatchRequest request { get; set; }
            public List<PartsMatchResult> results { get; set; }
        }
        
        public class PartsMatchResult
        {
            public List<Part> items { get; set; }
            public int hits { get; set; }
            public string reference { get; set; }
            public string error { get; set; }
        }

        public class ReferenceDesign : Asset
        {
            public string title { get; set; }
            public string description { get; set; }
            public Attribution attribution { get; set; }
        }

        public class Seller
        {
            public string uid { get; set; }
            public string name { get; set; }
            public string homepage_url { get; set; }
            public string display_flag { get; set; }
            public bool has_ecommerce { get; set; }
        }

        public class Source
        {
            public string uid { get; set; }
            public string name { get; set; }
        }
    }
}
