using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace OneS_Baselinker
{
        public partial class Baselinker_response
        {
            [JsonProperty("status")]
            public string Status { get; set; }

            [JsonProperty("orders")]
            public Order[] Orders { get; set; }
        }
        public partial class Baselinker_journal_response
        {
            [JsonProperty("status")]
            public string Status { get; set; }

            [JsonProperty("logs")]
            public Journal[] Logs { get; set; }
        }
    public partial class Journal
        {
        [JsonProperty("log_id")]
        public string LogId { get; set; }

        [JsonProperty("log_type")]
        public string LogType { get; set; }

        [JsonProperty("order_id")]
        public string OrderId { get; set; }

        [JsonProperty("object_id")]
        public string ObjectId { get; set; }

        [JsonProperty("date")]
        public string Date { get; set; }
    }
        public partial class Order
        {
            [JsonProperty("order_id")]
            public string OrderId { get; set; }

            [JsonProperty("shop_order_id")]
            public string ShopOrderId { get; set; }

            [JsonProperty("external_order_id")]
            public string ExternalOrderId { get; set; }

            [JsonProperty("order_source")]
            public string OrderSource { get; set; }

            [JsonProperty("order_source_id")]
            public string OrderSourceId { get; set; }

            [JsonProperty("order_source_info")]
            public string OrderSourceInfo { get; set; }

            [JsonProperty("order_status_id")]
            public string OrderStatusId { get; set; }

            [JsonProperty("date_add")]
            public string DateAdd { get; set; }

            [JsonProperty("date_confirmed")]
            public string DateConfirmed { get; set; }

            [JsonProperty("date_in_status")]
            public string DateInStatus { get; set; }

            [JsonProperty("user_login")]
            public string UserLogin { get; set; }

            [JsonProperty("phone")]
            public string Phone { get; set; }

            [JsonProperty("email")]
            public string Email { get; set; }

            [JsonProperty("user_comments")]
            public string UserComments { get; set; }

            [JsonProperty("admin_comments")]
            public string AdminComments { get; set; }

            [JsonProperty("currency")]
            public string Currency { get; set; }

            [JsonProperty("payment_method")]
            public string PaymentMethod { get; set; }

            [JsonProperty("payment_method_cod")]
            public string PaymentMethodCod { get; set; }

            [JsonProperty("payment_done")]
            public double PaymentDone { get; set; }

            [JsonProperty("payment_date")]
            public string PaymentDate { get; set; }

            [JsonProperty("delivery_method")]
            public string DeliveryMethod { get; set; }

            [JsonProperty("delivery_price")]
            public double DeliveryPrice { get; set; }

            [JsonProperty("delivery_package_module")]
            public string DeliveryPackageModule { get; set; }

            [JsonProperty("delivery_package_nr")]
            public string DeliveryPackageNr { get; set; }

            [JsonProperty("delivery_fullname")]
            public string DeliveryFullname { get; set; }

            [JsonProperty("delivery_company")]
            public string DeliveryCompany { get; set; }

            [JsonProperty("delivery_address")]
            public string DeliveryAddress { get; set; }

            [JsonProperty("delivery_city")]
            public string DeliveryCity { get; set; }

            [JsonProperty("delivery_postcode")]
            public string DeliveryPostcode { get; set; }

            [JsonProperty("delivery_country")]
            public string DeliveryCountry { get; set; }

            [JsonProperty("delivery_point_id")]
            public string DeliveryPointId { get; set; }

            [JsonProperty("delivery_point_name")]
            public string DeliveryPointName { get; set; }

            [JsonProperty("delivery_point_address")]
            public string DeliveryPointAddress { get; set; }

            [JsonProperty("delivery_point_postcode")]
            public string DeliveryPointPostcode { get; set; }

            [JsonProperty("delivery_point_city")]
            public string DeliveryPointCity { get; set; }

            [JsonProperty("invoice_fullname")]
            public string InvoiceFullname { get; set; }

            [JsonProperty("invoice_company")]
            public string InvoiceCompany { get; set; }

            [JsonProperty("invoice_nip")]
            public string InvoiceNip { get; set; }

            [JsonProperty("invoice_address")]
            public string InvoiceAddress { get; set; }

            [JsonProperty("invoice_city")]
            public string InvoiceCity { get; set; }

            [JsonProperty("invoice_postcode")]
            public string InvoicePostcode { get; set; }

            [JsonProperty("invoice_country")]
            public string InvoiceCountry { get; set; }

            [JsonProperty("want_invoice")]
            public string WantInvoice { get; set; }

            [JsonProperty("extra_field_1")]
            public string ExtraField1 { get; set; }

            [JsonProperty("extra_field_2")]
            public string ExtraField2 { get; set; }

            [JsonProperty("order_page")]
            public string OrderPage { get; set; }

            [JsonProperty("pick_status")]
            public string PickStatus { get; set; }

            [JsonProperty("pack_status")]
            public string PackStatus { get; set; }

            [JsonProperty("products")]
            public Product[] Products { get; set; }
        }
        public partial class InputParameters
        {
            public string date_confirmed_from { get; set; }

        }
        public partial class JournalParameters
        {
            public string[] logs_types { get; set; }
            public string order_id { get; set; }
         }
    public partial class Product
        {
            [JsonProperty("storage")]
            public string Storage { get; set; }

            [JsonProperty("storage_id")]
            public string StorageId { get; set; }

            [JsonProperty("order_product_id")]
            public string OrderProductId { get; set; }

            [JsonProperty("product_id")]
            public string ProductId { get; set; }

            [JsonProperty("variant_id")]
            public string VariantId { get; set; }

            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("attributes")]
            public string Attributes { get; set; }

            [JsonProperty("sku")]
            public string Sku { get; set; }

            [JsonProperty("ean")]
            public string Ean { get; set; }

            [JsonProperty("auction_id")]
            public string AuctionId { get; set; }

            [JsonProperty("price_brutto")]
            public double PriceBrutto { get; set; }

            [JsonProperty("tax_rate")]
            public string TaxRate { get; set; }

            [JsonProperty("quantity")]
            public string Quantity { get; set; }

            [JsonProperty("weight")]
            public string Weight { get; set; }
        }

        public partial class Baselinker_response
        {
            public static Baselinker_response FromJson(string json) => JsonConvert.DeserializeObject<Baselinker_response>(json);
        }
        public partial class Baselinker_journal_response
        {
            public static Baselinker_journal_response FromJson(string json) => JsonConvert.DeserializeObject<Baselinker_journal_response>(json);
        }

    public static class Serialize
        {
            public static string ToJson(this Baselinker_response self) => JsonConvert.SerializeObject(self);
        }
}
