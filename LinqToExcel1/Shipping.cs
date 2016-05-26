using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel.Attributes;

namespace LinqToExcel1 {
    class Shipping {

        [ExcelColumn("Search Key")]
        public String key { get; set; }

        [ExcelColumn("Ref 1")]
        public String ref1 { get; set; }

        [ExcelColumn("Ref 2")]
        public String ref2 { get; set; }

        [ExcelColumn("Ref 3")]
        public String ref3 { get; set; }

        [ExcelColumn("Tracking #")]
        public String tracking { get; set; }

        [ExcelColumn("Ship Date")]
        public String shipDate { get; set; }

        [ExcelColumn("Estimated $$")]
        public String estimatedCost { get; set; }

        [ExcelColumn("Company")]
        public String company { get; set; }

        [ExcelColumn("Attention")]
        public String attention { get; set;}

        [ExcelColumn("Address 1")]
        public String address1 { get; set; }

        [ExcelColumn("Address 2")]
        public String address2 { get; set; }

        [ExcelColumn("Address 3")]
        public String address3 { get; set; }

        [ExcelColumn("City")]
        public String city { get; set; }

        [ExcelColumn("State")]
        public String state { get; set; }

        [ExcelColumn("Postal")]
        public String postal { get; set; }

        [ExcelColumn("Country")]
        public String country { get; set; }

        [ExcelColumn("Residential")]
        public String residential { get; set; }

        [ExcelColumn("Phone")]
        public String phone { get; set; }

        [ExcelColumn("Email")]
        public String email { get; set; }

        [ExcelColumn("Service")]
        public String service { get; set; }

        [ExcelColumn("Signature Type")]
        public String signatureType { get; set; }

        [ExcelColumn("Weight")]
        public double weight { get; set; }

        [ExcelColumn("Declared Value")]
        public String declaredValue { get; set; }

        [ExcelColumn("Length")]
        public int length { get; set; }

        [ExcelColumn("Width")]
        public int width { get; set; }

        [ExcelColumn("Height")]
        public int height { get; set; }

        [ExcelColumn("Package Quantity")]
        public int packageQty { get; set; }

        [ExcelColumn("Payment Terms")]
        public String paymentTerms { get; set; }

        [ExcelColumn("Billing Account")]
        public String billingAcct { get; set; }

        [ExcelColumn("Billing Name")]
        public String billingName { get; set; }

        [ExcelColumn("Billing Adddress 1")]
        public String billingAddress1 { get; set; }

        [ExcelColumn("Billing Adddress 2")]
        public String billingAddress2 { get; set; }
        
        [ExcelColumn("Billing City")]
        public String billingCity { get; set; }

        [ExcelColumn("Billing State")]
        public String billingState { get; set; }

        [ExcelColumn("Billing Postal")]
        public String billingPostal { get; set; }

        [ExcelColumn("Billing Phone")]
        public String billingPhone { get; set; }

        [ExcelColumn("PROPER")]
        public String proper { get; set; }
    }
}
