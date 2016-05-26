using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using ExcelToLinq;
using excel = Microsoft.Office.Interop.Excel;
using LinqToExcel;
using Remotion.Data.Linq.Clauses;

namespace LinqToExcel1 {
    class Program {
        static StringBuilder csv = new StringBuilder();

        static void Main(string[] args) {
            /// String value is Spreadsheet's location on the machine
            var excel = new ExcelQueryFactory(@"C:\Users\John\Desktop\shippingtest.xlsx");

            /// Linq query
            /// String value is the name of the worksheet in the Excel document
            var shipping = from x in excel.Worksheet<Shipping>("2016.04 381305")
                           select x;

            /// Create header for CSV file
            /// Appended to string builder
            const string HEAD = "Search Key," + "Ref 1," + "Ref 2," + "Ref 3," + "Tracking #," + "Ship Date," +
                                "Estimated $$," + "Company," + "Attention," + "Address 1," + "Address 2," + "Address3," +
                                "City," + "State," + "Postal," + "Country," + "Residential," + "Phone," + "Email," +
                                "Service," + "Signature Type," + "Weight," + "Declared Value," + "Length," + "Width," +
                                "Height," + "Package Quantity," + "Payment Terms," + "Billing Account," + "Billing Name," +
                                "Billing Address1," + "Billing Address2," + "Billing City," + "Billing State," +
                                "Billing Postal," + "Billing Phone," + "PROPER";

            csv.AppendLine(HEAD);

            /// Integer for row count
            /// Not required
            int i = 1;
            foreach (Shipping name in shipping) {
                if (name.attention != null) {
                    name.attention = name.attention.Replace(",", "");
                }

                if (name.address1 != null) {
                    name.address1 = name.address1.Replace(",", "");
                }

                if (name.city != null) {
                    name.city = name.city.Replace(",", "");
                }

                /// Check if company name is null
                /// 
                /// This doesn't much help, but will prevent app from breaking
                /// when there is invalid data somewhere in the file
                if (name.company != null) {


                    if (name.company.Contains("M18-S") || name.company.Contains("M12-S")) {

                        name.weight = 9.5;
                        name.service = "firstclass";
                    }

                    if (name.company.Contains("M18-M") || name.company.Contains("M12-M")) {

                        name.weight = 10.00;
                        name.service = "firstclass";
                    }
                    if (name.company.Contains("M18-L") || name.company.Contains("M12-L")) {

                        name.weight = 10.50;
                        name.service = "firstclass";
                    }
                    if (name.company.Contains("M18-XL") || name.company.Contains("M12-XL")) {

                        name.weight = 11.20;
                        name.service = "firstclass";
                    }
                    if (name.company.Contains("M18-XXL") || name.company.Contains("M12-XXL")) {

                        name.weight = 12.30;
                        name.service = "firstclass";
                    }
                    if (name.company.Contains("M18-XXXL") || name.company.Contains("M12-XXXL")) {

                        name.weight = 14.00;
                        name.service = "parcelpost";
                    }
                }


                i++;


                /// Build line for a single record
                /// Will be appended to string builder
                /// The very definition of hacky
                var line = name.key + "," + name.ref1 + "," + name.ref2 + "," + name.ref3 + "," + name.tracking + "," +
                              name.shipDate + "," + name.estimatedCost + "," + name.company + "," +
                              name.attention + "," + name.address1 + "," + name.address2 + "," + name.address3 + "," +
                              name.city + "," + name.state + "," + name.postal + "," + name.country + "," + name.residential + "," +
                              name.phone + "," + name.email + "," + name.service + "," + name.signatureType + "," + name.weight +
                              "," + name.declaredValue + "," + name.length + "," +
                              name.width + "," + name.height + "," + name.packageQty + "," + name.paymentTerms + "," +
                              name.billingAcct + "," + name.billingName + "," + name.billingAddress1 + "," +
                              name.billingAddress2 + "," + name.billingCity + "," + name.billingState + "," +
                              name.billingPostal + "," + name.billingPhone + "," + name.proper;

                csv.AppendLine(line);
                Console.WriteLine(line);

                Console.WriteLine(i);

            }
            Console.Clear();
            Console.WriteLine("Process complete");
            Console.WriteLine("Total records: " + i);

            WriteCSV();

            Console.Read();


        }

        public static void WriteCSV() {
            string csvpath = @"C:\Users\John\Desktop\testing.csv";

            if (File.Exists(csvpath)) {
                File.Delete(csvpath);
            }

            using (StreamWriter sw = File.CreateText(csvpath)) {
                sw.Write(csv.ToString());

            }
        }
    }
}
