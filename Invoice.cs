using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace SageIntegration
{
    public class Invoice
    {
        public List<LineItem> LineItems = new List<LineItem>();

        public string EDI_ID = string.Empty;
        public string SendingQualifier = string.Empty;
        public string SendingID = string.Empty;

        public string SalesOrderNo = string.Empty;
        public string InvoiceNo = string.Empty;
        public string CustomerPONo = string.Empty;
        public string CustomerNo = string.Empty;

        public string ShipToName = string.Empty;
        public string ShipToStore = string.Empty;
        public string ShipToAddress1 = string.Empty;
        public string ShipToAddress2 = string.Empty;
        public string ShipToAddress3 = string.Empty;
        public string ShipToCity = string.Empty;
        public string ShipToState = string.Empty;
        public string ShipToZip = string.Empty;
        public string ShipToCountry = string.Empty;

        public string InvoiceMerchandise = string.Empty;
        public string InvoiceTotal = string.Empty;
        public string InvoiceDate = string.Empty;
        public string OrderDate = string.Empty;
        public string FirstTracking = string.Empty;

        public string TodayDate = DateTime.Today.ToString("yyyyMMdd");
        public string TimeNow = DateTime.Now.ToString("HHmm");

        // ReadValue is a wrapper to avoid using System.Reflection
        // when reading edi maps.

        // If the value is enclosed in square brackets, try to read
        // from the invoice, if not, read the static map value
        // This lets me just input vendor numbers into the mapping
        // instead of having to parse them from somewhere

        public Invoice(string InvoiceNo)
        {
            this.InvoiceNo = InvoiceNo;
        }

        public Invoice()
        {

        }
        public string ReadValue(string Element)
        {
            bool isAttribute = Element.StartsWith("[") && Element.EndsWith("]");

            if (isAttribute == true)
            {
                string Attribute = Element.Replace("[", string.Empty).Replace("]", string.Empty);
                switch (Attribute)
                {
                    case "InvoiceNo":
                        Attribute = this.InvoiceNo;
                        break;
                    case "CustomerPONo":
                        Attribute = this.CustomerPONo;
                        break;
                    case "CustomerNo":
                        Attribute = this.CustomerNo;
                        break;
                    case "ShipToStore":
                        Attribute = this.ShipToStore;
                        break;
                    case "ShipToName":
                        Attribute = this.ShipToName;
                        break;
                    case "ShipToAddress1":
                        Attribute = this.ShipToAddress1;
                        break;
                    case "ShipToAddress2":
                        Attribute = this.ShipToAddress2;
                        break;
                    case "ShipToAddress3":
                        Attribute = this.ShipToAddress3;
                        break;
                    case "ShipToCity":
                        Attribute = this.ShipToCity;
                        break;
                    case "ShipToState":
                        Attribute = this.ShipToState;
                        break;
                    case "ShipToZip":
                        Attribute = this.ShipToZip;
                        break;
                    case "ShipToCountry":
                        Attribute = this.ShipToCountry;
                        break;
                    case "InvoiceMerchandise":
                        Attribute = this.LineItems.Sum(it => it.UnitPrice).ToString();
                        break;
                    case "InvoiceLines":
                        Attribute = this.LineItems.Count.ToString();
                        break;
                    case "InvoiceTotal":
                        Attribute = this.InvoiceTotal;
                        break;
                    case "InvoiceDate":
                        Attribute = this.InvoiceDate;
                        break;
                    case "OrderDate":
                        Attribute = this.OrderDate;
                        break;
                    case "TodayDate":
                        Attribute = this.TodayDate;
                        break;
                    case "TimeNow":
                        Attribute = this.TimeNow;
                        break;
                    case "FirstTracking":
                        Attribute = this.FirstTracking;
                        break;
                    case "750DSFee":
                        Attribute = (Convert.ToInt32(this.InvoiceTotal) + 750).ToString();
                        break;
                    case "200DSFee":
                        Attribute = (Convert.ToInt32(this.InvoiceTotal) + 200).ToString();
                        break;
                    case "HL_Package":
                        break;
                    case "HL_No":
                        break;

                }
                return Attribute;
            }
            else
            {
                return Element;
            }
        }

        

    }
}
