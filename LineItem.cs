using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SageIntegration
{
    public class LineItem
    {
        public int Quantity = 0;
        public int QuantityShipped = 0;
        public int LineNo = 0;
        public string LineKey;
        public string _Monogram = string.Empty;
        public string ItemCode = string.Empty;
        public string SGCItemCode = string.Empty;
        public string CustItemNo = string.Empty;
        public string MFGCode = string.Empty;
        public string CustomerItemCode = string.Empty;
        public string UPC = string.Empty;
        public string Description = string.Empty;
        public string PSDescription1 = string.Empty;
        public string PSDescription2 = string.Empty;
        public string PSDescription3 = string.Empty;
        public string ImageURL = string.Empty;
        public decimal UnitPrice = 0;
        public decimal RetailPrice = 0;

        public string Monogram
        {
            get { return _Monogram; }
            set { _Monogram = LimitString(value, 200); }
        }

        public string LimitString(string Input, int Length)
        {
            Input = Input.Trim();
            if (Input.Length < Length)
            {
                return Input;
            }
            else
            {
                return Input.Substring(0, Length);
            }

        }

        // Plans to move this to a .dll and use it everywhere

        // ReadValue is a wrapper to avoid using System.Reflection
        // when reading edi maps.

        // If the value is enclosed in square brackets, try to read
        // from the lineitem, if not, read the static map value
        // This lets me just input vendor numbers into the mapping
        // instead of having to parse them
        public string ReadValue(string Element)
        {
            bool isAttribute = Element.StartsWith("[") && Element.EndsWith("]");

            if (isAttribute == true)
            {
                string Attribute = Element.Replace("[", string.Empty).Replace("]", string.Empty);
                switch (Attribute)
                {
                    case "Quantity":
                        Attribute = this.Quantity.ToString();
                        break;
                    case "LineNo":
                        Attribute = this.LineNo.ToString();
                        break;
                    case "VendorSKU":
                        Attribute = this.SGCItemCode;
                        break;
                    case "CustomerSKU":
                        Attribute = this.CustomerItemCode;
                        break;
                    case "Description":
                        Attribute = this.Description;
                        break;
                    case "UnitPrice":
                        Attribute = this.UnitPrice.ToString("0.##");
                        break;
                    case "UPC":
                        Attribute = this.UPC;
                        break;
                }
                return Attribute;
            }
            else
            {
                return Element;
            }
        }

        public LineItem()
        {

        }

        public LineItem(string ItemCode, int Quantity, string Monogram = "", string Description = "")
        {
            this.ItemCode = ItemCode;
            this.Quantity = Quantity;
            this.Monogram = Monogram;
            this.Description = Description;
        }

        public LineItem(string ItemCode, int Quantity, int LineNo, string Monogram = "", string Description = "")
        {
            this.ItemCode = ItemCode;
            this.Quantity = Quantity;
            this.LineNo = LineNo;
            this.Monogram = Monogram;
            this.Description = Description;
        }

    }
}
