using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Data;

namespace SageIntegration
{
    public class SalesOrder
    {
        public static string ProductionDatabase = @"\\SGC-APPLICATION\EDIService\bin\EDI.accdb";
        public static string TestDatabase = "";
        // This object is just built to be something I can reuse when
        // creating order import processes.

        private List<LineItem> _lineitems = new List<LineItem>();
        private List<LineItem> _errorlineitems = new List<LineItem>();
        private string _SalesOrderNo = string.Empty;
        private string _ShipToCompany = string.Empty;
        private string _ShipToName = string.Empty;
        private string _ShipToAddress1 = string.Empty;
        private string _ShipToAddress2 = string.Empty;
        private string _ShipToAddress3 = string.Empty;
        private string _ShipToCity = string.Empty;
        private string _ShipToState = string.Empty;
        private string _ShipToZipcode = string.Empty;
        private string _ShipToCountry = string.Empty;
        private string _CustomerNo = string.Empty;

        private string _CustomerPONo;

        public int TransactionSet = 0;
        public int ID = 0;

        public int SalesTax = 0;
        public int ShippingCost = 0;
        public int OrderTotal = 0;
        public int Credit = 0;
        public string SellingChannel = string.Empty;
        public string PaymentTender = string.Empty;
        public string ReservationNo = string.Empty;

        public string CustomerOrderNo = string.Empty;
        public DateTime ProcessDate = DateTime.Now;
        public DateTime RequiredDate = DateTime.Now;
        public DateTime OrderDate = DateTime.Now;
        public string CustomerShipvia = string.Empty;
        public string Shipvia = string.Empty;
        public string GiftMessage = string.Empty;
        public string GiftHeader = string.Empty;
        public string StoreNo = string.Empty;
        public double Freight = 0;

        public List<string> TrackingNumbers = new List<string>();

        public Dictionary<string, Address> Addresses = new Dictionary<string, Address>();
        public Dictionary<string, SalesOrder> StoreOrders = new Dictionary<string, SalesOrder>();

        public string ShipToEmail = string.Empty;
        public string ShipToPhone = string.Empty;

        public string BillToCompany = string.Empty;
        public string BillToName = string.Empty;
        public string BillToAddress1 = string.Empty;
        public string BillToAddress2 = string.Empty;
        public string BillToAddress3 = string.Empty;
        public string BillToCity = string.Empty;
        public string BillToState = string.Empty;
        public string BillToZipcode = string.Empty;
        public string BillToCountry = string.Empty;
        public string BillToEmail = string.Empty;
        public string BillToPhone = string.Empty;

        public string SalesOrderNo
        {
            get { return _SalesOrderNo; }
            set { _SalesOrderNo = TextUtilities.LimitString(value, 7); }
        }

        public SalesOrder(string CustomerNo, string CustomerPONo)
        {
            this.CustomerNo = CustomerNo;
        }

        public string CustomerPONo
        {
            get { return _CustomerPONo; }
            set { _CustomerPONo = TextUtilities.LimitString(value, 15); }
        }

        public SalesOrder() { }

        public string CustomerNo
        {
            get { return _CustomerNo; }
            set { _CustomerNo = TextUtilities.LimitString(value, 20); }
        }

        public string ShipToName
        {
            get { return _ShipToName; }
            set { _ShipToName = TextUtilities.LimitString(value, 30); }
        }

        public string ShipToCompany
        {
            get { return _ShipToCompany; }
            set { _ShipToCompany = value; }
        }

        public string ShipToAddress1
        {
            get { return _ShipToAddress1; }
            set { _ShipToAddress1 = TextUtilities.LimitString(value, 30); }
        }

        public string ShipToAddress2
        {
            get { return _ShipToAddress2; }
            set { _ShipToAddress2 = TextUtilities.LimitString(value, 30); }

        }

        public string ShipToAddress3
        {
            get { return _ShipToAddress3; }
            set { _ShipToAddress3 = TextUtilities.LimitString(value, 30); }
        }

        public string ShipToCity
        {
            get { return _ShipToCity; }
            set { _ShipToCity = TextUtilities.LimitString(value, 20); }
        }

        public string ShipToState
        {
            get { return _ShipToState; }
            set { _ShipToState = TextUtilities.LimitString(value, 2); }
        }

        public string ShipToZipcode
        {
            get { return _ShipToZipcode; }
            set
            {
                string Tempzip = value;
                if (Tempzip != string.Empty)
                {
                    Tempzip = Tempzip.Replace("-", "");
                    Tempzip = Tempzip.Replace(" ", "");
                    if (Regex.IsMatch(Tempzip, @"^\d+$"))
                    {
                        int ZipcodeInt;
                        int.TryParse(Tempzip, out ZipcodeInt);
                        if (Tempzip.Length <= 5)
                        {
                            Tempzip = ZipcodeInt.ToString("D5");
                        }

                        _ShipToZipcode = Tempzip;
                    }
                    else
                    {
                        _ShipToZipcode = Tempzip[0].ToString() + Tempzip[1].ToString() + Tempzip[2].ToString() + " " + Tempzip[3].ToString() + Tempzip[4].ToString() + Tempzip[5].ToString();
                    }
                }

            }
        }

        public string ShipToCountry
        {
            get { return _ShipToCountry; }
            set { _ShipToCountry = TextUtilities.LimitString(value, 3); }
        }

        public List<LineItem> LineItems()
        {
            return _lineitems;
        }
        public LineItem AddLine(LineItem LineItem = null)
        {
            if (LineItem == null)
            {
                LineItem = new LineItem();
            }
            _lineitems.Add(LineItem);
            return LineItem;
        }

        public void ErrorLine(LineItem lineitem)
        {
            _lineitems.Remove(lineitem);
            _errorlineitems.Add(lineitem);
        }

        public Address CreateAddress(string key)
        {
            Address address = new Address();
            this.Addresses.Add(key, address);

            return address;
        }

        public Address AddAddress(string key, Address address)
        {
            this.Addresses.Add(key, address);

            return address;
        }

        public SalesOrder CreateStoreOrder(string Key, string CustomerNo)
        {
            SalesOrder order = new SalesOrder();
            Address address = new Address();

            order.Addresses.Add("92", address);

            order.StoreNo = Key;

            this.StoreOrders.Add(Key, order);

            return order;
        }
        
    }

}
