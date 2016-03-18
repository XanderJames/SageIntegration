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
        public string Store = string.Empty;
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

        static Address LookupStore(string StoreNo, string CustomerNo)
        {
            Address address = new Address();

            using (OleDbConnection ediCn = new OleDbConnection(@"provider=Microsoft.ACE.OLEDB.12.0;data source=" + ProductionDatabase))
            using (OleDbCommand storeCmd = new OleDbCommand())
            {
                storeCmd.Connection = ediCn;
                storeCmd.CommandType = CommandType.Text;
                storeCmd.CommandText = @"SELECT Stores.CustomerNo, Stores.Number, Stores.Name, Stores.Address, Stores.Address2, Stores.City, Stores.State, Stores.Zip
                                        FROM Stores
                                        WHERE (Stores.CustomerNo=?) AND (Stores.Number=?)";

                storeCmd.Parameters.Add("@CustomerNo", OleDbType.VarChar).Value = CustomerNo;
                storeCmd.Parameters.Add("@Number", OleDbType.VarChar).Value = StoreNo;


                using (OleDbDataAdapter storeAdp = new OleDbDataAdapter(storeCmd))
                using (DataTable storeDt = new DataTable())
                {
                    storeAdp.Fill(storeDt);
                    if (storeDt.Rows.Count > 0)
                    {
                        DataRow storerow = storeDt.Rows[0];
                        address.Store = storerow["Number"].ToString();
                        address.Name = storerow["Name"].ToString();
                        address.Address1 = storerow["Address"].ToString();
                        address.Address2 = storerow["Address2"].ToString();
                        address.City = storerow["City"].ToString();
                        address.State = storerow["State"].ToString();
                        address.Zipcode = storerow["Zip"].ToString();
                    }
                }
            }

            return address;

        }


        public SalesOrder CreateStoreOrder(string Key, string CustomerNo)
        {
            SalesOrder order = new SalesOrder();
            Address address = new Address();

            address = LookupStore(Key, CustomerNo);
            order.Addresses.Add("92", address);

            order.Store = Key;
            order.ShipToName = address.Name;
            order.ShipToAddress1 = address.Address1;
            order.ShipToAddress2 = address.Address2;
            order.ShipToAddress3 = address.Address3;
            order.ShipToCity = address.City;
            order.ShipToState = address.State;
            order.ShipToZipcode = address.Zipcode;
            order.ShipToCountry = address.Country;

            this.StoreOrders.Add(Key, order);

            return order;
        }

        public virtual void Export(int TransactionID)
        {
            string database = string.Empty;
            if (System.Diagnostics.Debugger.IsAttached)
            {
                database = ProductionDatabase;
            }
            else
            {
                database = ProductionDatabase;
            }

            using (OleDbConnection connection = new OleDbConnection(@"provider=Microsoft.ACE.OLEDB.12.0;data source=" + database))
            {
                int salesorderid = 0;
                using (OleDbCommand command = new OleDbCommand())
                {
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = @"INSERT INTO [SalesOrder] ([Transaction_ID], [OrderDate], [RequiredDate], [CustomerOrderNo], [CustomerPONo], [CustomerShipvia],
                                        [Shipvia], [GiftHeader], [GiftMessage], [SellingChannel], [PaymentTender], [StoreNo], [ReservationNo], [ShippingCharge],
                                        [SalesTax], [OrderCredit], [OrderTotal], [ShipToCompany], [ShipToName], [ShipToAddress1], [ShipToAddress2], [ShipToCity],
                                        [ShipToState], [ShipToZipcode], [ShipToCountry],  [ShipToPhone], [ShipToEmail], [BillToCompany], [BillToName], [BillToAddress1],
                                        [BillToAddress2], [BillToCity], [BillToState], [BillToZipcode], [BillToCountry]) VALUES 
                                        (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

                    command.Parameters.Add("@Transaction_ID", OleDbType.Integer).Value = TransactionID;

                    command.Parameters.Add("@OrderDate", OleDbType.DBDate).Value = OrderDate;
                    command.Parameters.Add("@RequiredDate", OleDbType.DBDate).Value = RequiredDate;

                    command.Parameters.Add("@CustomerOrderNo", OleDbType.VarChar).Value = CustomerOrderNo;
                    command.Parameters.Add("@CustomerPONo", OleDbType.VarChar).Value = CustomerPONo;
                    command.Parameters.Add("@CustomerShipvia", OleDbType.VarChar).Value = CustomerShipvia;
                    command.Parameters.Add("@Shipvia", OleDbType.VarChar).Value = Shipvia;
                    command.Parameters.Add("@GiftHeader", OleDbType.VarChar).Value = GiftHeader;
                    command.Parameters.Add("@GiftMessage", OleDbType.VarChar).Value = GiftMessage;
                    command.Parameters.Add("@SellingChannel", OleDbType.VarChar).Value = SellingChannel;
                    command.Parameters.Add("@PaymentTender", OleDbType.VarChar).Value = PaymentTender;
                    command.Parameters.Add("@Store", OleDbType.VarChar).Value = Store;
                    command.Parameters.Add("@ReservationNo", OleDbType.VarChar).Value = ReservationNo;

                    command.Parameters.Add("@ShippingCharge", OleDbType.Integer).Value = ShippingCost;
                    command.Parameters.Add("@SalesTax", OleDbType.Integer).Value = SalesTax;
                    command.Parameters.Add("@OrderCredit", OleDbType.Integer).Value = Credit;
                    command.Parameters.Add("@OrderTotal", OleDbType.Integer).Value = OrderTotal;

                    command.Parameters.Add("@ShipToCompany", OleDbType.VarChar).Value = ShipToCompany;
                    command.Parameters.Add("@ShipToName", OleDbType.VarChar).Value = ShipToName;
                    command.Parameters.Add("@ShipToAddress1", OleDbType.VarChar).Value = ShipToAddress1;
                    command.Parameters.Add("@ShipToAddress2", OleDbType.VarChar).Value = ShipToAddress2;
                    command.Parameters.Add("@ShipToCity", OleDbType.VarChar).Value = ShipToCity;
                    command.Parameters.Add("@ShipToState", OleDbType.VarChar).Value = ShipToState;
                    command.Parameters.Add("@ShipToZipcode", OleDbType.VarChar).Value = ShipToZipcode;
                    command.Parameters.Add("@ShipToCountry", OleDbType.VarChar).Value = ShipToCountry;
                    command.Parameters.Add("@ShipToPhone", OleDbType.VarChar).Value = ShipToPhone;
                    command.Parameters.Add("@ShipToEmail", OleDbType.VarChar).Value = ShipToEmail;

                    command.Parameters.Add("@BillToCompany", OleDbType.VarChar).Value = BillToCompany;
                    command.Parameters.Add("@BillToName", OleDbType.VarChar).Value = BillToName;
                    command.Parameters.Add("@BillToAddress1", OleDbType.VarChar).Value = BillToAddress1;
                    command.Parameters.Add("@BillToAddress2", OleDbType.VarChar).Value = BillToAddress2;
                    command.Parameters.Add("@BillToCity", OleDbType.VarChar).Value = BillToCity;
                    command.Parameters.Add("@BillToState", OleDbType.VarChar).Value = BillToState;
                    command.Parameters.Add("@BillToZipcode", OleDbType.VarChar).Value = BillToZipcode;
                    command.Parameters.Add("@BillToCountry", OleDbType.VarChar).Value = BillToCountry;

                    connection.Open();
                    command.ExecuteNonQuery();

                    command.Parameters.Clear();
                    command.CommandText = "Select @@Identity";
                    salesorderid = Convert.ToInt32(command.ExecuteScalar());
                    connection.Close();

                }

                foreach (LineItem lineitem in this.LineItems())
                {
                    using (OleDbCommand command = new OleDbCommand())
                    {
                        command.Connection = connection;
                        command.CommandType = CommandType.Text;
                        command.CommandText = @"INSERT INTO [LineItem] ([SalesOrder_ID], [ItemCode], [Quantity], [LineNo], [UPC], [CustomerItemCode],
                                        [SGCItemCode], [CustomerItemNo], [MFGCode], [UnitCost], [RetailPrice], [PS_Description1], [PS_Description2],
                                        [PS_Description3], [Description], [Monogram], [ImageURL]) VALUES 
                                        (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

                        command.Parameters.Add("@SalesOrder_ID", OleDbType.Integer).Value = salesorderid;
                        command.Parameters.Add("@ItemCode", OleDbType.VarChar).Value = lineitem.ItemCode;
                        command.Parameters.Add("@Quantity", OleDbType.Integer).Value = lineitem.Quantity;
                        command.Parameters.Add("@LineNo", OleDbType.Integer).Value = lineitem.LineNo;
                        command.Parameters.Add("@UPC", OleDbType.VarChar).Value = lineitem.UPC;
                        command.Parameters.Add("@CustomerItemCode", OleDbType.VarChar).Value = lineitem.CustomerItemCode;
                        command.Parameters.Add("@SGCItemCode", OleDbType.VarChar).Value = lineitem.SGCItemCode;
                        command.Parameters.Add("@CustomerItemNo", OleDbType.VarChar).Value = lineitem.CustItemNo;
                        command.Parameters.Add("@MFGCode", OleDbType.VarChar).Value = lineitem.MFGCode;

                        command.Parameters.Add("@UnitPrice", OleDbType.Decimal).Value = lineitem.UnitPrice;
                        command.Parameters.Add("@RetailPrice", OleDbType.Decimal).Value = lineitem.RetailPrice;

                        command.Parameters.Add("@PS_Description1", OleDbType.VarChar).Value = lineitem.PSDescription1;
                        command.Parameters.Add("@PS_Description2", OleDbType.VarChar).Value = lineitem.PSDescription2;
                        command.Parameters.Add("@PS_Description3", OleDbType.VarChar).Value = lineitem.PSDescription3;
                        command.Parameters.Add("@Description", OleDbType.VarChar).Value = lineitem.Description;
                        command.Parameters.Add("@Monogram", OleDbType.VarChar).Value = lineitem.Monogram;
                        command.Parameters.Add("@ImageURL", OleDbType.VarChar).Value = lineitem.ImageURL;

                        connection.Open();
                        command.ExecuteNonQuery();
                        connection.Close();

                    }


                }

            }
        }

    }

}
