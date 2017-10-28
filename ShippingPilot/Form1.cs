using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShippingPilot
{
    public partial class Form1 : Form
    {
        string fileName = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        // This is for Submit Button Click
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable ExcelData = ReadExcel();
                bool _IsFirstRow = true;
                foreach (DataRow dr in ExcelData.Rows)
                {
                    if (_IsFirstRow)
                    {
                        _IsFirstRow = false;
                        continue;
                    }

                    if (dr.ItemArray[22].ToString().Trim() == "Pilot Freight Basic Delivery")
                    {
                        PilotOperations(dr);
                    }
                    else
                    {
                    }
                }
            }
            catch (IOException)
            {
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        // This is for Browse Button Click
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                int size = -1;
                openFileDialog1.Filter = "allfiles|*.xlsx";
                DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
                if (result == DialogResult.OK) // Test result.
                {
                    fileName = openFileDialog1.FileName;
                    textBox1.Text = openFileDialog1.FileName;
                }
            }
            catch (Exception Ex)
            {

            }
        }

        public DataTable ReadExcel()
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();

            if (Path.GetExtension(fileName).CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [PO Details$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch (Exception ext)
                {

                }
            }
            return dtexcel;
        }

        public void PilotOperations(DataRow dr)
        {
            try
            {
                TestPilotServiceref.dsShipment ds;
                TestPilotServiceref.ShipmentService ws = new TestPilotServiceref.ShipmentService();
                //returns dsShipment with default values
                ds = ws.GetNewShipment();

                /* Test Data for mail trigger
                 ds.Shipment[0].LocationID = 11223344;
                 ds.Shipment[0].ControlStation = "LAX";
                 ds.Shipment[0].AddressID = 12345;
                 ds.Shipment[0].TariffHeaderID = 98765;
                 * Test Data End  */
                #region - Shipment Constant Details - Don't Modify - Start
                ds.Shipment[0].LocationID = 12884574;
                ds.Shipment[0].ControlStation = "GOP";
                ds.Shipment[0].AddressID = 80149;
                ds.Shipment[0].TariffHeaderID = 21942;
                #endregion - Shipment Constant Details - Don't Modify - End 

                ds.Shipment[0].IsScreeningConsent = "Y";
                ds.Shipment[0].PayType = "T";
                ds.Shipment[0].Service = "EC";
                ds.Shipment[0].ProductName = "B";
                ds.Shipment[0].ProductDescription = "BASIC";
                ds.Shipment[0].SpecialInstructions = "Do Not Drop";
                ds.Shipment[0].ShipperRef = "shipref123";
                ds.Shipment[0].ConsigneeRef = "conref987";
                ds.Shipment[0].ConsigneeAttn = "";
                ds.Shipment[0].ThirdPartyAuth = "";
                ds.Shipment[0].ShipDate = System.DateTime.Now;
                ds.Shipment[0].ReadyTime = System.DateTime.Now;
                ds.Shipment[0].CloseTime = System.DateTime.Now;
                ds.Shipment[0].EmailBOL = "rasool@ufe2.com";
                ds.Shipment[0].COD = 0;
                ds.Shipment[0].DeclaredValue = 0;
                ds.Shipment[0].DebrisRemoval = false; // Set to false if not exists 
                ds.Shipment[0].Hazmat = false; // Set to false if not exists 
                ds.Shipment[0].HazmatPhone = string.Empty;// Set to Empty String if not exists
                ds.Shipment[0].AirtrakQuoteNo = 0; //should be defaulted to the integer value 0


                //populate shipper information;
                ds.Shipper[0].Name = "US FURNISHINGS EXPRESS INC";
                ds.Shipper[0].Address1 = "5125 SCHAEFER AVE";
                ds.Shipper[0].Address2 = "STE 104";
                ds.Shipper[0].City = "CHINO";
                ds.Shipper[0].State = "CALIFORNIA";
                ds.Shipper[0].Zipcode = "91710";
                ds.Shipper[0].Country = "UNITED STATES OF AMERICA";
                ds.Shipper[0].Contact = "JIA GUO";
                ds.Shipper[0].Phone = "9498781998";
                ds.Shipper[0].Email = "JGUO@UFE2.COM";
                ds.Shipper[0].PrivateRes = false;
                ds.Shipper[0].Hotel = false;
                ds.Shipper[0].Inside = false;
                ds.Shipper[0].Liftgate = false;
                ds.Shipper[0].TwoManHours = 0;
                ds.Shipper[0].WaitTimeHours = 0;

                ds.Consignee[0].Name = "PILOT AIR FREIGHT (ORD)";
                ds.Consignee[0].Address1 = "695 TOUHY";
                ds.Consignee[0].City = "ELK GROVE VILLAGE";
                ds.Consignee[0].State = "IL";
                ds.Consignee[0].Zipcode = "60007";
                ds.Consignee[0].Country = "US";
                ds.Consignee[0].Contact = "";
                ds.Consignee[0].Phone = "630-238-3245";
                ds.Consignee[0].Email = "rasool@ufe2.com";
                ds.Consignee[0].PrivateRes = false;
                ds.Consignee[0].Hotel = false;
                ds.Consignee[0].Inside = false;
                ds.Consignee[0].Liftgate = false;
                ds.Consignee[0].TwoManHours = 0;
                ds.Consignee[0].WaitTimeHours = 0;

                ds.ThirdParty[0].Name = "BILL PAYER INC";
                ds.ThirdParty[0].Address1 = "123 MONEYBAGS LANE";
                ds.ThirdParty[0].Address2 = "SUITE 100";
                ds.ThirdParty[0].City = "BEVERLY HILLS";
                ds.ThirdParty[0].State = "CA";
                ds.ThirdParty[0].Zipcode = "90210";
                ds.ThirdParty[0].Country = "US";
                ds.ThirdParty[0].Contact = "MR. MONOPOLY";
                ds.ThirdParty[0].Phone = "123-456-7890";
                ds.ThirdParty[0].Email = "rasool@ufe2.com";


                //creating lineitems
                TestPilotServiceref.dsShipment.LineItemsRow drline = ds.LineItems.NewLineItemsRow();
                drline.Pieces = 1;
                drline.Weight = 600;
                drline.Description = "Pallet";
                drline.Length = 48;
                drline.Width = 48;
                drline.Height = 36;
                ds.LineItems.Rows.Add(drline);

                ///save the shipment
                TestPilotServiceref.PilotShipmentResult res = ws.Save(ds);

                if (res.IsError == false)
                {
                    Console.WriteLine("Shipment has been saved.");
                    res.dsResult.WriteXml(@"G:\Mallesh\Coding\result.xml");
                    //shou
                    //return true;
                }
                else
                {
                    Console.WriteLine("Shipment has been not saved.");
                    //return false;
                }

                //Add wsShipment as Web Reference pointing to service address 

                TestPilotServiceref.dsVoid ds2;
                //returns dsVoid with default values
                ds2 = ws.GetNewVoid();
                ds2.Void[0].LocationID = 11223344;
                ds2.Void[0].ControlStation = "LAX";
                ds2.Void[0].AddressID = 12345;
                ds2.Void[0].TariffHeaderID = 98765;
                ds2.Void[0].ProNumber = res.dsResult.Shipment[0].ProNumber.ToString(); //"062002488";
                                                                                       //void the shipment 
                TestPilotServiceref.PilotShipmentResult res2 = ws.Void(ds2);
                if (!res2.IsError && res2.Message == "Shipment Void Success")
                {
                    Console.WriteLine("Shipment Void Success");
                }
                else
                {
                    Console.WriteLine("Shipment void Failed !.");
                }
            }
            catch (Exception ex)
            {


            }
        }
    }
}
