using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShippingPilot
{
    public partial class Form1 : Form
    {
        string fileName = "";
        bool _IsSaved = false, _IsVoid = false;
        string Remarks = "", ProNumber = "", PrintPath = "";
        DataTable dtLineItemData = null;

        DataTable table;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            lblInfo.Text = "";
            try
            {
                DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
                if (result == DialogResult.OK) // Test result.
                {
                    fileName = openFileDialog1.FileName;
                    txtFilePath.Text = openFileDialog1.FileName;
                }
            }
            catch (Exception Ex)
            {
                lblInfo.Text = "Exception while Browsing file";
                Remarks += $"BtnBrowse - {Ex}";
            }
        }

        private void btnResponse_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                txtResponsePath.Text = openFileDialog1.FileName;
                btnSubmit.Enabled = true;
                lblInfo.Text = "Click on Submit";
            }
            var FileExtensin = Path.GetExtension(openFileDialog1.FileName);
            if (FileExtensin == "xlsx" || FileExtensin == ".xls")
            {
                txtResponsePath.Text = "";
                MessageBox.Show(@"Please Select Response File. Ex:: C:\PilotResponse\abc.xlsx");
            }
            if (!File.Exists(openFileDialog1.FileName) && txtResponsePath.Text != "")
            {
                txtResponsePath.Text = "";
                MessageBox.Show("File Not Exists in given location");
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                table = new DataTable();
                if (txtFilePath.Text.Trim() == "")
                {
                    MessageBox.Show("OOOPS!!! First Select Pilot File to Submit !");
                    return;
                }

                if (txtLineItem.Text.Trim() == "")
                {
                    MessageBox.Show("OOOPS!!! First Select Line Item file Path !");
                    return;
                }

                if (txtFilePath.Text.Trim() == "")
                {
                    MessageBox.Show("OOOPS!!! First Select Response file Path !");
                    return;
                }

                DialogResult res = MessageBox.Show("Are you Sure you want to Submit ?", "Ready To Submit!", MessageBoxButtons.YesNo);
                if (res.Equals(DialogResult.Yes))
                {
                    btnSubmit.Enabled = false;
                    DataTable dtPilotExcelData = ReadPilotExcel();
                    dtLineItemData = new DataTable();
                    dtLineItemData = ReadLineItemExcel();
                    if (dtPilotExcelData == null || dtLineItemData == null)
                        return;

                    lblInfo.Text = "Please be patient while processing your request !!";
                    bool _IsFirstRow = true;
                    foreach (DataRow dr in dtPilotExcelData.Rows)
                    {
                        if (_IsFirstRow)
                        {
                            _IsFirstRow = false;
                            continue;
                        }

                        if (dr.ItemArray[22].ToString().Trim() == string.Empty)
                            continue;

                        if (dr.ItemArray[22].ToString().Trim() == "Pilot Freight Basic Delivery")
                        {
                            PilotOperations(dr);
                        }
                        else
                        {
                            _IsSaved = false;
                            _IsVoid = false;
                            Remarks = "Invalid Carrier";
                            TableOperations(dr.ItemArray[0].ToString(), dr.ItemArray[1].ToString(), dr.ItemArray[25].ToString());
                        }
                    }
                    lblInfo.Text = "Request Submitted Succesfully.";
                }
            }
            catch (Exception ex)
            {
                lblInfo.Text = "Exception while submitting!";
                Remarks += $"BtnSubmit - {ex}";
            }
            finally
            {
                ExportToExcel();
                txtFilePath.Text = "";
                //txtResponsePath.Text = "";
                btnBrowse.Enabled = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            lblUserName.Text = "Welcome Jia Guo";
            lblDateTime.Text = DateTime.Now.DayOfWeek.ToString() + " , " + DateTime.Now.ToShortDateString();
            lblInfo.Text = "Browse File to Submit & Select Response File Path";
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.csv";
        }

        public DataTable ReadPilotExcel()
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();

            //if (fileextensin == "xlsx" || fileextensin == ".xls" || fileextensin == "xlsm")
            //conn = @"provider=microsoft.jet.oledb.4.0;data source=" + fileName + ";extended properties='excel 8.0;hrd=yes;imex=1';"; //for below excel 2007  
            //else
            conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [PO Details$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch (Exception ex)
                {
                    lblInfo.Text = "*** Sheet name Should be PO Details";
                    Remarks += $"ReadPilotExcel - {ex}";
                }
            }
            return dtexcel;
        }

        public DataTable ReadLineItemExcel()
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();

            //if (fileextensin == "xlsx" || fileextensin == ".xls" || fileextensin == "xlsm")
            //conn = @"provider=microsoft.jet.oledb.4.0;data source=" + fileName + ";extended properties='excel 8.0;hrd=yes;imex=1';"; //for below excel 2007  
            //else
            conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtLineItem.Text.Trim() + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Item Report$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch (Exception ex)
                {
                    lblInfo.Text = "***Sheet name Should be Item Report";
                    Remarks += $"ReadLineItemExcel - {ex}";
                }
            }
            return dtexcel;
        }

        private void btnLineItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult LineItemResult = openFileDialog1.ShowDialog(); // Show the dialog.
                if (LineItemResult == DialogResult.OK) // Test result.
                {
                    fileName = openFileDialog1.FileName;
                    txtLineItem.Text = openFileDialog1.FileName;
                }
            }
            catch (Exception Ex)
            {
                lblInfo.Text = "Exception while Browsing Line Item";
                Remarks += $"BtnBrowse - {Ex}";
            }
        }

        public void PilotOperations(DataRow dr)
        {
            try
            {
                //TestPilotServiceref  - For Testing
                //CoPilotProd       - For Production
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

                ds.Shipment[0].IsScreeningConsent = "Yes"; // Doubt
                ds.Shipment[0].PayType = "THIRD PARTY";   // Doubt
                ds.Shipment[0].Service = "BA"; // Doubt
                ds.Shipment[0].ProductName = dr.ItemArray[16].ToString();
                ds.Shipment[0].ProductDescription = dr.ItemArray[18].ToString(); // Item Description
                ds.Shipment[0].SpecialInstructions = ""; // Doubt
                ds.Shipment[0].ShipperRef = dr.ItemArray[0].ToString();
                ds.Shipment[0].ConsigneeRef = dr.ItemArray[0].ToString();
                ds.Shipment[0].ConsigneeAttn = dr.ItemArray[4].ToString(); // Customer Phone Number
                ds.Shipment[0].ThirdPartyAuth = "";
                ds.Shipment[0].ShipDate = System.DateTime.Now;  // Ship By
                ds.Shipment[0].ReadyTime = System.DateTime.Now; // Doubt  15:00
                ds.Shipment[0].CloseTime = System.DateTime.Now; // Doubt  17:00
                ds.Shipment[0].EmailBOL = "bols@ufe2.com"; // Doubt
                ds.Shipment[0].COD = 0;
                ds.Shipment[0].DeclaredValue = 0;
                ds.Shipment[0].DebrisRemoval = false; // Set to false if not exists 
                ds.Shipment[0].Hazmat = false; // Set to false if not exists 
                ds.Shipment[0].HazmatPhone = string.Empty;// Set to Empty String if not exists
                ds.Shipment[0].AirtrakQuoteNo = 0; //should be defaulted to the integer value 0

                #region Shipper Address will remain same for all the Orders - Start
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
                #endregion Shipper Address will remain same for all the Orders - End

                ds.Consignee[0].Name = dr.ItemArray[4].ToString(); // Customer Name
                ds.Consignee[0].Address1 = dr.ItemArray[8].ToString() + " , " + dr.ItemArray[9].ToString();  // Customer Shipping Address
                ds.Consignee[0].City = dr.ItemArray[10].ToString();     // City
                ds.Consignee[0].State = dr.ItemArray[11].ToString();    // State
                ds.Consignee[0].Zipcode = dr.ItemArray[12].ToString();  // Zip
                ds.Consignee[0].Country = "UNITED STATES OF AMERICA"; // Doubt
                ds.Consignee[0].Contact = dr.ItemArray[4].ToString();  // Doubt ??
                ds.Consignee[0].Phone = dr.ItemArray[6].ToString(); // Customer Phone Number
                ds.Consignee[0].Email = "";
                ds.Consignee[0].PrivateRes = false;
                ds.Consignee[0].Hotel = false;
                ds.Consignee[0].Inside = false;
                ds.Consignee[0].Liftgate = false;
                ds.Consignee[0].TwoManHours = 0;
                ds.Consignee[0].WaitTimeHours = 0;

                // Updated
                #region - Third Party Details Remain Same for all - Start 
                ds.ThirdParty[0].Name = "Walmart.com";
                ds.ThirdParty[0].Address1 = "850 Cherry Avenue";
                ds.ThirdParty[0].Address2 = "";
                ds.ThirdParty[0].City = "SAN BRUNO";
                ds.ThirdParty[0].State = "CALIFORNIA";
                ds.ThirdParty[0].Zipcode = "94066";
                ds.ThirdParty[0].Country = "UNITED STATES OF AMERICA";
                ds.ThirdParty[0].Contact = "Account # 2983341";
                ds.ThirdParty[0].Phone = "877-549-0409";
                ds.ThirdParty[0].Email = string.Empty;
                #endregion - Third Party Details Remain Same for all - End

                //creating lineitems
                TestPilotServiceref.dsShipment.LineItemsRow drline = ds.LineItems.NewLineItemsRow();

                foreach (DataRow Row in dtLineItemData.Rows)
                {
                    if (Row.ItemArray[0].ToString().Trim() == dr.ItemArray[16].ToString().Trim())
                    {
                        drline.Pieces = Convert.ToInt32(Row.ItemArray[1]);
                        drline.Weight = Convert.ToInt32(Row.ItemArray[5]);
                        drline.Description = Row.ItemArray[0].ToString();
                        drline.Length = Convert.ToInt32(Row.ItemArray[2]);
                        drline.Width = Convert.ToInt32(Row.ItemArray[3]);
                        drline.Height = Convert.ToInt32(Row.ItemArray[4]);
                        ds.LineItems.Rows.Add(drline);
                        /*
                        if (Row.ItemArray[6].ToString() != string.Empty)
                        {
                            TestPilotServiceref.dsShipment.LineItemsRow drline2 = ds.LineItems.NewLineItemsRow();
                            drline2.Pieces = Convert.ToInt32(Row.ItemArray[1]);
                            drline2.Weight = Convert.ToInt32(Row.ItemArray[9]);
                            drline2.Description = Row.ItemArray[0].ToString();
                            drline2.Length = Convert.ToInt32(Row.ItemArray[6]);
                            drline2.Width = Convert.ToInt32(Row.ItemArray[7]);
                            drline2.Height = Convert.ToInt32(Row.ItemArray[8]);
                            ds.LineItems.Rows.Add(drline2);
                        }
                        */
                        break;
                    }
                }



                ///save the shipment
                TestPilotServiceref.PilotShipmentResult SaveResp = ws.Save(ds);

                if (SaveResp.IsError == false)
                {
                    _IsSaved = true;
                    ProNumber = SaveResp.dsResult.Shipment[0].ProNumber.ToString();
                    Console.WriteLine("Shipment has been saved.");
                    Remarks += "Shipment has been saved";
                }
                else
                {
                    _IsSaved = false;
                    Console.WriteLine("Shipment has been not saved.");
                    Remarks += "Shipment Not saved";
                }

                //Add wsShipment as Web Reference pointing to service address 
                TestPilotServiceref.dsVoid ds2;
                //returns dsVoid with default values
                ds2 = ws.GetNewVoid();
                ds2.Void[0].LocationID = 12884574;
                ds2.Void[0].ControlStation = "GOP";
                ds2.Void[0].AddressID = 80149;
                ds2.Void[0].TariffHeaderID = 21942;
                ds2.Void[0].ProNumber = SaveResp.dsResult.Shipment[0].ProNumber.ToString();
                //void the shipment 
                TestPilotServiceref.PilotShipmentResult VoidResp = ws.Void(ds2);
                if (!VoidResp.IsError && VoidResp.Message == "Shipment Void Success")
                {
                    Console.WriteLine("Shipment Void Success");
                    _IsVoid = true;
                    Remarks += $" ; {VoidResp.Message}";
                }
                else
                {
                    _IsVoid = false;
                    Console.WriteLine("Shipment void Failed !.");
                    Remarks += $" ; {VoidResp.Message}";
                }
            }
            catch (Exception ex)
            {
                Remarks += $"PilotOperations - {ex}";
            }
            finally
            {
                TableOperations(dr.ItemArray[0].ToString(), dr.ItemArray[1].ToString(), dr.ItemArray[25].ToString());
            }
        }

        public void TableOperations(string PO, string Order, string Carrier)
        {

            if (table.Columns.Count == 0)
            {
                table.Columns.Add("PO#", typeof(long));
                table.Columns.Add("Order#", typeof(long));
                table.Columns.Add("Carrier", typeof(string));
                table.Columns.Add("Saved?", typeof(bool));
                table.Columns.Add("Voided?", typeof(bool));
                table.Columns.Add("Remarks", typeof(string));
                table.Columns.Add("ProNumber", typeof(long));
                table.Columns.Add("PrintPath", typeof(string));
            }
            PrintPath = $@"https://copilot2.pilotdelivers.com/webairbill/reprint.aspx?section=ship&subsection=reprint&PrintWhat=Pro&ShipmentID={ProNumber}&ShipperCountry=US&ConsigneeCountry=US";
            table.Rows.Add(PO, Order, Carrier, _IsSaved, _IsVoid, Remarks, ProNumber, PrintPath);

            _IsSaved = _IsVoid = false;
            Remarks = ProNumber = PrintPath = string.Empty;
        }

        public void ExportToExcel()
        {
            DataSet ds = new DataSet("PilotResponse");
            ds.Tables.Add(table);

            //Creae an Excel application instance
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(txtResponsePath.Text);

            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = $"Pilot Response_{DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace(" ", "")}";

                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }
            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApp.Quit();
        }
    }
}
