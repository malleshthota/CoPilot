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
        string Remarks = "", ProNumber = "";

        DataTable table = new DataTable();

        public Form1()
        {
            InitializeComponent();
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtFilePath.Text.Trim() == "")
                {
                    MessageBox.Show("OOOPS!!! First Select File to Submit !");
                    return;
                }
                DialogResult res = MessageBox.Show("Are you Sure you want to Submit ?", "Ready To Submit!", MessageBoxButtons.YesNo);
                if (res.Equals(DialogResult.Yes))
                {
                    btnSubmit.Enabled = false;
                    lblInfo.Text = "Please be patient while processing your request !!";
                    DataTable ExcelData = ReadExcel();
                    bool _IsFirstRow = true;
                    foreach (DataRow dr in ExcelData.Rows)
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
                    ExportToExcel();
                }
            }
            catch (Exception ex)
            {
                lblInfo.Text = "Exception while submitting!";
                Remarks += $"BtnSubmit - {ex}";
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            lblUserName.Text = "Welcome Jia Guo";
            lblDateTime.Text = DateTime.Now.DayOfWeek.ToString() + " , " + DateTime.Now.ToShortDateString();
            lblInfo.Text = "Browse File to Submit";
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            lblInfo.Text = "";
            try
            {
                int size = -1;
                openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.csv";
                DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
                if (result == DialogResult.OK) // Test result.
                {
                    fileName = openFileDialog1.FileName;
                    txtFilePath.Text = openFileDialog1.FileName;
                    btnSubmit.Enabled = true;
                    lblInfo.Text = "Click on Submit";
                }
            }
            catch (Exception Ex)
            {
                lblInfo.Text = "Exception while Browsing file";
                Remarks += $"BtnBrowse - {Ex}";
            }
        }

        public DataTable ReadExcel()
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();

            //if (FileExtensin == "xlsx" || FileExtensin == ".xls" || FileExtensin == "xlsm")
            //    conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
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
                    Remarks += $"ReadExcel - {ex}";
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

                ds.Shipment[0].IsScreeningConsent = ""; // Doubt
                ds.Shipment[0].PayType = "";   // Doubt
                ds.Shipment[0].Service = "BASIC"; // Doubt
                ds.Shipment[0].ProductName = "B";
                ds.Shipment[0].ProductDescription = dr.ItemArray[18].ToString(); // Item Description
                ds.Shipment[0].SpecialInstructions = ""; // Doubt
                ds.Shipment[0].ShipperRef = "";
                ds.Shipment[0].ConsigneeRef = "conref987";
                ds.Shipment[0].ConsigneeAttn = dr.ItemArray[6].ToString(); // Customer Phone Number
                ds.Shipment[0].ThirdPartyAuth = "";
                ds.Shipment[0].ShipDate = Convert.ToDateTime(dr.ItemArray[3]); // Ship By
                ds.Shipment[0].ReadyTime = System.DateTime.Now; // Doubt
                ds.Shipment[0].CloseTime = System.DateTime.Now; // Doubt
                ds.Shipment[0].EmailBOL = ""; // Doubt
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
                ds.Consignee[0].Address1 = dr.ItemArray[5].ToString();  // Customer Shipping Address
                ds.Consignee[0].City = dr.ItemArray[10].ToString();     // City
                ds.Consignee[0].State = dr.ItemArray[11].ToString();    // State
                ds.Consignee[0].Zipcode = dr.ItemArray[12].ToString();  // Zip
                ds.Consignee[0].Country = "US"; // Doubt
                ds.Consignee[0].Contact = "";  // Doubt ??
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
                ds.ThirdParty[0].Contact = "877-549-0409";
                ds.ThirdParty[0].Phone = "877-549-0409";
                ds.ThirdParty[0].Email = string.Empty;
                #endregion - Third Party Details Remain Same for all - End

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
                table.Columns.Add("PO#", typeof(string));
                table.Columns.Add("Order#", typeof(string));
                table.Columns.Add("Carrier", typeof(string));
                table.Columns.Add("Saved?", typeof(bool));
                table.Columns.Add("Voided?", typeof(bool));
                table.Columns.Add("Remarks", typeof(string));
                table.Columns.Add("ProNumber", typeof(string));
            }
            table.Rows.Add(PO, Order, Carrier, _IsSaved, _IsVoid, Remarks, ProNumber);

            _IsSaved = _IsVoid = false;
            Remarks = ProNumber = string.Empty;
        }

        public void ExportToExcel()
        {
            DataSet ds = new DataSet("Organization");
            ds.Tables.Add(table);

            //Creae an Excel application instance
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            string ImportExcelPath = @"F:\PilotResponse";
            if (!Directory.Exists(@"C:"))
            {
                ImportExcelPath = @"D:\PilotResponse";
            }
            if (!Directory.Exists(ImportExcelPath))
                Directory.CreateDirectory(ImportExcelPath);

            string ImportExcelFileName = $"PilotResponse_{DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace(" ", "")}.xlsx";
            string FilePath = $@"{ImportExcelPath}\{ImportExcelFileName}";
            File.Create(FilePath);
            FileIOPermission f2 = new FileIOPermission(FileIOPermissionAccess.Write, FilePath);

            Microsoft.Office.Interop.Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(FilePath);

            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

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
