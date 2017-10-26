using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShippingAPIIntegration
{
    public class Program
    {
        public static void Main(string[] args)
        {
            newrest.dsShipment ds;
            newrest.ShipmentService ws = new newrest.ShipmentService();
            //returns dsShipment with default values
            ds = ws.GetNewShipment();

            ds.Shipment[0].LocationID = 11223344;
            ds.Shipment[0].ControlStation = "LAX";
            ds.Shipment[0].AddressID = 12345;
            ds.Shipment[0].TariffHeaderID = 98765;
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
            newrest.dsShipment.LineItemsRow drline = ds.LineItems.NewLineItemsRow();
            drline.Pieces = 1;
            drline.Weight = 600;
            drline.Description = "Pallet";
            drline.Length = 48;
            drline.Width = 48;
            drline.Height = 36;
            ds.LineItems.Rows.Add(drline);

            ///save the shipment
            newrest.PilotShipmentResult res = ws.Save(ds);

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

            newrest.dsVoid ds2;
            //returns dsVoid with default values
            ds2 = ws.GetNewVoid();
            ds2.Void[0].LocationID = 11223344;
            ds2.Void[0].ControlStation = "LAX";
            ds2.Void[0].AddressID = 12345;
            ds2.Void[0].TariffHeaderID = 98765;
            ds2.Void[0].ProNumber = res.dsResult.Shipment[0].ProNumber.ToString(); //"062002488";
            //void the shipment 
            newrest.PilotShipmentResult res2 = ws.Void(ds2);
            if (!res2.IsError && res2.Message== "Shipment Void Success")
            {
                Console.WriteLine("Shipment Void Success");
            }
            else
            {
                Console.WriteLine("Shipment void Failed !.");
            }

        }

    }
}
