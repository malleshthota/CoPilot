﻿using System;
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
            ds.Shipment[0].PayType = "";
            ds.Shipment[0].SpecialInstructions = "";
            ds.Shipment[0].ReadyTime = System.DateTime.Now;
            ds.Shipment[0].CloseTime = System.DateTime.Now;
            ds.Shipment[0].IsScreeningConsent = "";


            //populate shipper information;
            ds.Shipper[0].Name = "DAVE KASARSKY";
            ds.Shipper[0].Address1 = "PILOT FREIGHT SERVICES";
            ds.Shipper[0].Address2 = "314 N MIDDLETOWN RD";
            ds.Shipper[0].City = "LIMA";
            ds.Shipper[0].State = "PA";
            ds.Shipper[0].Zipcode = "19037";
            ds.Shipper[0].Country = "US";
            ds.Shipper[0].Contact = "DAVE KASARSKY";
            ds.Shipper[0].Phone = "610-891-8100";

            ds.Consignee[0].Name = "PILOT AIR FREIGHT (ORD)";
            ds.Consignee[0].Address1 = "695 TOUHY";
            ds.Consignee[0].City = "ELK GROVE VILLAGE";
            ds.Consignee[0].State = "IL";
            ds.Consignee[0].Zipcode = "60007";
            ds.Consignee[0].Country = "US";
            ds.Consignee[0].Contact = "";
            ds.Consignee[0].Phone = "630-238-3245";

            ds.ThirdParty[0].Name = "BILL PAYER INC";
            ds.ThirdParty[0].Address1 = "123 MONEYBAGS LANE";
            ds.ThirdParty[0].Address2 = "SUITE 100";
            ds.ThirdParty[0].City = "BEVERLY HILLS";
            ds.ThirdParty[0].State = "CA";
            ds.ThirdParty[0].Zipcode = "90210";
            ds.ThirdParty[0].Country = "US";
            ds.ThirdParty[0].Contact = "MR. MONOPOLY";
            ds.ThirdParty[0].Phone = "123-456-7890";

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
                Console.Write("Shipment has been saved.");
                res.dsResult.WriteXml(@"G:\Mallesh\Coding\result.xml");

                //return true;
            }
            else
            {
                //return false;
            }

        }

    }
}