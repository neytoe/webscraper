using HtmlAgilityPack;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WebScrapping
{
    public class Utils
    {
        public string GetRecords(int page)
        {
            string data = "";
            int beginning = 1;
            int ending = 10;

            // int.MaxValue
            for (int i = 0; i < 3; i++)
            {
                string url = $"https://www.cbn.gov.ng/rates/GovtSecuritiesDrillDown.asp?beginrec={beginning}&endrec={ending}&market=";
                int count = 0;
                HtmlWeb web = new HtmlWeb();
                HtmlDocument doc = web.Load(url);
                var headerNames = doc.DocumentNode.SelectNodes("//table[@id='mytables']//tr/td");

                foreach (var item in headerNames)
                {
                    if (item.InnerText == "" && count == 0)
                    {
                        data += "+";
                        count++;
                    } 
                    else
                    {
                        data += item.InnerText + "=";
                        count = 0;
                    }                    
                }

                beginning += 10;
                ending += 10;
            }

            return data;
        }

        public List<ObjectModel> ChangeToObject(string data)
        {
            List<ObjectModel> modelObject = new List<ObjectModel>();
            var newData = data.Remove(data.Length - 1, 1);
            var dataArray = newData.Split('+');

            foreach (var item in dataArray)
            {
                if (item == "")
                    continue;
                else
                {
                    ObjectModel model = new ObjectModel();
                    var newItem = item.Remove(item.Length - 1, 1);
                    var itemObject = newItem.Split('=');

                    model.AuctionDate = itemObject[0];
                    model.SecurityType = itemObject[1];
                    model.Tenor = itemObject[2];
                    model.AuctionNo = itemObject[3];
                    model.Auction = itemObject[4];
                    model.MaturityDate = itemObject[5];
                    model.TotalSubscription = itemObject[6];
                    model.TotalSuccessful = itemObject[7];
                    model.RangeBid = itemObject[8];
                    model.SuccessfulBidRates = itemObject[9];
                    model.Description = itemObject[10];
                    model.Rate = itemObject[11];
                    model.TrueYield = itemObject[12];
                    model.AmountOffered = itemObject[13];

                    modelObject.Add(model);
                }
            }

            return modelObject;
        }

        public void CreateSpreadsheet(List<ObjectModel> datas)
        {
            string spreadsheetPath = "securities.xlsx";

            FileInfo spreadsheetInfo = new FileInfo(spreadsheetPath);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Worksheets.Add(spreadsheetPath);
            var securities = pck.Workbook.Worksheets.Add("Securities");

            securities.Cells["A1"].Value = "Auction Date";
            securities.Cells["B1"].Value = "Security Type";
            securities.Cells["C1"].Value = "Tenor";
            securities.Cells["D1"].Value = "Auction No";
            securities.Cells["E1"].Value = "Auction";
            securities.Cells["F1"].Value = "Maturity Date";
            securities.Cells["G1"].Value = "Total Subscription";
            securities.Cells["H1"].Value = "Total Successful";
            securities.Cells["I1"].Value = "Range Bid";
            securities.Cells["J1"].Value = "Successful Bid Rates";
            securities.Cells["K1"].Value = "Description";
            securities.Cells["L1"].Value = "Rate";
            securities.Cells["M1"].Value = "True Yield";
            securities.Cells["N1"].Value = "Amount Offered (mn)";
            securities.Cells["A1:N1"].Style.Font.Bold = true;

            int currentRow = 2;

            foreach (var data in datas)
            {
                securities.Cells["A" + currentRow.ToString()].Value = data.AuctionDate;
                securities.Cells["B" + currentRow.ToString()].Value = data.SecurityType;
                securities.Cells["C" + currentRow.ToString()].Value = data.Tenor;
                securities.Cells["D" + currentRow.ToString()].Value = data.AuctionNo;
                securities.Cells["E" + currentRow.ToString()].Value = data.Auction;
                securities.Cells["F" + currentRow.ToString()].Value = data.MaturityDate;
                securities.Cells["G" + currentRow.ToString()].Value = data.TotalSubscription;
                securities.Cells["H" + currentRow.ToString()].Value = data.TotalSuccessful;
                securities.Cells["I" + currentRow.ToString()].Value = data.RangeBid;
                securities.Cells["J" + currentRow.ToString()].Value = data.SuccessfulBidRates;
                securities.Cells["K" + currentRow.ToString()].Value = data.Description;
                securities.Cells["L" + currentRow.ToString()].Value = data.Rate;
                securities.Cells["M" + currentRow.ToString()].Value = data.TrueYield;
                securities.Cells["N" + currentRow.ToString()].Value = data.AmountOffered;

                currentRow++;
            }

            pck.SaveAs(spreadsheetInfo);
            Console.WriteLine("Save");
        }
    }
}
