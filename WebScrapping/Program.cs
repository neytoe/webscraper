using HtmlAgilityPack;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace WebScrapping
{
    class Program
    {
        // https://www.cbn.gov.ng/rates/GovtSecuritiesDrillDown.asp?beginrec=11&endrec=20&market=
       
        static void Main(string[] args)
        {
            int page = 200;
            var util = new Utils();
            string result = util.GetRecords(page);
            var objectChange = util.ChangeToObject(result);
            util.CreateSpreadsheet(objectChange);
        }
    }
}

