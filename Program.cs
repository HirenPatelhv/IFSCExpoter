

using ClosedXML.Excel;
using IFSCUpdater;
using Newtonsoft.Json;
using System;
using System.Data;
using System.Data.SqlClient;

string folder = @"C:\Users\admin\Desktop\by-bank\Bank";
List<Info> bankInfos = new List<Info>();

//**************************************************************************
//**************************************************************************
//**************************************************************************
//                  For Start With File Name
//**************************************************************************
//**************************************************************************
//**************************************************************************


//foreach (var item in Directory.EnumerateFiles(folder, "*.json"))
//{
//    var filename = Path.GetFileNameWithoutExtension(item);
//    List<string> strings = new List<string>();
//    var strInfo = File.ReadAllLines(item);
//    int i = 0;
//    foreach (string item1 in strInfo)
//    {


//        if (item1.Trim().StartsWith("\"" + filename))
//        {
//            strings.Add("{");
//        }
//        else
//        {
//            if (i == 0)
//            {
//                strings.Add("[");
//            }
//            else if (i == strInfo.Length - 1)
//            {
//                strings.Add("]");
//            }
//            else
//            {
//                strings.Add(item1);
//            }
//        }
//        i++;
//    }

//    var final = string.Join("", strings);

//    bankInfos.AddRange(JsonConvert.DeserializeObject<List<Info>>(final));

//}

//**************************************************************************
//**************************************************************************
//**************************************************************************
//                  Start With " AND End With {
//**************************************************************************
//**************************************************************************
//**************************************************************************

foreach (var item in Directory.EnumerateFiles(folder, "*.json"))
{
    var filename = Path.GetFileNameWithoutExtension(item);
    List<string> strings = new List<string>();
    var strInfo = File.ReadAllLines(item);
    int i = 0;
    foreach (string item1 in strInfo)
    {


        if (item1.Trim().StartsWith("\"") && item1.Trim().EndsWith("{"))
        {
            strings.Add("{");
        }
        else
        {
            if (i == 0)
            {
                strings.Add("[");
            }
            else if (i == strInfo.Length - 1)
            {
                strings.Add("]");
            }
            else
            {
                strings.Add(item1);
            }
        }
        i++;
    }

    var final = string.Join("", strings);

    bankInfos.AddRange(JsonConvert.DeserializeObject<List<Info>>(final));

}
//**************************************************************************
//**************************************************************************
//**************************************************************************


DataTable dataTable =new DataTable();
dataTable.Columns.Add("BankCode", typeof(string));
dataTable.Columns.Add("BranchName", typeof(string));
dataTable.Columns.Add("Address1", typeof(string));
dataTable.Columns.Add("Address2", typeof(string));
dataTable.Columns.Add("City", typeof(string));
dataTable.Columns.Add("Phone1", typeof(string));
dataTable.Columns.Add("Phone2", typeof(string));
dataTable.Columns.Add("IFSCCode", typeof(string));
dataTable.AcceptChanges();



foreach (var item in bankInfos)
{
    var datarow = dataTable.NewRow();

    datarow["BankCode"] = item.MICR;
    datarow["BranchName"] = item.BANK + " - " + item.BRANCH;
    datarow["Address1"] = item.ADDRESS;
    datarow["Address2"] = string.Empty;
    datarow["City"] = item.CITY;
    datarow["Phone1"] = item.CONTACT;
    datarow["Phone2"] = string.Empty;
    datarow["IFSCCode"] = item.IFSC;
    dataTable.Rows.Add(datarow);
    dataTable.AcceptChanges();
}

XLWorkbook wb = new XLWorkbook();
wb.Worksheets.Add(dataTable, "IFSC Details");
wb.SaveAs(folder+"\\BankInfoMICRIDFC.xlsx");

Console.WriteLine(bankInfos.Count.ToString());