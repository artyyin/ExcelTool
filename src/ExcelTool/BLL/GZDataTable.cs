using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace ExcelTool
{
    
    public class GZDataTable : ExpandDataTable
    {
        public string DateStr;
        public string TableTile;
        public GZDataTable(string dstr, string title, DataTable dt)
        {
            DateStr = dstr;
            TableTile = title;
            base.Merge(dt);
        }
        public static GZDataTable ImportFromExcel(string filepath)
        {
            ExcelDataHelper excel = new ExcelDataHelper(filepath);
            DataSet ds = excel.ToDataSet(true, 2);
            GZDataTable gzdt = new GZDataTable(ds.DataSetName, ds.Tables[0].TableName, ds.Tables[0]);
            return gzdt;
            //return NPOIHelper.Import(filepath);
        }
        public static GZDataTable ImportFromExcel(string filepath,string sheetname)
        {
            ExcelDataHelper excel = new ExcelDataHelper(filepath);
            DataTable dt = excel.ToDataTable(sheetname, true, 2);           

            GZDataTable gzdt = new GZDataTable(Path.GetFileName(filepath),sheetname, dt);
            return gzdt;
            //return NPOIHelper.Import(filepath);
        }

        public void ExportToExcel(string filename)
        {
            ExcelDataHelper excel = new ExcelDataHelper();
            DataSet ds = new System.Data.DataSet();
            ds.Tables.Add(this);
            excel.FromDataSet(ds, true);
            //excel.SaveAs("c_" + TableTile + DateStr + ".xls");
            excel.SaveAs(filename);
            //NPOIHelper.Export(this, TableTile, DateStr, "c_" + TableTile+DateStr+".xls");
        }
    }
}
