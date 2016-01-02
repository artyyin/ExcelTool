using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    public class ExcelDataHelper
    {
        public string ExcelFilePath { get; private set; }
        private IWorkbook Workbook;

        public ExcelDataHelper()
        {
            ExcelFilePath = "";
            Workbook = new HSSFWorkbook();
        }
        public ExcelDataHelper(string filepath)
        {
            ExcelFilePath = filepath;
            if (Path.GetExtension(filepath) == ".xls")
            {
                using (FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Read))
                {
                    Workbook = new HSSFWorkbook(file);
                }
            }
            else if (Path.GetExtension(filepath) == ".xlsx")
            {
                using (FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Read))
                {
                    Workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(file);
                }
            }
        }                        
        
        //获取Excel文件的扩展名
        private string GetExcelExtension()
        {
            if (Workbook is NPOI.XSSF.UserModel.XSSFWorkbook)
            {
                return ".xlsx";
            }
            else
            {
                return "xls";
            }
        }
        //获取所有Sheet的Name
        public string[] GetSheetNames()
        {
            List<string> result = new List<string>();
            foreach (ISheet sheet in Workbook)
            {
                result.Add(sheet.SheetName);
            }
            return result.ToArray();
        }
        //去除空格及特殊字符'\0'
        private string Trim(string s)
        {
            s = s.Trim();
            s = s.Replace("\0", "");
            return s;
        }
        //添加文件属性信息
        private void AddSummaryInformation()
        {
            if (Workbook is HSSFWorkbook)
            {
                HSSFWorkbook w = (Workbook as HSSFWorkbook);
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "NPOI";
                w.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "文件作者信息"; //填加xls文件作者信息
                si.ApplicationName = "创建程序信息"; //填加xls文件创建程序信息
                si.LastAuthor = "最后保存者信息"; //填加xls文件最后保存者信息
                si.Comments = "作者信息"; //填加xls文件作者信息
                si.Title = "标题信息"; //填加xls文件标题信息
                si.Subject = "主题信息";//填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                w.SummaryInformation = si;
            }
            else if (Workbook is NPOI.XSSF.UserModel.XSSFWorkbook)
            {
                var w = (Workbook as NPOI.XSSF.UserModel.XSSFWorkbook);
            }
        }

        #region Excel To Data
        //从Sheet转换到数据表格
        public DataTable ToDataTable(ISheet sheet, bool HasTitle, int TitleRowIndex)
        {
            DataTable dt = new DataTable();
            dt.TableName = sheet.SheetName;
            int RowIndex = sheet.FirstRowNum + TitleRowIndex;
            if (HasTitle)
            {
                IRow irow = sheet.GetRow(RowIndex);
                for (int i = 0; i < irow.LastCellNum; i++)
                {
                    dt.Columns.Add(Trim(irow.Cells[i].ToString()));
                }
                RowIndex++;
            }
            else
            {
                IRow irow = sheet.GetRow(RowIndex);
                for (int i = 0; i < irow.LastCellNum; i++)
                {
                    dt.Columns.Add("列" + i.ToString("00"));
                }
            }
            for (int i = RowIndex; i < sheet.LastRowNum; i++)
            {
                DataRow drow = dt.NewRow();
                IRow irow = sheet.GetRow(i);
                for (int j = 0; j < irow.Cells.Count; j++)
                {
                    drow[j] = Trim(irow.GetCell(j).ToString());
                }
                dt.Rows.Add(drow);
                RowIndex++;
            }
            return dt;
        }
        public DataTable ToDataTable(string sheetname, bool hasTitle, int TitleRowIndex)
        {
            ISheet sheet = Workbook.GetSheet(sheetname);
            return ToDataTable(sheet, hasTitle, TitleRowIndex);
        }
        //从Excel转换为DataSet
        public DataSet ToDataSet(bool HasTitle, int FirstRowIndex)
        {
            return ToDataSet(Workbook, HasTitle, FirstRowIndex);
        }
        private DataSet ToDataSet(IWorkbook book, bool HasTitle, int FirstRowIndex)
        {
            DataSet ds = new DataSet();
            ds.DataSetName = Path.GetFileName(ExcelFilePath);
            foreach (ISheet sheet in book)
            {
                DataTable dt = ToDataTable(sheet, HasTitle, FirstRowIndex);
                ds.Tables.Add(dt);
            }
            return ds;
        }
        #endregion

        #region Data To Excel
        public ISheet FromDataTable(DataTable dt, int FirstRowIndex, bool OutputTitle)
        {
            Debug.Assert(FirstRowIndex >= 0);
            Debug.Assert(dt != null);
            //设置SheetName
            HSSFSheet sheet = Workbook.CreateSheet() as HSSFSheet;
            int sheetidx = Workbook.GetSheetIndex(sheet);
            Workbook.SetSheetName(sheetidx, dt.TableName);

            int RowIndex = BeforeTitle(sheet);
            //输出标题行
            RowIndex += FirstRowIndex;
            if (OutputTitle)
            {
                IRow row = sheet.CreateRow(RowIndex);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    ICell cell = row.CreateCell(i);
                    cell.SetCellValue(dt.Columns[i].ColumnName);
                }
                RowIndex++;
            }
            //输出数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet.CreateRow(RowIndex);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    //TODO::需要针对类型转换                   
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
                RowIndex++;
            }
            AfterFootbar(sheet);
            return sheet;
        }
        public IWorkbook FromDataSet(DataSet set, bool ClearOldData)
        {
            if (ClearOldData)
            {
                while (Workbook.NumberOfSheets > 0)
                    Workbook.RemoveSheetAt(0);
            }
            foreach (DataTable dt in set.Tables)
            {
                FromDataTable(dt, 0, true);
            }
            return Workbook;
        }
        
        protected virtual int BeforeTitle(ISheet sheet)
        {
            return 0;
        }
        
        protected virtual int AfterFootbar(ISheet sheet)
        {
            return 0;
        }

        #endregion

        public void SaveAs(string strFileName)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                Workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
        }
    }
}
