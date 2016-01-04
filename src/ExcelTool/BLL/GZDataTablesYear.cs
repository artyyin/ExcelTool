using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{   
    public class GZDataTablesYear
    {
        private Dictionary<string, GZDataTable> YearTables;
        public DataTable SummaryTable;       
        public void Summary(string key_column, string[] reserve_columns, string[] expend_columns)
        {
            List<string> months = new List<string>(YearTables.Keys);
            months.Sort();
            HashSet<string> keyset = new HashSet<string>();
            foreach (string month in months)
            {
                GZDataTable gzdt = YearTables[month];
                foreach (DataRow row in gzdt.Rows)
                {
                    keyset.Add(row[key_column].ToString());
                }
            }
            List<string> Keys = new List<string>(keyset);
            DataTable result = new DataTable();
            result.Columns.Add(key_column);            
            foreach (var col in reserve_columns)
            {
                result.Columns.Add(col);
            }
            foreach (string month in months)
            {
                //result.Columns.Add(expandcol);
                foreach (var expandcol in expend_columns)
                {
                    result.Columns.Add(month + expandcol);
                }
            }
            months.Sort();
            Keys.Sort();
            //填充数据
            foreach (string key in Keys)
            {
                DataRow row = result.NewRow();
                row[key_column] = key;
                foreach (string month in months)
                {
                    GZDataTable gzdt = YearTables[month];
                    foreach (DataRow mrow in gzdt.Rows)
                    {
                        if (mrow[key_column].ToString() == key)
                        {
                            foreach (var col in reserve_columns)
                            {
                                row[col] = mrow[col];
                            }
                            foreach (string expandcol in expend_columns)
                            {
                                if (mrow.Table.Columns.Contains(expandcol))
                                    row[month + expandcol] = mrow[expandcol];
                                else
                                    row[month + expandcol] = "";
                            }
                            break;
                        }
                    }
                }
                result.Rows.Add(row);
            }
            SummaryTable = result;
        }

        #region 输入输出
        public void ExportToExcel()
        {
            NPOIHelper.Export(SummaryTable, "", "", "c_汇总.xls");            
        }
        public void ImportFromExcel(string[] filepaths, string[] RejectNames)
        {
            YearTables = new Dictionary<string, GZDataTable>();
            foreach (string file in filepaths)
            {
                GZDataTable gzdt = GZDataTable.ImportFromExcel(file);
                gzdt.DeleteBlankOrZeroColumns(RejectNames);
                YearTables.Add(gzdt.DateStr, gzdt);
            }
        }
        #endregion

        public string[] GetColumnsName()
        {
            HashSet<string> result = new HashSet<string>();
            foreach (var item in YearTables)
            {
                foreach (DataColumn dc in item.Value.Columns)
                {
                    result.Add(dc.ColumnName);
                }
            }
            return result.ToArray();
        }        
        public DataTable GetColumnsMatrix()
        {
            DataTable d = new DataTable();
            d.Columns.Add("字段");
            string[] cols = GetColumnsName();
            List<string[]> collist = new List<string[]>();
            foreach (var item in YearTables)
            {
                d.Columns.Add(item.Value.TableTile + item.Value.DateStr);
                collist.Add(item.Value.GetColumnsName());
            }
            foreach (string col in cols)
	        {
                DataRow drow = d.NewRow();
                drow[0] = col;
                int i=1;
                foreach (string[] col_l in collist)
                {
                    if (Array.IndexOf<string>(col_l, col) >= 0)
                    {
                        drow[i] = "●";
                    }
                    i++;
                }
                d.Rows.Add(drow);
	        }            
            return d;
        }
    }
}
