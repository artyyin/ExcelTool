using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ExcelTool
{
    public class ExpandDataTable : DataTable
    {
        public string[] GetColumnsName()
        {
            HashSet<string> result = new HashSet<string>();
            foreach (DataColumn dc in base.Columns)
            {
                result.Add(dc.ColumnName);
            }
            return result.ToArray();
        }
        public void DeleteBlankOrZeroColumns(string[] RejectNames)
        {
            Filter();
            DeleteColumns(RejectNames);
        }
        public void DeleteColumns(string[] RejectNames)
        {
            foreach (string s in RejectNames)
            {
                if (Columns.IndexOf(s) >= 0)
                    Columns.Remove(s);
            }
        }
        private void Filter()
        {
            HashSet<int> ccn = new HashSet<int>();
            for (int i = 0; i < Rows.Count; i++)
            {
                for (int j = 0; j < Columns.Count; j++)
                {
                    double x;
                    if (double.TryParse(Rows[i][j].ToString(), out x))
                    {
                        if (Math.Abs(x) > 1e-6)
                        {
                            ccn.Add(j);
                        }
                    }
                    else
                    {
                        ccn.Add(j);
                    }
                }
            }
            List<int> cn = new List<int>(ccn);
            cn.Sort();

            DataTable d = new DataTable();
            foreach (int x in cn)
            {
                d.Columns.Add(Columns[x].Caption.Trim().Replace("\0", ""));
            }
            for (int i = 0; i < Rows.Count; i++)
            {
                DataRow row = d.NewRow();
                for (int j = 0; j < cn.Count; j++)
                {
                    row[j] = Rows[i][cn[j]].ToString().Trim().Replace("\0", "");
                }
                d.Rows.Add(row);
            }
            base.Clear();
            base.Merge(d);
        }
    }
    public class ExpandDataSet : DataSet
    {
        public string[] GetColumnsName()
        {
            HashSet<string> result = new HashSet<string>();
            foreach (DataTable item in base.Tables)
            {
                foreach (DataColumn dc in item.Columns)
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
            foreach (DataTable item in base.Tables)
            {
                d.Columns.Add(item.TableName);
                List<string> ccols = new List<string>();
                foreach (DataColumn dc in item.Columns)
                {
                    //result.Add(dc.ColumnName);
                    ccols.Add(dc.ColumnName);
                }
                collist.Add(ccols.ToArray());
            }

            foreach (string col in cols)
            {
                DataRow drow = d.NewRow();
                drow[0] = col;
                int i = 1;
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
