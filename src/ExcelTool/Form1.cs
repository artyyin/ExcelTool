using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTool
{
    public partial class Form1 : Form
    {
        private string[] RejectNames = new string[] { "单位代码", "单位名称", "类别代码", "类别名称", "个人账号", "身份证号码", "代发银行名称" };
        public Form1()
        {
            InitializeComponent();
        }

        GZDataTablesYear gz = new GZDataTablesYear();
        private void button4_Click(object sender, EventArgs e)
        {
            string[] fnames = GetSelectedFiles();
            gz.ImportFromExcel(fnames.ToArray(), RejectNames);
            string[] colnames = gz.GetColumnsName();
            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(colnames);
            checkedListBox3.Items.Clear();
            checkedListBox4.Items.Clear();
            checkedListBox3.Items.AddRange(colnames);
            checkedListBox4.Items.AddRange(colnames);
            this.dataGridView2.AutoGenerateColumns = true;
            this.dataGridView2.DataSource = gz.GetColumnsMatrix();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<string> reserved = GetCheckedItemString(checkedListBox3);
            List<string> expended = GetCheckedItemString(checkedListBox4);
            expended.Remove(comboBox1.Text);
            reserved.Remove(comboBox1.Text);
            gz.Summary(comboBox1.Text, reserved.ToArray(), expended.ToArray());
            this.dataGridView1.DataSource = gz.SummaryTable;
            gz.ExportToExcel();
        }

        private void 汇总ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 选择文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Dir;
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                treeView1.Nodes.Clear();
                Dir = folderBrowserDialog1.SelectedPath;
                ShowFilesInTreeView(Dir);
                Directory.SetCurrentDirectory(Dir);
                CreateExcelTree();
            }
        }

        private void 精简ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<TabPage> tabs = new List<TabPage>();
            foreach (TabPage tab in tabControl2.TabPages)
            {
                tabs.Add(tab);
            }            
            tabControl2.TabPages.Clear();
            foreach (TabPage tab in tabs)
            {
                tab.Dispose();
            }
            tabs.Clear();

            foreach (var item in GetSelected())
            {
                CompressDatatable(item[0],item[1]);
            }
        }
        private string[][] GetSelected()
        {
            List<string[]> result = new List<string[]>();
            TreeNode root = this.treeView1.Nodes[0];
            foreach (TreeNode fnode in root.Nodes)
            {
                foreach (TreeNode snode in fnode.Nodes)
                {
                    if (snode.Checked)
                    {
                        string[] one = new string[2] {fnode.Tag as string, snode.Text };
                        result.Add(one);
                    }
                }
            }
            return result.ToArray();
        }
        private void CompressDatatable(string filename,string sheetname)
        {
            GZDataTable gzdt = GZDataTable.ImportFromExcel(filename,sheetname);
            gzdt.DeleteBlankOrZeroColumns(RejectNames);

            toolStripStatusLabel1.Text = gzdt.Rows.Count.ToString();
            string newfilename = Path.GetDirectoryName(filename) + "\\" + Path.GetFileNameWithoutExtension(filename) + "[精简]" + Path.GetExtension(filename);
            gzdt.ExportToExcel(newfilename);

            DynamicShowDatatable(filename,sheetname, gzdt);
        }

        private void DynamicShowDatatable(string item,string sheetname, GZDataTable gzdt)
        {
            DataGridView dataGV = new DataGridView();
            dataGV.AutoGenerateColumns = true;
            dataGV.DataSource = gzdt;
            dataGV.AutoResizeRows();
            dataGV.AllowUserToAddRows = false;
            dataGV.AllowUserToDeleteRows = false;
            dataGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGV.Dock = System.Windows.Forms.DockStyle.Fill;
            dataGV.Location = new System.Drawing.Point(0, 0);
            dataGV.Name = "dataGridView2";
            dataGV.ReadOnly = true;
            dataGV.RowTemplate.Height = 23;
            dataGV.Size = new System.Drawing.Size(321, 349);

            TabPage tabPage = new TabPage();
            tabPage.Controls.Add(dataGV);
            tabPage.Location = new System.Drawing.Point(4, 22);
            tabPage.Padding = new System.Windows.Forms.Padding(3);
            tabPage.Size = new System.Drawing.Size(587, 355);
            tabPage.TabIndex = 0;
            tabPage.Name = Path.GetFileNameWithoutExtension(item) +" >> "+ sheetname;
            tabPage.Text = Path.GetFileNameWithoutExtension(item) +" >> "+ sheetname;
            tabPage.UseVisualStyleBackColor = true;
            tabControl2.TabPages.Add(tabPage);
        }
        #region 功能性代码
        private void ShowFilesInTreeView(string Dir)
        {
            string[] filelist = (new DirHelper(Dir)).GetAllFiles("*.xls");
            TreeNode root = new TreeNode(Dir);
            treeView1.Nodes.Add(root);
            foreach (string file in filelist)
            {
                TreeNode node = new TreeNode();
                node.Name = Path.GetFileName(file);
                node.Text = Path.GetFileName(file);
                node.Tag = file;
                root.Nodes.Add(node);
            }
            root.Expand();
        }
        private string[] GetSelectedFiles()
        {
            List<string> result = new List<string>();
            if (treeView1.Nodes.Count > 0)
            {
                TreeNode root = treeView1.Nodes[0];
                foreach (TreeNode node in root.Nodes)
                {
                    if (node.Checked)
                    {
                        result.Add(node.Tag as string);
                    }
                }
            }
            return result.ToArray();
        }
        private List<string> GetCheckedItemString(CheckedListBox chklist)
        {
            List<string> expended = new List<string>();
            foreach (string str in chklist.CheckedItems)
            {
                expended.Add(str);
            }
            return expended;
        }
        private void CreateExcelTree()
        {
            List<ExcelDataHelper> Excels = new List<ExcelDataHelper>();

            foreach (TreeNode rnode in treeView1.Nodes[0].Nodes)
            {
                string fname = rnode.Tag as string;
                TreeNode node = treeView1.Nodes[0].Nodes[Path.GetFileName(fname)];
                ExcelDataHelper edh = new ExcelDataHelper(fname);
                Excels.Add(edh);
                string[] sheetnames = edh.GetSheetNames();

                foreach (string sheet in sheetnames)
                {
                    TreeNode cnode = new TreeNode();
                    cnode.Text = sheet;
                    cnode.Name = sheet;
                    node.Nodes.Add(cnode);
                    //GZDataTable gzdt = GZDataTable.ImportFromExcel(fname, sheet);
                    //string[] cols = gzdt.GetColumnsName();
                    //foreach (string col in cols)
                    //{
                    //    cnode.Nodes.Add(col);
                    //}
                }
            }
        }
        #endregion

        private void autoSizeToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
