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

        private void button1_Click(object sender, EventArgs e)
        {
            //checkedListBox1.Items.Clear();
            treeView1.Nodes.Clear();
            string Dir = textBox1.Text;
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Dir = folderBrowserDialog1.SelectedPath;
                textBox1.Text = Dir;
            }
            ShowFilesInTreeView(Dir);
            Directory.SetCurrentDirectory(Dir);
            CreateExcelTree();
        }
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
        private string[] GetCheckedFiles()
        {
            List<string> result = new List<string>();
            TreeNode root = treeView1.Nodes[0];
            foreach (TreeNode node in root.Nodes)
            {
                if (node.Checked)
                {
                    result.Add(node.Tag as string);
                }
            }
            return result.ToArray();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            foreach (var item in GetCheckedFiles())
            {
                GZDataTable gzdt = GZDataTable.ImportFromExcel(item.ToString());
                gzdt.DeleteBlankOrZeroColumns(RejectNames);

                this.dataGridView1.AutoGenerateColumns = true;
                this.dataGridView1.DataSource = gzdt;
                this.dataGridView1.AutoResizeRows();

                toolStripStatusLabel1.Text = gzdt.Rows.Count.ToString();
                gzdt.ExportToExcel();
            }            
        }

        
        GZDataTablesYear gz = new GZDataTablesYear();
        private void button4_Click(object sender, EventArgs e)
        {
            string[] fnames = GetCheckedFiles();
            gz.ImportFromExcel(fnames.ToArray(),RejectNames);
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

        private void CreateExcelTree()
        {
            List<ExcelDataHelper> Excels = new List<ExcelDataHelper>();           
            
            foreach(TreeNode rnode in treeView1.Nodes[0].Nodes)
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
                    GZDataTable gzdt = GZDataTable.ImportFromExcel(fname, sheet);
                    string[] cols = gzdt.GetColumnsName();
                    foreach (string col in cols)
                    {
                        cnode.Nodes.Add(col);
                    }
                }
            }
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
        private List<string> GetCheckedItemString(CheckedListBox chklist)
        {
            List<string> expended = new List<string>();
            foreach (string str in chklist.CheckedItems)
            {
                expended.Add(str);
            }
            return expended;
        }
    }
}
