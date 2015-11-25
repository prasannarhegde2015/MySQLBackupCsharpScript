using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Management.Instrumentation;
using System.Management;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;


namespace getProcessListGrid
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            label1.Text = "";
            terminatexcel();
            progressBar1.Visible = false;
            label1.Visible = false;
            ClearDatagrid.Enabled = false;
            ProcessExport.Enabled = false;
            ProductExport.Enabled = false;
           

        }
        
        #region ProcessCode
        private void GetProcessList_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("sr no");
            dt.Columns.Add("Caption");
            dt.Columns.Add("ProcessId");
            dt.Columns.Add("ExecutablePath");
            dt.Columns.Add("WorkingSetSize_MB");
      
            ManagementObjectSearcher searcher = 
                     new ManagementObjectSearcher("root\\CIMV2",
                     "SELECT * FROM Win32_Process");
            int i = 0;
            foreach (ManagementObject queryObj in searcher.Get())
            {
                DataRow dr = dt.NewRow();
                dr["sr no"] = (i + 1).ToString();
                dr["Caption"] = queryObj["Caption"];
                dr["ProcessId"] = queryObj["ProcessId"];
                dr["ExecutablePath"] = queryObj["ExecutablePath"];
                dr["WorkingSetSize_MB"] = Convert.ToDouble(queryObj["WorkingSetSize"]) / 1000000 ;
                dt.Rows.Add(dr);
                i++;  
            }
            dataGridView1.DataSource = dt;
            dataGridView1.ReadOnly = true;

            if (dataGridView1.RowCount > 1)
            {
                ClearDatagrid.Enabled = true;
                ProcessExport.Enabled = true;
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {

            string strname = "";
            DialogResult dlgr = saveFileDialog1.ShowDialog();
            saveFileDialog1.Filter = "Excel File|*.xls,*.xlsx";
            saveFileDialog1.Title = "Save Excel File";

            if (dlgr == DialogResult.OK) // Test result.
            {
                strname = saveFileDialog1.FileName;
            }
            else if ((dlgr == DialogResult.Cancel))
            {
                //MessageBox.Show("Export Cancelled by user");
                return;
            }

            Excel.Application oxls = new Excel.Application();
            Excel.Workbook owb = oxls.Workbooks.Add();
            Excel.Worksheet ows = owb.ActiveSheet;
            ows.Name = "Process List";
            ows.Cells[1, 1] = "sr no";
            ows.Cells[1, 2] = "Caption";
            ows.Cells[1, 3] = "ProcessId";
            ows.Cells[1, 4] = "ExecutablePath";
            ows.Cells[1, 5] = "WorkingSetSize_MB";

            DataTable td1 = new DataTable();
            td1 = (DataTable)dataGridView1.DataSource;
            for (int i = 1; i < td1.Rows.Count; i++)
            {
                ows.Cells[i + 1, 1] = td1.Rows[i]["Sr No"].ToString();
                ows.Cells[i + 1, 2] = td1.Rows[i]["Caption"].ToString();
                ows.Cells[i + 1, 3] = td1.Rows[i]["ProcessId"].ToString();
                ows.Cells[i + 1, 4] = td1.Rows[i]["ExecutablePath"].ToString();
                ows.Cells[i + 1, 5] = td1.Rows[i]["WorkingSetSize_MB"].ToString();
            }
            oxls.DisplayAlerts = false;
            owb.SaveAs(strname);
            oxls.Quit();
            //   MessageBox.Show("Successfully Exported to Excel");
            label1.Visible = true;
            label1.ForeColor = Color.Indigo;
            label1.Text = "Successfully Exported Process List to Excel";
            terminatexcel();
            saveFileDialog1.Dispose();
        }
       
        #endregion

        #region Product code

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("sr no");
            dt.Columns.Add("Product Name");
            dt.Columns.Add("Vendor");
            dt.Columns.Add("Version");
            dt.Columns.Add("Install Date");
            dt.Columns.Add("Description");
            progressBar1.Visible = true;
            label1.Visible = true;
            label1.Text = "Please Wait....";
            label1.ForeColor = Color.BlueViolet;
            ManagementObjectSearcher searcher =
                     new ManagementObjectSearcher("root\\CIMV2",
                     "SELECT * FROM Win32_Product");
            int i = 0;

            progressBar1.Minimum = 0;
            progressBar1.Maximum = searcher.Get().Count;
            progressBar1.Value = 0;
            label1.Text = "Processing...";
            label1.ForeColor = Color.IndianRed;

            foreach (ManagementObject queryObj in searcher.Get())
            {

                DataRow dr = dt.NewRow();
                dr["sr no"] = (i + 1).ToString();
                dr["Product Name"] = queryObj["Caption"];
                dr["Vendor"] = queryObj["Vendor"];
                dr["Version"] = queryObj["Version"];
                dr["Install Date"] = queryObj["InstallDate"];
                dr["Description"] = queryObj["Description"];
                dt.Rows.Add(dr);

                i++;
                progressBar1.Value++;

                progressBar1.ForeColor = Color.Green;

            }
            //   backgroundWorker1.RunWorkerAsync();
            label1.Text = "Processing Complete";
            label1.ForeColor = Color.Green;
            dataGridView1.DataSource = dt;
            dataGridView1.ReadOnly = true;
            if (dataGridView1.RowCount > 1)
            {
                ClearDatagrid.Enabled = true;
                ProductExport.Enabled = true;
            }

            System.Threading.Thread.Sleep(5000);
            progressBar1.Visible = false;
            label1.Visible = false;
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                string strname = "";
                DialogResult dlgr = saveFileDialog2.ShowDialog();
                saveFileDialog2.Filter = "Excel File|*.xls,*.xlsx";
                saveFileDialog2.Title = "Save Excel File";
                if (dlgr == DialogResult.OK) // Test result.
                {
                    strname = saveFileDialog2.FileName;
                }
                else if (dlgr == DialogResult.Cancel)
                {
                    return;
                }


                Excel.Application oxls = new Excel.Application();
                Excel.Workbook owb = oxls.Workbooks.Add();
                Excel.Worksheet ows = owb.ActiveSheet;

                ows.Cells[1, 1] = "Sr No";
                ows.Cells[1, 2] = "Product Name";
                ows.Cells[1, 3] = "Vendor";
                ows.Cells[1, 4] = "Version";
                ows.Cells[1, 5] = "Install Date";
                ows.Cells[1, 6] = "Description";

                DataTable td1 = new DataTable();
                td1 = (DataTable)dataGridView1.DataSource;

                for (int i = 1; i < td1.Rows.Count; i++)
                {
                    ows.Cells[i+1, 1] = td1.Rows[i]["Sr No"].ToString();
                    ows.Cells[i+1, 2] = td1.Rows[i]["Product Name"].ToString();
                    ows.Cells[i+1, 3] = td1.Rows[i]["Vendor"].ToString();
                    ows.Cells[i+1, 4] = td1.Rows[i]["Version"].ToString();
                    ows.Cells[i+1, 5] = td1.Rows[i]["Install Date"].ToString();
                    ows.Cells[i+1, 6] = td1.Rows[i]["Description"].ToString();
                }
                oxls.DisplayAlerts = false;
                owb.SaveAs(strname);
                oxls.Quit();
                terminatexcel();
              //  MessageBox.Show("Successfully Exported to Excel");
                label1.Visible = true;
                label1.Text = "Successfully Exported to Excel..";
                saveFileDialog1.Dispose();
  
            }
            catch (Exception ex)
            {
                terminatexcel();
                MessageBox.Show("Please Generate Products List by clicking [Product List] button before exporting to excel Table generated is process table [ reason: "+ex.Message +"]", "Please check", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }
#endregion
        private void terminatexcel()
        {
             ManagementObjectSearcher searcher = 
                     new ManagementObjectSearcher("root\\CIMV2",
                     "SELECT * FROM Win32_Process");


            foreach (ManagementObject queryObj in searcher.Get())
            {
                if (  queryObj["Caption"].ToString() == "EXCEL.EXE")
                {
                    queryObj.InvokeMethod("Terminate", null);
                }
            }
        }

        private void ClearDatagrid_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            progressBar1.Visible = false;
            label1.Visible = false;
            ProcessExport.Enabled = false;
            ProductExport.Enabled = false;
        }

        
    }
}
-
