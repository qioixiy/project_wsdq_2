using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 

namespace ExportExcel
{
    public partial class FormMain : Form
    {
        EnergyData mEnergyData;
        public FormMain()
        {
            InitializeComponent();
        }

        public void SetEnergyDataFromFile(String filename)
        {
            mEnergyData = new EnergyData(filename);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void buttonExportExcel_Click(object sender, EventArgs e)
        {
            CBExcel excel = new CBExcel();
            excel.Create();
            
            excel.SelectWorksheet(1);
            excel.SetData(1, 1, "");
            excel.SetData(1, 2, "Student1");
            excel.SetData(1, 3, "Student2");
            excel.SetData(1, 4, "Student3");

            excel.SetData(2, 1, "Term1");
            excel.SetData(2, 2, "80");
            excel.SetData(2, 3, "65");
            excel.SetData(2, 4, "45");

            excel.SetData(3, 1, "Term2");
            excel.SetData(3, 2, "81");
            excel.SetData(3, 3, "61");
            excel.SetData(3, 4, "41");

            excel.SetData(4, 1, "Term3");
            excel.SetData(4, 2, "82");
            excel.SetData(4, 3, "62");
            excel.SetData(4, 4, "42");

            excel.SelectWorksheet(2);
            excel.SetChart("A1", "D4", Excel.XlChartType.xlColumnClustered);

            excel.SelectWorksheet(3); 
            excel.SetChart("A1", "D4", Excel.XlChartType.xlLine);
            excel.SaveAs(GetExcelFileName(textBoxNumber.Text));

            excel.Release(); 
        }

        public String GetExcelFileName(String append)
        {
            if (append.Equals("")) {
                append = "xxxx";
            }
            String curDate = DateTime.Now.ToString("yyyyMMdd");
            String ret = System.Environment.CurrentDirectory
                + "\\" + "CRH380D-" + append + "_" + curDate + ".xlsx";

            return ret;
        }

        private void labelTitle_Click(object sender, EventArgs e)
        {

        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "数据文件(*.txt)|*.txt|所有文件(*.*)|*.*";
            this.openFileDialog1.FileName = "电能列表2016-03-02.TXT";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string FileName = this.openFileDialog1.FileName;
                SetEnergyDataFromFile(FileName);
            }
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }
    }
}
