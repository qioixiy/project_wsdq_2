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
        public FormMain()
        {
            InitializeComponent();
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

            excel.SetChart("A1", "D4", Excel.XlChartType.xlLine);
            excel.SaveAs();
            excel.Release();
        }
    }
}
