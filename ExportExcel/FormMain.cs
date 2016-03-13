using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 

namespace ExportExcel
{
    public partial class FormMain : Form
    {
        struct PowerStruct{//电能列表结构
                            public char year;//年
                            public char mouth;//月
                            public char day;//日
                            public char power1;//正向电能
                            public char power2;//负向电能
                            public char powerAll;//总电能
                          };
        EnergyData mEnergyData;
        public FormMain()
        {
            InitializeComponent();
            mEnergyData = new EnergyData();
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
            excel.SaveAs(GetExcelFileName());
            excel.Release();
        }

        public String GetExcelFileName()
        {
            return System.Environment.CurrentDirectory + "\\export.xlsx";
        }

<<<<<<< HEAD
        private void buttonOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ElectircData = new OpenFileDialog();
            ElectircData.Filter = "文本文件|*.txt";
            if (ElectircData.ShowDialog() == DialogResult.OK)
            {
                //
                Stream st = new FileStream(ElectircData.FileName, FileMode.Open);
                BinaryReader br = new BinaryReader(st);

                PowerStruct pr = new PowerStruct();//一个未被具体指明的  

                //pr.year = br.;
                //pr.mouth = br.ReadChar;
                //pr.day = br.ReadChar;
                //pr.power1 = br.ReadChar;
                //pr.power2 = br.ReadChar;
                //pr.powerAll = br.ReadChar;
                br.Close();
                st.Close();  
            }
=======
        private void labelTitle_Click(object sender, EventArgs e)
        {

>>>>>>> ba72b6b0bd14c58cd408fac3d3f50f53c84b5248
        }
    }
}
