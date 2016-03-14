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
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["A", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["B", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["C", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["D", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["E", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["F", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["G", Type.Missing]).ColumnWidth = 20;
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, 7]).Font.Name = "微软雅黑";
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, 7]).Font.Bold = true;
            //excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, 7]).VerticalAlignment = true;
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1,7]).Interior.ColorIndex = 16;

            excel.SetData(1, 1, "日期");
            excel.SetData(1, 2, "正向电能");
            excel.SetData(1, 3, "反向电能");
            excel.SetData(1, 4, "总电能");
            excel.SetData(1, 5, "日耗正向电能");
            excel.SetData(1, 6, "日馈反向电能");
            excel.SetData(1, 7, "单耗总电能");

            if (mEnergyData == null || mEnergyData.mEnergyDataRawList.Count == 0)
            {
                MessageBox.Show("请先导入正确的数据文件！");
                return;
            }

            int row = 2;
            for (int i = 0; i < mEnergyData.mEnergyDataRawList.Count; i++)
            {
                row = i + 2;
                object obj1 = BitConverter.ToUInt16(ToHostEndian(mEnergyData.mEnergyDataRawList[i].year), 0) + "年"
                                                            + BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].mouth) + "月"
                                                            + BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].day) + "日";
                object obj2 = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0) + "kW.h";
                object obj3 = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0) + "kW.h";
                object obj4 = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0) + "kW.h";

                object obj5, obj6, obj7;
                if (i == 0)
                {
                    obj5 = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0) + "kW.h";
                    obj6 = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0) + "kW.h";
                    obj7 = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0) + "kW.h";
                }
                else
                {
                    obj5 = (BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0)
                       - BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power1), 0)) + "kW.h";
                    obj6 = (BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0)
                       - BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power2), 0)) + "kW.h";
                    obj7 = (BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0)
                       - BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].powerAll), 0)) + "kW.h";
                }

                excel.SetData(row, 1, (string)obj1);
                excel.SetData(row, 2, (string)obj2);
                excel.SetData(row, 3, (string)obj3);
                excel.SetData(row, 4, (string)obj4);
                excel.SetData(row, 5, (string)obj5);
                excel.SetData(row, 6, (string)obj6);
                excel.SetData(row, 7, (string)obj7);
            }

            excel.SelectWorksheet(2);
            object charts= excel.CurXlsWorkSheet.ChartObjects(Type.Missing);
            // excel.SetChart("A2", "F5", Excel.XlChartType.xlColumnClustered);

            //excel.SelectWorksheet(3); 
            //excel.SetChart("A2", "F5", Excel.XlChartType.xlLine);
            //excel.SaveAs(GetExcelFileName(textBoxNumber.Text));

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
            ;
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "数据文件(*.txt)|*.txt|所有文件(*.*)|*.*";
            this.openFileDialog1.FileName = "电能列表2016-03-02.TXT";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string FileName = this.openFileDialog1.FileName;
                SetEnergyDataFromFile(FileName);
                this.dataGridViewEnergy.RowCount = mEnergyData.mEnergyDataRawList.Count;
                if (this.dataGridViewEnergy.RowCount == 0)
                {
                    MessageBox.Show("数据文件内容为空");
                    return;
                }
                for (int i = 0; i < this.dataGridViewEnergy.RowCount; i++)
                {
                    dataGridViewEnergy.Rows[i].Cells[0].Value = BitConverter.ToUInt16(ToHostEndian(mEnergyData.mEnergyDataRawList[i].year), 0) + "年" 
                                                            + BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].mouth) + "月"
                                                            + BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].day) + "日";
                    dataGridViewEnergy.Rows[i].Cells[1].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0) + "kW.h";
                    dataGridViewEnergy.Rows[i].Cells[2].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0) + "kW.h";
                    dataGridViewEnergy.Rows[i].Cells[3].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0) + "kW.h";
                    if (i == 0)
                    {
                        dataGridViewEnergy.Rows[i].Cells[4].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0) + "kW.h";
                        dataGridViewEnergy.Rows[i].Cells[5].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0) + "kW.h";
                        dataGridViewEnergy.Rows[i].Cells[6].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0) + "kW.h";
                    }
                    else
                    {
                        dataGridViewEnergy.Rows[i].Cells[4].Value = (BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0)
                           - BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power1), 0)) + "kW.h";
                        dataGridViewEnergy.Rows[i].Cells[5].Value = (BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0)
                           - BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power2), 0)) + "kW.h";
                        dataGridViewEnergy.Rows[i].Cells[6].Value = (BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0)
                           - BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].powerAll), 0)) + "kW.h";
                    }
                    
                    //dataGridViewEnergy.Rows[i].Cells[2].Value = "test2";
                    //dataGridViewEnergy.Rows[i].Cells[3].Value = mEnergyData.mEnergyDataRawList[i].power1;
                    //dataGridViewEnergy.Rows[i].Cells[4].Value = mEnergyData.mEnergyDataRawList[i].power2;
                    //dataGridViewEnergy.Rows[i].Cells[0].Value = mEnergyData.mEnergyDataRawList[i].powerAll;
                   //string Str =  " mouth:" + BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].mouth)
                   //             + " day:" + BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].day);
                }
            }
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        byte[] ToHostEndian(byte[] src)
        {
            byte[] dest = new byte[src.Length];
            for (int i = src.Length - 1, j = 0; i >= 0; i--, j++)
            {
                dest[j] = src[i];
            }

            return dest;
        }
    }
}
