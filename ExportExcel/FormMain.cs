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
        public EnergyData mEnergyData;
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
            excel.GenExcel(mEnergyData, GetExcelFileName(textBoxNumber.Text));
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
                                                            + Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].mouth),System.Globalization.NumberStyles.HexNumber) + "月"
                                                            + Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].day),System.Globalization.NumberStyles.HexNumber) + "日";
                    dataGridViewEnergy.Rows[i].Cells[1].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0) + " kW.h";
                    dataGridViewEnergy.Rows[i].Cells[2].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0) + " kW.h";
                    dataGridViewEnergy.Rows[i].Cells[3].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0) + " kW.h";
                    if (i == 0)
                    {
                        dataGridViewEnergy.Rows[i].Cells[4].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0) + " kW.h";
                        dataGridViewEnergy.Rows[i].Cells[5].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0) + " kW.h";
                        dataGridViewEnergy.Rows[i].Cells[6].Value = BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0) + " kW.h";
                    }
                    else
                    {
                        dataGridViewEnergy.Rows[i].Cells[4].Value = (BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0)
                           - BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power1), 0)) + " kW.h";
                        dataGridViewEnergy.Rows[i].Cells[5].Value = (BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0)
                           - BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power2), 0)) + " kW.h";
                        dataGridViewEnergy.Rows[i].Cells[6].Value = (BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0)
                           - BitConverter.ToUInt32(ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].powerAll), 0)) + " kW.h";
                    }
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

        private void FormMain_Shown(object sender, EventArgs e)
        {
            dataGridViewEnergy.Font = new Font("Arial",9);
        }

        private void dataGridViewEnergy_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            DataGridViewRow dgr = dataGridViewEnergy.Rows[e.RowIndex];
            try
            {
                if (e.ColumnIndex == 0)//定位到第1列日期 
                {
                        e.CellStyle.Font = new Font("微软雅黑",9);  
                }
            }
            catch
            {

            }  
        }
    }
}
