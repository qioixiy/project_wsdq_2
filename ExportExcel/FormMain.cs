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
using System.Threading; 

namespace ExportExcel
{
    public partial class FormMain : Form
    {
        public EnergyData mEnergyData;
        
        public FormMain()
        {
            InitializeComponent();
            this.label3.Text = ExportExcel.Properties.Resources.Version;
            CheckForIllegalCrossThreadCalls = false;
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
        
        public void setExportExcelStatus(string status)
        {
            bool enable = true;
            if (status.Equals("unknown-data"))
            {
                MessageBox.Show("请先导入正确的数据文件！");
                enable = true;
            }
            else if (status.Equals("exporting"))
            {
                buttonExportExcel.Text = "导出...";
                enable = false;
            }
            else if (status.Equals("export-success"))
            {
                MessageBox.Show("导出成功！");
                buttonExportExcel.Text = "导出为Excel";
                enable = true;
            }
            else if (status.Equals("export-fail"))
            {
                buttonExportExcel.Text = "导出为Excel";
                enable = true;
            }

            if (enable == true)
            {
                buttonExportExcel.Text = "导出为Excel";
                buttonExportExcel.Enabled = true;
            }
            else {
                buttonExportExcel.Enabled = false;
            }
        }

        private void buttonExportExcel_Click(object sender, EventArgs e)
        {
            if (null == mEnergyData) {
                MessageBox.Show("请先导入数据文件");
                return;
            }
            setExportExcelStatus("exporting");
           
            ExportExcelThread mExportExcelThread = new ExportExcelThread(this, mEnergyData, GetExcelFileName(textBoxNumber.Text));
            Thread th = new Thread(mExportExcelThread.ThreadMethod);

            th.Start();
        }

        public String GetExcelFileName(String append)
        {
            if (append.Equals("")) {
                append = "xxxx";
            }
            String curDate = DateTime.Now.ToString("yyyyMMdd");
            String ret = System.Environment.CurrentDirectory
                + "\\" + textBox_V_type.Text + "-" + append + "_" + curDate + ".xls";

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
                for (int j = 0, count = 0, i = this.dataGridViewEnergy.RowCount - 1; i >= 0; i--, count++)
                {
                    string v0_0, v0_1, v0_2, v1, v2, v3, v4, v5, v6;

                    v0_0 = BitConverter.ToInt16(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].year), 0).ToString();
                    v0_1 = Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].mouth), System.Globalization.NumberStyles.HexNumber).ToString();
                    v0_2 = Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].day), System.Globalization.NumberStyles.HexNumber).ToString();
                    v1 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0).ToString();
                    v2 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0).ToString();
                    v3 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0).ToString();
                    if (count == this.dataGridViewEnergy.RowCount - 1)
                    {
                        v4 = "0";
                        v5 = "0";
                        v6 = "0";
                    }
                    else
                    {
                        v4 = (BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0)
                            - BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power1), 0)).ToString();
                        v5 = (BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0)
                            - BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power2), 0)).ToString();
                        v6 = (BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0)
                            - BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].powerAll), 0)).ToString();
                    }
                    
                    dataGridViewEnergy.Rows[j].Cells[0].Value = v0_0 + "年" +　v0_1 + "月" + v0_2 + "日";
                    dataGridViewEnergy.Rows[j].Cells[1].Value = v1 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[2].Value = v2 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[3].Value = v3 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[4].Value = v4 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[5].Value = v5 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[6].Value = v6 + " kW.h";

                    j++;
                 }
            }
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void FormMain_Shown(object sender, EventArgs e)
        {
            dataGridViewEnergy.Font = new Font("Arial",9);
        }

        private void dataGridViewEnergy_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                e.CellStyle.Font = new Font("微软雅黑",9);  
                return;
            }

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
