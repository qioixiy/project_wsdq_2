﻿using System;
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
            //this.dateTimePicker1.MaxDate = new DateTime(;


            CheckForIllegalCrossThreadCalls = false;

            if ((Myutility.GetMajorVersionNumber() == "V1.1")
                || (Myutility.GetMajorVersionNumber() == "V1.3"))
            {
                labelNumber.Visible = true;
                textBoxNumber.Visible = true;
                label1.Visible = true;
                textBox_V_type.Visible = true;
            }
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

        public void setExportExcelStatus(string status, string detail = "")
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
            else if (status.Equals("processing"))
            {
                buttonExportExcel.Text = detail;
                return;
            }

            if (enable == true)
            {
                buttonExportExcel.Text = "导出为Excel";
                buttonExportExcel.Enabled = true;
            }
            else
            {
                buttonExportExcel.Enabled = false;
            }
        }
        private string GetExcelFileNameV1_2()
        {
            string ret;
            string carType = "";
            if (mEnergyData.carType[0] == 0x01)
            {
                carType = "CRH1A";
            }
            else if (mEnergyData.carType[0] == 0x02)
            {
                carType = "CRH1E";
            }
            else if (mEnergyData.carType[0] == 0x03)
            {
                carType = "CRH380D";
            }
            else
            {
                MessageBox.Show("未识别的车型");
                return null;
            }

            short num = System.BitConverter.ToInt16(mEnergyData.carNum, 0);
            string carNum = num.ToString();
            string pre = System.Environment.CurrentDirectory + "\\";
            ret = pre + carType + "-" + carNum + "_" + DateTime.Now.ToString("yyyyMMdd");
            return ret;
        }

        private void filterEnergyDataWithDateTime(int dateTime)
        {
            List<ExportExcel.EnergyData.EnergyDataRaw> tEnergyDataRawList = new List<ExportExcel.EnergyData.EnergyDataRaw>();

            for (int i = 0; i < mEnergyData.mEnergyDataRawList.Count; i++)
            {
                if (dateTime == (int)mEnergyData.mEnergyDataRawList[i].getDay())
                {
                    tEnergyDataRawList.Add(mEnergyData.mEnergyDataRawList[i]);
                }
            }

            mEnergyData.mEnergyDataRawList = tEnergyDataRawList;
        }
        private void buttonExportExcel_Click(object sender, EventArgs e)
        {
            bool filter = false;
            if (filter)
            {
                if (comboBox1.Text == "请选择日期")
                {
                    MessageBox.Show("请先选择日期");
                    return;
                }
                else
                {
                    filterEnergyDataWithDateTime(Int32.Parse(comboBox1.Text));
                }
            }

            if (null == mEnergyData)
            {
                MessageBox.Show("请先导入数据文件");
                return;
            }
            setExportExcelStatus("exporting");

            ExportExcelThread mExportExcelThread;
            //mExportExcelThread = new ExportExcelThread(this, mEnergyData, GetExcelFileName(textBoxNumber.Text));
            // V1.2
            string filename = null;
            if ((Myutility.GetMajorVersionNumber() == "V1.1")
                || (Myutility.GetMajorVersionNumber() == "V1.3"))
            {
                filename = GetExcelFileName(textBoxNumber.Text);
            }
            else
            {
                filename = GetExcelFileNameV1_2();
            }
            if (null == filename)
            {
                MessageBox.Show("无效文件名");
                setExportExcelStatus("export-fail");
                return;
            }
            mExportExcelThread = new ExportExcelThread(this, mEnergyData, filename);
            Thread th = new Thread(mExportExcelThread.ThreadMethod);

            th.Start();
        }

        public String GetExcelFileName(String append)
        {
            if (append.Equals(""))
            {
                append = "xxxx";
            }
            String curDate = DateTime.Now.ToString("yyyyMMdd");
            String ret = System.Environment.CurrentDirectory
                + "\\" + textBox_V_type.Text + "-" + append + "_" + curDate;

            return ret;
        }

        private void labelTitle_Click(object sender, EventArgs e)
        {
            ;
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();

            this.openFileDialog1.Filter = "数据文件(*.txt)|*.txt|所有文件(*.*)|*.*";
            this.openFileDialog1.FileName = "*.TXT";
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

                // 记录时间段
                bool[] dateTimesFlag = new bool[31];

                for (int j = 0, count = 0, i = this.dataGridViewEnergy.RowCount - 1; i >= 0; i--, count++)
                {
                    dateTimesFlag[1] = true;

                    string day = mEnergyData.mEnergyDataRawList[i].getDay().ToString();
                    string hour = mEnergyData.mEnergyDataRawList[i].getHour().ToString();
                    string minutes = mEnergyData.mEnergyDataRawList[i].getMinuts().ToString();
                    string second = mEnergyData.mEnergyDataRawList[i].getSecond().ToString();
                    string power1 = mEnergyData.mEnergyDataRawList[i].getPower1().ToString();
                    string power2 = mEnergyData.mEnergyDataRawList[i].getPower2().ToString();
                    string powerAll = mEnergyData.mEnergyDataRawList[i].getPowerAll().ToString();
                    string step_power1 = "0";
                    string step_power2 = "0";
                    string step_powerAll = "0";

                    if (count != dataGridViewEnergy.RowCount - 1)
                    {
                        step_power1 = (mEnergyData.mEnergyDataRawList[i].getPower1() - mEnergyData.mEnergyDataRawList[i - 1].getPower1()).ToString();
                        step_power2 = (mEnergyData.mEnergyDataRawList[i].getPower2() - mEnergyData.mEnergyDataRawList[i - 1].getPower2()).ToString();
                        step_powerAll = (mEnergyData.mEnergyDataRawList[i].getPowerAll() - mEnergyData.mEnergyDataRawList[i - 1].getPowerAll()).ToString();
                    }
                    dataGridViewEnergy.Rows[j].Cells[0].Value = day + "日" + hour + "时" + minutes + "分" + second + "秒";
                    dataGridViewEnergy.Rows[j].Cells[1].Value = power1 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[2].Value = power2 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[3].Value = powerAll + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[4].Value = step_power1 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[5].Value = step_power2 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[6].Value = step_powerAll + " kW.h";
                    j++;
                }

                for (int index = 0; index < dateTimesFlag.Length; index++)
                {
                    if (dateTimesFlag[index])
                    {
                        comboBox1.Items.Add(index + 1);
                    }
                }
            }
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void FormMain_Shown(object sender, EventArgs e)
        {
            dataGridViewEnergy.Font = new Font("Arial", 9);
        }

        private void dataGridViewEnergy_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                e.CellStyle.Font = new Font("微软雅黑", 9);
                return;
            }

            try
            {
                if (e.ColumnIndex == 0)//定位到第1列日期 
                {
                    e.CellStyle.Font = new Font("微软雅黑", 9);
                }
            }
            catch
            {

            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


    }
}
