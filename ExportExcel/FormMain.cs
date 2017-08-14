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
            bool filter = true;
            if (filter)
            {
                if (comboBoxSelectDateTime.Text == "请选择日期")
                {
                    MessageBox.Show("请先选择日期");
                    return;
                }
                else
                {
                    filterEnergyDataWithDateTime(Int32.Parse(comboBoxSelectDateTime.Text));
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

        private String GetExportTimePoint()
        {
            String ret = "";
            Int32 chooseDay = Int32.Parse(comboBoxSelectDateTime.Text);

            String year = DateTime.Now.ToString("yyyy");
            String mouth = DateTime.Now.ToString("MM");
            String day = DateTime.Now.ToString("dd");

            String exportYear = year;
            String exportMouth = mouth;
            String exportDay = chooseDay.ToString().PadLeft(2, '0'); ;

            Int32 curDay = Int32.Parse(day);
            
            // 1.33修订版
            // 根据选择的日期来生产EXCEL的后缀日期，以前是根据软件在电脑的系统时间生产的日期，现在需要做判断，
            // 如果选择的日小于等于电脑系统的日，那么月份也就是电脑的系统的月份，
            // 如果选择的日大于电脑的日，那么月份肯定是电脑现在月份基础上的上一个月份，
            // 如果是1月份，上一个月份是12月
            if (chooseDay <= curDay)
            {
                exportMouth = mouth;
            }
            else
            {
                Int32 mouthInt32 = Int32.Parse(mouth);
                Int32 exportMouthInt32 = mouthInt32 - 1;
                if (mouthInt32 <= 1) {
                    exportMouthInt32 = 12;
                    exportYear = (Int32.Parse(year) - 1).ToString();
                }
                exportMouth = exportMouthInt32.ToString().PadLeft(2, '0');
            }

            ret = exportYear + exportMouth + exportDay;

            return ret;
        }

        public String GetExcelFileName(String append)
        {
            if (append.Equals(""))
            {
                append = "xxxx";
            }
            String curDate = DateTime.Now.ToString("yyyyMMdd");
            
            String exportDate = GetExportTimePoint();

            String ret = System.Environment.CurrentDirectory
                + "\\" + textBox_V_type.Text + "-" + append + "_" + exportDate;

            return ret;
        }

        private void labelTitle_Click(object sender, EventArgs e)
        {
            ;
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            comboBoxSelectDateTime.Items.Clear();

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
                    dateTimesFlag[mEnergyData.mEnergyDataRawList[i].getDay()-1] = true;

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
                    string step_v = mEnergyData.mEnergyDataRawList[i].getV().ToString();
                    string step_i = mEnergyData.mEnergyDataRawList[i].getI().ToString();
                    string step_powerFactor = mEnergyData.mEnergyDataRawList[i].getPowerFactor().ToString();
                    string step_powerRealTime = mEnergyData.mEnergyDataRawList[i].getPowerRealTime().ToString();
                    string step_v3rd = mEnergyData.mEnergyDataRawList[i].getV3rd().ToString();
                    string step_v5rd = mEnergyData.mEnergyDataRawList[i].getV5rd().ToString();
                    string step_v7rd = mEnergyData.mEnergyDataRawList[i].getV7rd().ToString();
                    string step_v9rd = mEnergyData.mEnergyDataRawList[i].getV9rd().ToString();
                    string step_i3rd = mEnergyData.mEnergyDataRawList[i].getI3rd().ToString();
                    string step_i5rd = mEnergyData.mEnergyDataRawList[i].getI5rd().ToString();
                    string step_i7rd = mEnergyData.mEnergyDataRawList[i].getI7rd().ToString();
                    string step_i9rd = mEnergyData.mEnergyDataRawList[i].getI9rd().ToString();
                    string step_Rosebowcar = "0";
                    switch (mEnergyData.mEnergyDataRawList[i].getRosebowcar())
                    {
                        default:
                        case 0: break;
                        case 1:
                            step_Rosebowcar = "3";
                            break;
                        case 2:
                            step_Rosebowcar = "6";
                            break;
                    }

                    int index = 0;
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = day + "日" + hour + "时" + minutes + "分" + second + "秒";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = power1 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = power2 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = powerAll + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_power1 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_power2 + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_powerAll + " kW.h";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_v + "kV";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_i + "A";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_powerFactor;
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_powerRealTime + "kW";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_v3rd + "%";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_v5rd + "%";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_v7rd + "%";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_v9rd + "%";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_i3rd + "%";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_i5rd + "%";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_i7rd + "%";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_i9rd + "%";
                    dataGridViewEnergy.Rows[j].Cells[index++].Value = step_Rosebowcar;

                    j++;
                }

                for (int index = 0; index < dateTimesFlag.Length; index++)
                {
                    if (dateTimesFlag[index])
                    {
                        comboBoxSelectDateTime.Items.Add(index + 1);
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
