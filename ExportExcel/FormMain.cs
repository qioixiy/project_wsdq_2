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
using System.Windows.Forms.DataVisualization.Charting;

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

            if ((Myutility.GetMajorVersionNumber() == "V1.1")
                || (Myutility.GetMajorVersionNumber() == "V1.3"))
            {
                labelNumber.Visible = true;
                textBoxNumber.Visible = true;
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
            else {
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

            mEnergyData.mEnergyDataRawListExportExcel = tEnergyDataRawList;
        }
        private void buttonExportExcel_Click(object sender, EventArgs e)
        {
            if (null == mEnergyData)
            {
                MessageBox.Show("请先导入数据文件");
                return;
            }

            if (comboBoxDays.SelectedIndex < 0)
            {
                MessageBox.Show("请先选择需要导出日期");
                return;
            }

            // 过滤出指定日期的数据
            filterEnergyDataWithDateTime(Int32.Parse(comboBoxDays.Text));

            setExportExcelStatus("exporting");

            ExportExcelThread mExportExcelThread;
            // V1.2
            string filename = GetExcelFileName(textBoxNumber.Text);
            
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
            if (append.Equals("")) {
                append = "xxxx";
            }
            String curDate = DateTime.Now.ToString("yyyyMMdd");
            String ret = System.Environment.CurrentDirectory
                + "\\" + "CRH-" + append + "_" + curDate;

            return ret;
        }

        private void labelTitle_Click(object sender, EventArgs e)
        {
            ;
        }

        private void updateDataGridView()
        {
            for (int j = 0, count = 0, i = this.dataGridViewEnergy.RowCount - 1; i >= 0; i--, count++)
            {
                string consumePower = "0";
                string revivePower = "0";
                string totalPower = "0";

                if (i != 0)
                {
                    consumePower = (mEnergyData.mEnergyDataRawList[i].consumeEnergy - mEnergyData.mEnergyDataRawList[i - 1].consumeEnergy).ToString();
                    revivePower = (mEnergyData.mEnergyDataRawList[i].reviveEgergy - mEnergyData.mEnergyDataRawList[i - 1].reviveEgergy).ToString();
                    totalPower = (mEnergyData.mEnergyDataRawList[i].totalEnergy - mEnergyData.mEnergyDataRawList[i - 1].totalEnergy).ToString();
                }

                dataGridViewEnergy.Rows[j].Cells[0].Value = mEnergyData.mEnergyDataRawList[i].GetDateTime();
                dataGridViewEnergy.Rows[j].Cells[1].Value = consumePower + " kW.h";
                dataGridViewEnergy.Rows[j].Cells[2].Value = revivePower + " kW.h";
                dataGridViewEnergy.Rows[j].Cells[3].Value = totalPower + " kW.h";
                j++;
            }
        }

        private void updateDaysComboBox()
        {
            // 获取有效的时间天数
            bool[] days = new bool[31];
            foreach (var item in mEnergyData.mEnergyDataRawList)
            {
                days[item.getDay()] = true;
            }

            for (int index = 0; index < 31; index++)
            {
                if (days[index] == true)
                {
                    comboBoxDays.Items.Add(index);
                }
            }
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            comboBoxDays.Items.Clear();

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

                // update DataGridView
                updateDataGridView();
                updateDaysComboBox();

                //初始默认显示全部数据
                dateTimePickerFrom.Value = mEnergyData.mEnergyDataRawList[0].EnergyDate;
                dateTimePickerTo.Value = mEnergyData.mEnergyDataRawList[mEnergyData.mEnergyDataRawList.Count - 1].EnergyDate;
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void buttonSaveImage_Click(object sender, EventArgs e)
        {
            
            string ImagePath = Directory.GetCurrentDirectory();
            string FileName = ImagePath + "\\" + "Test" + ".jpeg";
            chartPower.SaveImage(FileName, ChartImageFormat.Jpeg);

            MessageBox.Show(FileName, "保存成功");
        }

        private void comboBoxUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((mEnergyData == null) || (mEnergyData.mEnergyDataRawList.Count == 0))
            {
                MessageBox.Show("请先导入数据！");
                return;
            }
            chartPower.ChartAreas["ChartArea1"].CursorX.AutoScroll = true;
            chartPower.ChartAreas["ChartArea1"].AxisX.ScrollBar.Enabled = true;
            chartPower.ChartAreas["ChartArea1"].CursorX.IsUserEnabled = true;
            chartPower.ChartAreas["ChartArea1"].CursorX.IsUserSelectionEnabled = true;
            chartPower.ChartAreas["ChartArea1"].AxisX.Interval = 1;
            chartPower.ChartAreas["ChartArea1"].AxisX.ScaleView.Zoomable = true;
            chartPower.ChartAreas["ChartArea1"].AxisX.ScaleView.Position = 0;
            chartPower.ChartAreas["ChartArea1"].AxisX.ScaleView.Size = 1 * 10;

            chartPower.Series.Clear();
            Series SeriesConsumePower = new Series("ConsumePower");
            SeriesConsumePower.ChartType = SeriesChartType.Column;
            SeriesConsumePower.BorderWidth = 3;
            SeriesConsumePower.ShadowOffset = 2;

            Series seriesRevivePower = new Series("RevivePower");
            seriesRevivePower.ChartType = SeriesChartType.Column;
            seriesRevivePower.BorderWidth = 3;
            seriesRevivePower.ShadowOffset = 2;

            Series seriesTotalPower = new Series("TotalPower");
            seriesTotalPower.ChartType = SeriesChartType.Column;
            seriesTotalPower.BorderWidth = 3;
            seriesTotalPower.ShadowOffset = 2;
            if (comboBoxUnit.SelectedIndex == 0)
            {
                var EnergyByYear = mEnergyData.mEnergyDataRawList.GroupBy(energy => energy.EnergyDate.Year).ToList();

                int i = 0;
                foreach (var energy in EnergyByYear)
                {                    
                    SeriesConsumePower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy));

                    //X轴显示的名称
                    SeriesConsumePower.Points[i].AxisLabel = energy.Key.ToString() + "年";

                    //顶部显示的数字
                    SeriesConsumePower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy).ToString();
                    //鼠标放上去的提示内容
                    SeriesConsumePower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy).ToString();



                    seriesRevivePower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy));

                    seriesRevivePower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy).ToString();

                    seriesRevivePower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy).ToString();

                    seriesTotalPower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy));

                    seriesTotalPower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy).ToString();

                    seriesTotalPower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy).ToString();

                    i++;
                }
            }
            else if (comboBoxUnit.SelectedIndex == 1)
            {
                var EnergyByMonth = mEnergyData.mEnergyDataRawList.GroupBy(energy => { return new { energy.EnergyDate.Year, energy.EnergyDate.Month }; }).ToList();

                int i = 0;
            foreach (var energy in EnergyByMonth)
            {
                SeriesConsumePower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy));

                //X轴显示的名称
                SeriesConsumePower.Points[i].AxisLabel = energy.Key.Year.ToString() + "/" + energy.Key.Month.ToString();

                //顶部显示的数字
                SeriesConsumePower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy).ToString();
                //鼠标放上去的提示内容
                SeriesConsumePower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy).ToString();



                seriesRevivePower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy));

                seriesRevivePower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy).ToString();

                seriesRevivePower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy).ToString();

                seriesTotalPower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy));

                seriesTotalPower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy).ToString();

                seriesTotalPower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy).ToString();

                i++;
            } 
            }
            else if (comboBoxUnit.SelectedIndex == 2)
            {
                var EnergyByDay = mEnergyData.mEnergyDataRawList.GroupBy(energy => { return new { energy.EnergyDate.Year, energy.EnergyDate.Month, energy.EnergyDate.Day }; }).ToList();

                int i = 0;
                foreach (var energy in EnergyByDay)
                {
                    SeriesConsumePower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy));

                    //X轴显示的名称
                    SeriesConsumePower.Points[i].AxisLabel = energy.Key.Year.ToString() + "/" + energy.Key.Month.ToString() + "/" + energy.Key.Day.ToString();

                    //顶部显示的数字
                    SeriesConsumePower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy).ToString();
                    //鼠标放上去的提示内容
                    SeriesConsumePower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy).ToString();



                    seriesRevivePower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy));

                    seriesRevivePower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy).ToString();

                    seriesRevivePower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy).ToString();

                    seriesTotalPower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy));

                    seriesTotalPower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy).ToString();

                    seriesTotalPower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy).ToString();

                    i++;
                }
            }
            else if (comboBoxUnit.SelectedIndex == 3)
            {
                var EnergyByHour = mEnergyData.mEnergyDataRawList.GroupBy(energy => { return new { energy.EnergyDate.Year, energy.EnergyDate.Month, energy.EnergyDate.Day, energy.EnergyDate.Hour }; }).ToList();

                int i = 0;
                foreach (var energy in EnergyByHour)
                {
                    SeriesConsumePower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy));

                    //X轴显示的名称
                    SeriesConsumePower.Points[i].AxisLabel = energy.Key.Year.ToString() + "/" + energy.Key.Month.ToString() + "/" + energy.Key.Day.ToString() + " " + energy.Key.Hour.ToString();

                    //顶部显示的数字
                    SeriesConsumePower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy).ToString();
                    //鼠标放上去的提示内容
                    SeriesConsumePower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy).ToString();



                    seriesRevivePower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy));

                    seriesRevivePower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy).ToString();

                    seriesRevivePower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy).ToString();

                    seriesTotalPower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy));

                    seriesTotalPower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy).ToString();

                    seriesTotalPower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy).ToString();

                    i++;
                }
            }
            else if (comboBoxUnit.SelectedIndex == 4)
            {
                var EnergyByMinute = mEnergyData.mEnergyDataRawList.GroupBy(energy => { return new { energy.EnergyDate.Year, energy.EnergyDate.Month, energy.EnergyDate.Day, energy.EnergyDate.Hour, energy.EnergyDate.Minute }; }).ToList();

                int i = 0;
                foreach (var energy in EnergyByMinute)
                {
                    SeriesConsumePower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy));

                    //X轴显示的名称
                    SeriesConsumePower.Points[i].AxisLabel = energy.Key.Year.ToString() + "/" + energy.Key.Month.ToString() + "/" + energy.Key.Day.ToString() + " " + energy.Key.Hour.ToString() + ":" + energy.Key.Minute.ToString();

                    //顶部显示的数字
                    SeriesConsumePower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy).ToString();
                    //鼠标放上去的提示内容
                    SeriesConsumePower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.consumeEnergy).ToString();



                    seriesRevivePower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy));

                    seriesRevivePower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy).ToString();

                    seriesRevivePower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.reviveEgergy).ToString();

                    seriesTotalPower.Points.AddY(energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy));

                    seriesTotalPower.Points[i].Label = energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy).ToString();

                    seriesTotalPower.Points[i].ToolTip = energy.ToList().Sum(tempEnergy => tempEnergy.totalEnergy).ToString();

                    i++;
                }
            }
            chartPower.Series.Add(SeriesConsumePower);
            chartPower.Series.Add(seriesRevivePower);
            chartPower.Series.Add(seriesTotalPower);
        }

        private bool inDateTime(DateTime dateFrom, DateTime dateTo, DateTime cur)
        {
            bool ret = false;
            if (dateFrom <= cur && cur <= dateTo)
            {
                ret = true;
            }
            else
            {
                ret = false;
            }
            return ret;
        }
        private string getTextBoxReport()
        {

            int consumePower = 0;
            int revivePower = 0;

            for (int j = 0, count = 0, i = this.dataGridViewEnergy.RowCount - 1; i >= 0; i--, j++, count++)
            {
                if (i != 0)
                {
                    EnergyData.EnergyDataRaw tEnergyDataRaw = mEnergyData.mEnergyDataRawList[i];
                    DateTime curDateTime = new DateTime(
                        tEnergyDataRaw.getYear(),
                        tEnergyDataRaw.getMouth(),
                        tEnergyDataRaw.getDay(),
                        tEnergyDataRaw.getHour(),
                        tEnergyDataRaw.getMinute(),
                        0
                    );

                    DateTime dateFrom = DateTime.Parse(dateTimePickerFrom.Text);
                    DateTime dateTo = DateTime.Parse(dateTimePickerTo.Text);

                    if (!inDateTime(dateFrom, dateTo, curDateTime))
                    {
                        continue;
                    }
                    consumePower += (mEnergyData.mEnergyDataRawList[i].consumeEnergy - mEnergyData.mEnergyDataRawList[i - 1].consumeEnergy);
                    revivePower += (mEnergyData.mEnergyDataRawList[i].reviveEgergy - mEnergyData.mEnergyDataRawList[i - 1].reviveEgergy);
                }
            }

            return "消耗电能:" + consumePower + "kwh,再生电能:" + revivePower + "kwh";
        }

        private void dateTimePickerFrom_ValueChanged(object sender, EventArgs e)
        {
            textBoxReport.Text = getTextBoxReport();
        }
    }
}
