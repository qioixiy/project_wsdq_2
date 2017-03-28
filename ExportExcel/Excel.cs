using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExportExcel
{
    public class CBExcel
    {
        Excel.Application xlsApp;
        Excel.Workbook xlsWorkBook;
        public Excel.Worksheet CurXlsWorkSheet;
        object misValue = System.Reflection.Missing.Value;

        public CBExcel()
        {
            ;
        }

        ~CBExcel()
        {
            ;
        }

        private class RowColumData
        {
            public RowColumData(int row, int col, string data)
            {
                this.row = row;
                this.col = col;
                this.data = data;
            }
            public int row;
            public int col;
            public string data;
        }

        public void SetData(int i, int j, string data)
        {
            CurXlsWorkSheet.Cells[i, j] = data;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public void Create()
        {
            xlsApp = new Excel.ApplicationClass();

            // default sheet1
            xlsWorkBook = xlsApp.Workbooks.Add(misValue);
            // add 2s sheet
            xlsApp.Worksheets.Add(misValue);
            xlsApp.Worksheets.Add(misValue);
            // add function sheet
            xlsApp.Worksheets.Add(misValue);
            GetWorksheet(1).Name = "电能列表";

            xlsWorkBook.EnableAutoRecover = false;
        }

        public bool SaveAs(String filename)
        {
            xlsApp.DisplayAlerts = false;
            xlsApp.AlertBeforeOverwriting = false;
            if (File.Exists(filename))
            {
                try
                {
                    File.Delete(filename);
                }
                catch (IOException)
                {
                    MessageBox.Show(filename + "已经打开");
                    return false;
                }
            }

            xlsApp.ActiveWorkbook.SaveCopyAs(filename);
            xlsApp.Quit();
            xlsApp = null;
            xlsWorkBook = null;
            CurXlsWorkSheet = null;

            return true;
        }

        public Excel.Worksheet GetWorksheet(int index)
        {
            return (Excel.Worksheet)xlsWorkBook.Worksheets.get_Item(index);
        }

        public void SelectWorksheet(int index)
        {
            CurXlsWorkSheet = (Excel.Worksheet)xlsWorkBook.Worksheets.get_Item(index);
        }
        public void KillProcess()
        {
            try
            {
                foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName("Excel"))
                {
                    if (!p.CloseMainWindow())
                    {
                        p.Kill();
                    }
                }
                GC.Collect();
            }
            catch (Exception vErr)
            {
                throw new Exception("", vErr);
            }
        }
        public void Release()
        {
            releaseObject(CurXlsWorkSheet);
            releaseObject(xlsWorkBook);
            releaseObject(xlsApp);
            KillProcess();
        }

        public void SetChart(int sheetSrc, Excel.XlChartType type, string start, string end, int width)
        {
            Excel.ChartObjects xlCharts = (Excel.ChartObjects)CurXlsWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(0, 0, width, 400);
            Excel.Chart chartPage = myChart.Chart;

            Excel.Range chartRange = GetWorksheet(sheetSrc).get_Range(start, end);
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = type;

            try
            {
                if (Myutility.GetMajorVersionNumber() == "V1.3")
                {
                    myChart.Chart.ChartWizard(
                        chartRange, type, Type.Missing, XlRowCol.xlColumns, 1, 1, true, "电能数据", "时间", "KW.h", Type.Missing);
                }
                else
                {
                    myChart.Chart.ChartWizard(
                        chartRange, type, Type.Missing, XlRowCol.xlColumns, 1, 1, true, "电能数据", "日期", "KW.h", Type.Missing);
                }
            }
            catch
            {
                Console.WriteLine("ChartWizard error");
            }
        }

        private static string datetime_str = "时间";
        private static string power1_str = "正向电能(kW.h)";
        private static string power2_str = "反向电能(kW.h)";
        private static string powerall_str = "总电能(kW.h)";
        private static string step_power1_str = "阶段耗电能(kW.h)";
        private static string step_power2_str = "阶段馈电能(kW.h)";
        private static string step_powerall_str = "阶段总耗电能(kW.h)";
        // 1.33增加的项
        private static string step_v_str = "电压(kV)";
        private static string step_i_str = "电流(A)";
        private static string step_powerFactor_str = "功率因素";
        private static string step_powerRealTime_str = "实时功率(kW)";
        private static string step_v3rd_str = "电压3次谐波含有率(%)";
        private static string step_v5rd_str = "电压5次谐波含有率(%)";
        private static string step_v7rd_str = "电压7次谐波含有率(%)";
        private static string step_v9rd_str = "电压9次谐波含有率(%)";
        private static string step_i3rd_str = "电流3次谐波含有率(%)";
        private static string step_i5rd_str = "电流5次谐波含有率(%)";
        private static string step_i7rd_str = "电流7次谐波含有率(%)";
        private static string step_i9rd_str = "电流9次谐波含有率(%)";
        private static string step_Rosebowcar_str = "升弓车厢";

        public static int GenExcel(FormMain form, EnergyData mEnergyData, string filename)
        {
            CBExcel excel = new CBExcel();

            excel.Create();

            // append file ext
            string ext = ".xls";
            if (Convert.ToDouble(excel.xlsApp.Version) >= 12.0)//office 2007
            {
                ext = ".xlsx";
            }
            filename += ext;

            // worksheet1 电能列表
            excel.SelectWorksheet(1);
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["A", Type.Missing]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["A", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["B", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["C", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["D", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["E", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["F", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["G", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["H", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["I", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["J", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["K", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["L", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["M", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["N", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["O", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["P", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["Q", Type.Missing]).ColumnWidth = 20;
            int sum_title = 17;
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, sum_title]).Font.Name = "微软雅黑";
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, sum_title]).Font.Bold = true;
            //excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, sum_title]).VerticalAlignment = true;
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, sum_title]).Interior.ColorIndex = 16;

            excel.SetData(1, 1, datetime_str);
            excel.SetData(1, 2, power1_str);
            excel.SetData(1, 3, power2_str);
            excel.SetData(1, 4, powerall_str);
            //excel.SetData(1, 5, step_power1_str);
            //excel.SetData(1, 6, step_power2_str);
            //excel.SetData(1, 7, step_powerall_str);

            excel.SetData(1, 5, step_v_str);
            excel.SetData(1, 6, step_i_str);
            excel.SetData(1, 7, step_powerFactor_str);
            excel.SetData(1, 8, step_powerRealTime_str);
            excel.SetData(1, 9, step_v3rd_str);
            excel.SetData(1, 10, step_v5rd_str);
            excel.SetData(1, 11, step_v7rd_str);
            excel.SetData(1, 12, step_v9rd_str);
            excel.SetData(1, 13, step_i3rd_str);
            excel.SetData(1, 14, step_i5rd_str);
            excel.SetData(1, 15, step_i7rd_str);
            excel.SetData(1, 16, step_i9rd_str);
            excel.SetData(1, 17, step_Rosebowcar_str);

            if (mEnergyData == null || mEnergyData.mEnergyDataRawList.Count == 0)
            {
                form.setExportExcelStatus("unknown-data");
                return -1;
            }

            List<RowColumData> tRowColumDataList = new List<RowColumData>();
            int row = 2;
            for (int i = 0; i < mEnergyData.mEnergyDataRawList.Count; i++)
            {
                row = i + 2;

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

                if (i != 0)
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
                
                excel.SetData(row, 1, day + "日" + hour + "时" + minutes + "分" + second + "秒");
                excel.SetData(row, 2, power1);
                excel.SetData(row, 3, power2);
                excel.SetData(row, 4, powerAll);
                excel.SetData(row, 5, step_power1);
                excel.SetData(row, 6, step_power2);
                excel.SetData(row, 7, step_powerAll);

                excel.SetData(row, 5, step_v);
                excel.SetData(row, 6, step_i);
                excel.SetData(row, 7, step_powerFactor);
                excel.SetData(row, 8, step_powerRealTime);
                excel.SetData(row, 9, step_v3rd);
                excel.SetData(row, 10, step_v5rd);
                excel.SetData(row, 11, step_v7rd);
                excel.SetData(row, 12, step_v9rd);
                excel.SetData(row, 13, step_i3rd);
                excel.SetData(row, 14, step_i5rd);
                excel.SetData(row, 15, step_i7rd);
                excel.SetData(row, 16, step_i9rd);
                excel.SetData(row, 17, step_Rosebowcar);

                // 及时更新状态
                form.setExportExcelStatus("processing", "S1:" + i + "/" + mEnergyData.mEnergyDataRawList.Count);
            }
            if (excel.SaveAs(filename))
            {
                form.setExportExcelStatus("export-success");
            }
            else
            {
                form.setExportExcelStatus("export-fail");
            }

            excel.Release();

            return 0;
        }
    }
}
