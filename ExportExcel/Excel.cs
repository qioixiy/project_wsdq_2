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
            GetWorksheet(2).Name = "单日电能柱状图";
            GetWorksheet(3).Name = "总电能曲线图";
            GetWorksheet(4).Name = "...";

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

            //myChart.Chart.ChartWizard(chartRange, type, Type.Missing, XlRowCol.xlColumns, 1, 1, true, "电能数据", "日期", "KW.h", Type.Missing);
        }

        public static int GenExcel(FormMain form, EnergyData mEnergyData, string filename)
        {
            CBExcel excel = new CBExcel();
            excel.Create();

            excel.SelectWorksheet(1);
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["A", Type.Missing]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
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
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, 7]).Interior.ColorIndex = 16;

            excel.SetData(1, 1, "日期");
            excel.SetData(1, 2, "正向电能(kW.h)");
            excel.SetData(1, 3, "反向电能(kW.h)");
            excel.SetData(1, 4, "总电能(kW.h)");
            excel.SetData(1, 5, "日耗电能(kW.h)");
            excel.SetData(1, 6, "日馈电能(kW.h)");
            excel.SetData(1, 7, "单日总耗电能(kW.h)");

            if (mEnergyData == null || mEnergyData.mEnergyDataRawList.Count == 0)
            {
                form.setExportExcelStatus("unknown-data");
                return -1;
            }

            int row = 2;
            for (int i = 0; i < mEnergyData.mEnergyDataRawList.Count; i++)
            {
                row = i + 2;

                string v0_0, v0_1, v0_2, v1, v2, v3, v4, v5, v6;

                v0_0 = BitConverter.ToInt16(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].year), 0).ToString();
                v0_1 = Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].mouth), System.Globalization.NumberStyles.HexNumber).ToString();
                v0_2 = Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].day), System.Globalization.NumberStyles.HexNumber).ToString();
                v1 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0).ToString();
                v2 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0).ToString();
                v3 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0).ToString();

                if (i == 0)
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
                object s1, s2, s3, s4, s5, s6, s7;
                s1 = BitConverter.ToInt16(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].year), 0) + "年"
                     + Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].mouth), System.Globalization.NumberStyles.HexNumber) + "月"
                     + Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].day), System.Globalization.NumberStyles.HexNumber) + "日";
                s2 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0);
                s3 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0);
                s4 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0);
                if (i == 0)
                {
                    s5 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0);
                    s6 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0);
                    s7 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0);
                }
                else
                {
                    s5 = (BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0)
                       - BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power1), 0));
                    s6 = (BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0)
                       - BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power2), 0));
                    s7 = (BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0)
                       - BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].powerAll), 0));
                }
                excel.SetData(row, 1, v0_0 + "年" + v0_1 + "月" + v0_2 + "日");
                excel.SetData(row, 2, v1);
                excel.SetData(row, 3, v2);
                excel.SetData(row, 4, v3);
                excel.SetData(row, 5, v4);
                excel.SetData(row, 6, v5);
                excel.SetData(row, 7, v6);
            }

            // sheet 4
            excel.SelectWorksheet(4);
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["A", Type.Missing]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["A", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["B", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["C", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["D", Type.Missing]).ColumnWidth = 20;
            //((Excel.Range)excel.CurXlsWorkSheet.Columns["E", Type.Missing]).ColumnWidth = 20;
            //((Excel.Range)excel.CurXlsWorkSheet.Columns["F", Type.Missing]).ColumnWidth = 20;
            //((Excel.Range)excel.CurXlsWorkSheet.Columns["G", Type.Missing]).ColumnWidth = 20;
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, 4]).Font.Name = "微软雅黑";
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, 4]).Font.Bold = true;
            //excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, 4]).VerticalAlignment = true;
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, 4]).Interior.ColorIndex = 16;

            excel.SetData(1, 1, "日期");
            excel.SetData(1, 2, "日耗电能(kW.h)");
            excel.SetData(1, 3, "日馈电能(kW.h)");
            excel.SetData(1, 4, "单日总耗电能(kW.h)");

            if (mEnergyData == null || mEnergyData.mEnergyDataRawList.Count == 0)
            {
                form.setExportExcelStatus("unknown-data");
                return -1;
            }

            row = 2;
            for (int i = 0; i < mEnergyData.mEnergyDataRawList.Count; i++)
            {
                row = i + 2;

                string v0_0, v0_1, v0_2, v1, v2, v3, v4, v5, v6;

                v0_0 = BitConverter.ToInt16(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].year), 0).ToString();
                v0_1 = Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].mouth), System.Globalization.NumberStyles.HexNumber).ToString();
                v0_2 = Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].day), System.Globalization.NumberStyles.HexNumber).ToString();
                v1 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0).ToString();
                v2 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0).ToString();
                v3 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0).ToString();

                if (i == 0)
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
                object s1, s2, s3, s4, s5, s6, s7;
                s1 = BitConverter.ToInt16(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].year), 0) + "年"
                     + Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].mouth), System.Globalization.NumberStyles.HexNumber) + "月"
                     + Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].day), System.Globalization.NumberStyles.HexNumber) + "日";
                s2 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0);
                s3 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0);
                s4 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0);
                if (i == 0)
                {
                    s5 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0);
                    s6 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0);
                    s7 = BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0);
                }
                else
                {
                    s5 = (BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0)
                       - BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power1), 0));
                    s6 = (BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0)
                       - BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power2), 0));
                    s7 = (BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0)
                       - BitConverter.ToInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].powerAll), 0));
                }
                excel.SetData(row, 1, v0_0 + "年" + v0_1 + "月" + v0_2 + "日");
                excel.SetData(row, 2, v4);
                excel.SetData(row, 3, v5);
                excel.SetData(row, 4, v6);
            }
            excel.CurXlsWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;

            // sheet 2
            excel.SelectWorksheet(2);
            int chart_width = row * 10;
            excel.SetChart(4, Excel.XlChartType.xlColumnClustered, "A1", "D" + row, chart_width);

            excel.SelectWorksheet(3);
            excel.SetChart(1, Excel.XlChartType.xlLine, "A1", "D" + row, chart_width);

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
