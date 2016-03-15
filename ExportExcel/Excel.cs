using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

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
            GetWorksheet(1).Name = "电能";
            GetWorksheet(2).Name = "柱状图";
            GetWorksheet(3).Name = "曲线图";
        }

        public void SaveAs(String filename)
        {
            xlsApp.DisplayAlerts = false;
            xlsApp.AlertBeforeOverwriting = false;
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }

            xlsApp.ActiveWorkbook.SaveCopyAs(filename);
            xlsApp.Quit();
            xlsApp = null;
            xlsWorkBook = null;
            CurXlsWorkSheet = null;
        }

        public Excel.Worksheet GetWorksheet(int index)
        {
            return (Excel.Worksheet)xlsWorkBook.Worksheets.get_Item(index);
        }
        
        public void SelectWorksheet(int index)
        {
            CurXlsWorkSheet = (Excel.Worksheet)xlsWorkBook.Worksheets.get_Item(index);
        }
        
        public void Release()
        {
            releaseObject(CurXlsWorkSheet);
            releaseObject(xlsWorkBook);
            releaseObject(xlsApp);
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
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1,7]).Interior.ColorIndex = 16;

            excel.SetData(1, 1, "日期");
            excel.SetData(1, 2, "正向电能(kW.h)");
            excel.SetData(1, 3, "反向电能(kW.h)");
            excel.SetData(1, 4, "总电能(kW.h)");
            excel.SetData(1, 5, "日耗正向电能(kW.h)");
            excel.SetData(1, 6, "日馈反向电能(kW.h)");
            excel.SetData(1, 7, "单耗总电能(kW.h)");

            if (mEnergyData == null || mEnergyData.mEnergyDataRawList.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("请先导入正确的数据文件！");
                form.setExportExcelStatus(2);
                return -1;
            }
            else
            {
                form.setExportExcelStatus(1);
            }

            int row = 2;
            for (int i = 0; i < mEnergyData.mEnergyDataRawList.Count; i++)
            {
                row = i + 2;
                object s1, s2, s3, s4, s5, s6, s7;
                s1 = BitConverter.ToUInt16(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].year), 0) + "年"
                     + Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].mouth), System.Globalization.NumberStyles.HexNumber) + "月"
                     + Int32.Parse(BitConverter.ToString(mEnergyData.mEnergyDataRawList[i].day), System.Globalization.NumberStyles.HexNumber) + "日";
                s2 = BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0);
                s3 = BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0);
                s4 = BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0);
                if (i == 0)
                {
                    s5 = BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0);
                    s6 = BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0);
                    s7 = BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0);
                }
                else
                {
                    s5 = (BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power1), 0)
                       - BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power1), 0));
                    s6 = (BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].power2), 0)
                       - BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].power2), 0));
                    s7 = (BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i].powerAll), 0)
                       - BitConverter.ToUInt32(Myutility.ToHostEndian(mEnergyData.mEnergyDataRawList[i - 1].powerAll), 0));
                }

                excel.SetData(row, 1, s1.ToString());
                excel.SetData(row, 2, s2.ToString());
                excel.SetData(row, 3, s3.ToString());
                excel.SetData(row, 4, s4.ToString());
                excel.SetData(row, 5, s5.ToString());
                excel.SetData(row, 6, s6.ToString());
                excel.SetData(row, 7, s7.ToString());
            }

            excel.SelectWorksheet(2);
            excel.SetChart(Excel.XlChartType.xlColumnClustered, "A1", "G" + row, row*10);
           
            excel.SelectWorksheet(3);
            excel.SetChart(Excel.XlChartType.xlLine, "A1", "G" + row, row * 10);
            
            excel.SaveAs(filename);

            excel.Release();

            form.setExportExcelStatus(2);
            System.Windows.Forms.MessageBox.Show("生成成功！");
            return 0;
        }

        public void SetChart( Excel.XlChartType type, string start, string end, int width)
        {
            Excel.ChartObjects xlCharts = (Excel.ChartObjects)CurXlsWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(0, 0, width, 400);
            Excel.Chart chartPage = myChart.Chart;

            Excel.Range chartRange = GetWorksheet(1).get_Range(start, end);
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = type;

            myChart.Chart.ChartWizard(chartRange, type, Type.Missing, XlRowCol.xlColumns, 1, 1, true, "电能数据", "日期", "KW.h", Type.Missing);
        }

    }
}
