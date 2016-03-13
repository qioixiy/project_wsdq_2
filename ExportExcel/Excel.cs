using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportExcel
{
    public class CBExcel
    {
        Excel.Application xlsApp;
        Excel.Workbook xlsWorkBook;
        Excel.Worksheet xlsWorkSheet;
        object misValue = System.Reflection.Missing.Value;

        public CBExcel()
        {
        }

        public void SetData(int i, int j, string data)
        {
            xlsWorkSheet.Cells[i, j] = data;
        }

        public void SetChart(string start, string end, Excel.XlChartType type)
        {
            Excel.Range chartRange;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlsWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
            Excel.Chart chartPage = myChart.Chart;

            chartRange = xlsWorkSheet.get_Range(start, end);
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = type;
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
            xlsWorkBook = xlsApp.Workbooks.Add(misValue);
            xlsApp.Worksheets.Add(misValue);
            xlsWorkSheet = (Excel.Worksheet)xlsWorkBook.Worksheets.get_Item(1);
            //xlsWorkSheet.Name = "电能";
        }

        public void SaveAs(String filename)
        {
            //xlsApp.DisplayAlerts = false;
            //xlsWorkBook.Close(true, misValue, misValue);
            //xlsApp.Quit();

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
            xlsWorkSheet = null;
        }

        public void Release()
        {
            releaseObject(xlsWorkSheet);
            releaseObject(xlsWorkBook);
            releaseObject(xlsApp);
        }

    }
}
