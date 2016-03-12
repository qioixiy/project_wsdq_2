using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportExcel
{
    public class CBExcel
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;

        public CBExcel()
        {
        }

        public void SetData(int i, int j, string data)
        {
            xlWorkSheet.Cells[i, j] = data;
        }

        public void SetChart(string start, string end, Excel.XlChartType type)
        {
            Excel.Range chartRange;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
            Excel.Chart chartPage = myChart.Chart;

            chartRange = xlWorkSheet.get_Range(start, end);
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
            xlApp = new Excel.ApplicationClass();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        }

        public void SaveAs()
        {
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }

        public void Release()
        {
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
    }
}
