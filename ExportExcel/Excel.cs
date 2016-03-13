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
        Excel.Worksheet CurXlsWorkSheet;
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

        public void SetChart(string start, string end, Excel.XlChartType type)
        {

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)CurXlsWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 300, 250);
            Excel.Chart chartPage = myChart.Chart;

            Excel.Range chartRange = GetWorksheet(1).get_Range(start, end);
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
            // default sheet1
            xlsWorkBook = xlsApp.Workbooks.Add(misValue);
            // add 2s sheet
            xlsApp.Worksheets.Add(misValue);
            xlsApp.Worksheets.Add(misValue);
            GetWorksheet(1).Name = "电能";
            GetWorksheet(2).Name = "柱状图";
            GetWorksheet(3).Name = "电量曲线";
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

    }
}
