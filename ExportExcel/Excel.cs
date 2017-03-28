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
            public RowColumData(int row, int col, string data) {
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
            GetWorksheet(2).Name = "阶段电能柱状图";
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
                
        public static int GenExcel(FormMain form, EnergyData mEnergyData, string filename)
        {
            CBExcel excel = new CBExcel();

            excel.Create();

            // append file ext
            if (Convert.ToDouble(excel.xlsApp.Version) >= 12.0)//office 2007
            {
                filename += ".xlsx";
            }
            else
            {
                filename += ".xls";
            }

            // worksheet1 电能列表
            excel.SelectWorksheet(1);
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["A", Type.Missing]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["A", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["B", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["C", Type.Missing]).ColumnWidth = 20;
            ((Excel.Range)excel.CurXlsWorkSheet.Columns["D", Type.Missing]).ColumnWidth = 20;
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, 7]).Font.Name = "微软雅黑";
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, 7]).Font.Bold = true;
            excel.CurXlsWorkSheet.get_Range(excel.CurXlsWorkSheet.Cells[1, 1], excel.CurXlsWorkSheet.Cells[1, 7]).Interior.ColorIndex = 16;

            excel.SetData(1, 1, "时间");
            excel.SetData(1, 2, "消耗电能(kW.h)");
            excel.SetData(1, 3, "再生电能(kW.h)");
            excel.SetData(1, 4, "总消耗能量(kW.h)");

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

                string year = mEnergyData.mEnergyDataRawList[i].getYear().ToString();
                string mouth = mEnergyData.mEnergyDataRawList[i].getMouth().ToString();
                string day = mEnergyData.mEnergyDataRawList[i].getDay().ToString();
                string hour = mEnergyData.mEnergyDataRawList[i].getHour().ToString();
                string minute = mEnergyData.mEnergyDataRawList[i].getMinute().ToString();
                string consumePower = "0";
                string revivePower = "0";
                string totalPower = "0";

                if (i != 0)
                {
                    consumePower = (mEnergyData.mEnergyDataRawList[i].getConsumePower() - mEnergyData.mEnergyDataRawList[i-1].getConsumePower()).ToString();
                    revivePower = (mEnergyData.mEnergyDataRawList[i].getRevivePower() - mEnergyData.mEnergyDataRawList[i - 1].getRevivePower()).ToString();
                    totalPower = (mEnergyData.mEnergyDataRawList[i].getTotalPower() - mEnergyData.mEnergyDataRawList[i - 1].getTotalPower()).ToString();
                }

                string datetime = year + "年" + mouth + "月" + day + "日" + hour + "时" + minute  + "分";


                excel.SetData(row, 1, datetime);
                excel.SetData(row, 2, consumePower);
                excel.SetData(row, 3, revivePower);
                excel.SetData(row, 4, totalPower);

                form.setExportExcelStatus("processing", "S1:" + i + "/" + mEnergyData.mEnergyDataRawList.Count);
            }

            excel.CurXlsWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;

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
