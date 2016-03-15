using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExportExcel
{
    class ExportExcelThread
    {
        Form form;
        string filename;
        EnergyData mEnergyData;

        public ExportExcelThread(Form _form, EnergyData _EnergyData, string _filename)
        {
            form = _form;
            mEnergyData = _EnergyData;
            filename = _filename;
        }
        public void ThreadMethod()
        {
            CBExcel.GenExcel((FormMain)form, mEnergyData, filename);
        }
    }
}
