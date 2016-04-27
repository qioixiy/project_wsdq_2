[merge exe and dll common]
"C:\Program Files (x86)\Microsoft\ILMerge\ILMerge.exe" /ndebug /target:winexe  /out:ExportExcelNew.exe ExportExcel.exe /log Microsoft.Office.Interop.Excel.dll

[merge exe and dll]
"C:\Program Files (x86)\Microsoft\ILMerge\ILMerge.exe" /ndebug /target:winexe  /out:ExportExcelNew.exe ExportExcel.exe /log Microsoft.Office.Interop.Excel.dll /lib:"C:\Program Files (x86)\Microsoft Visual Studio 11.0\Visual Studio Tools for Office\PIA\Office14"