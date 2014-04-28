using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    class ExcelAccess : IDisposable
    {
        String m_path = null;
        Application m_objExcel = null;
        Workbook m_objWBook = null;


        public ExcelAccess(String path)
        {
            m_path = path;
        }

        public void Do(String newPath)
        {
            TemplateWorkbook.SaveAs(newPath);

           // 与えられたワークシート名から、Worksheetオブジェクトを得る
            string sheetName = "Sheet2";
            int indexSheet = GetSheetIndex(sheetName, TemplateWorkbook.Sheets);
            if (indexSheet > 0)
            {
                Sheets objSheets = TemplateWorkbook.Sheets;
                Worksheet objSheet = objSheets[indexSheet];

                object[,] rng = null;
                Range objRange = objSheet.get_Range("A1", "C1");
                rng = (System.Object[,])objRange.get_Value(Type.Missing);

                rng[1, 1] = "MORE";

                objRange.set_Value(Type.Missing, rng);

                Marshal.ReleaseComObject(objRange);
                Marshal.ReleaseComObject(objSheet);
                Marshal.ReleaseComObject(objSheets);

            }
            TemplateWorkbook.Save();
        }

        // 指定されたワークシート名のインデックスを返すメソッド
        private int GetSheetIndex(string searchName, Sheets sheets)
        {
            int i = 0;
            foreach (Worksheet sheet in sheets)
            {
                String sheetName = sheet.Name;
                Marshal.ReleaseComObject(sheet);
                if (searchName == sheetName)
                {
                    Marshal.ReleaseComObject(sheets);
                    return i + 1;                    
                }
            }
            Marshal.ReleaseComObject(sheets);
            return 0;
        }

        private Application ExcelApplication
        {
            get
            {
                if (this.m_objExcel == null)
                {
                    this.m_objExcel = new Application();
                }
                //m_objExcel.Visible = true;
                return this.m_objExcel;
            }
        }

        private Workbook TemplateWorkbook
        {
            get
            {
                if (this.m_objWBook == null)
                {
                    String templatePath = m_path;
                    if (String.IsNullOrEmpty(m_path))
                    {
                        templatePath = AppDomain.CurrentDomain.BaseDirectory;
                        var directory = new DirectoryInfo(templatePath + @"\Teplates\");
                        var file = directory.GetFiles().FirstOrDefault();
                        templatePath = file.FullName;
                    }

                    ExcelApplication.DisplayAlerts = false;
                    Workbooks books = ExcelApplication.Workbooks;
                    m_objWBook = books.Open(templatePath);
                    Marshal.ReleaseComObject(books);
                    ExcelApplication.DisplayAlerts = true;
//                  ExcelApplication.Calculation = XlCalculation.xlCalculationManual;
                }
                return m_objWBook;
            }
        }

        public void Dispose()
        {
            if (m_objWBook != null)
            {
                m_objWBook.Close(false);
                Marshal.ReleaseComObject(m_objWBook);
            }

            if (m_objExcel != null)
            {
                m_objExcel.Quit();
                System.Threading.Thread.Sleep(2000);
                Marshal.ReleaseComObject(m_objExcel);
            }
        }

    }
}
