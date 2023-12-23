using System;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp3
{
    class RepExcel : IDisposable
    {
        public Excel.Application excelapp;
        Excel.Workbooks excelappworkbooks;
        Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets; // лист в екселе
        private Excel.Worksheet excelworksheet; // ячейка
        private Excel.Range excelcells; // диапазон ячеек
                                        // Конструктор
        public RepExcel()
        {
            excelapp = new Excel.Application();
            excelapp.Visible = false;
        }
        // Деструктор
        public void Dispose()
        {
            // Release COM objects (very important!)
            if (excelapp != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp);
            if (excelappworkbooks != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelappworkbooks);
            if (excelappworkbook != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelappworkbook);
            if (excelsheets != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelsheets);
            if (excelworksheet != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworksheet);
            if (excelcells != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelcells);
            excelapp = null;
            excelappworkbooks = null;
            excelappworkbook = null;
            excelsheets = null;
            excelworksheet = null;
            excelcells = null;
            GC.Collect();
            // GC.GetTotalMemory(true);
        }
        // Coхранение книги с заданным именем
        public void CreateNewBook(string fullPathAndFilename)
        {

            excelapp.SheetsInNewWorkbook = 5;

            excelapp.Workbooks.Add(Type.Missing);
            excelapp.DisplayAlerts = false;
            //Получаем набор ссылок на объекты Workbook (на созданные книги)
            excelappworkbooks = excelapp.Workbooks;
            //Получаем ссылку на книгу 1 - нумерация от 1
            excelappworkbook = excelappworkbooks[1];
            excelsheets = excelappworkbook.Worksheets;
            //Получаем ссылку на лист 1
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            excelappworkbook.Saved = true;
            excelappworkbook.SaveAs(fullPathAndFilename, Excel.XlFileFormat.xlExcel7,
            //object FileFormat
            Type.Missing, //object Password
            Type.Missing, //object WriteResPassword
            Type.Missing, //object ReadOnlyRecommended
            Type.Missing, //object CreateBackup
            Excel.XlSaveAsAccessMode.xlNoChange,//XlSaveAsAccessMode AccessMode
            Type.Missing, //object ConflictResolution
            Type.Missing, //object AddToMru
            Type.Missing, //object TextCodepage
            Type.Missing, //object TextVisualLayout
            Type.Missing);
            excelapp.Workbooks.Close();
            excelapp.Quit();



            Dispose();

        }
        public void OpenBook(string fullPathAndFilename)
        {

            excelapp.Workbooks.Open(fullPathAndFilename,
            Type.Missing, false, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

        }
        public void CloseBook()
        {
            excelapp.Workbooks.Close();
            excelapp.Quit();
        }
        public void Save(string fullPathAndFilename)
        {

            excelappworkbooks = excelapp.Workbooks;
            //Получаем ссылку на книгу 1 - нумерация от 1
            excelappworkbook = excelappworkbooks[1];
            excelappworkbook.Saved = false;
            excelappworkbook.Save();

        }
        public void SaveAs(string fullPathAndFilename)
        {

            excelappworkbooks = excelapp.Workbooks;
            //Получаем ссылку на книгу 1 - нумерация от 1
            excelappworkbook = excelappworkbooks[1];
            excelappworkbook.Saved = true;
            excelappworkbook.SaveAs(fullPathAndFilename,
            Excel.XlFileFormat.xlExcel7, //object FileFormat
            Type.Missing, //object Password
            Type.Missing, //object WriteResPassword
            false, //object ReadOnlyRecommended
            Type.Missing, //object CreateBackup
            Excel.XlSaveAsAccessMode.xlNoChange,//XlSaveAsAccessMode AccessMode
            Type.Missing, //object ConflictResolution
            Type.Missing, //object AddToMru
            Type.Missing, //object TextCodepage
            Type.Missing, //object TextVisualLayout
            Type.Missing);

        }
        public void SetValue(string pageName, string address, string StrValues, string
        typeValue, bool isBold = false) // &quot;A10&quot;, &quot;значение&quot;
        {
            excelappworkbooks = excelapp.Workbooks;
            //Получаем ссылку на книгу 1 - нумерация от 1
            excelappworkbook = excelappworkbooks[1];
            excelsheets = excelappworkbook.Worksheets;

            excelworksheet = (Excel.Worksheet)excelsheets[pageName];
            //MessageBox.Show(&quot;Страница найдена&quot;);

            excelsheets.Add();
            excelworksheet =
            (Excel.Worksheet)excelsheets.get_Item(excelsheets.Count);
            excelworksheet.Name = pageName;

            excelcells = excelworksheet.get_Range(address, address);

        }
        public string GetValue(string pageName, string address)
        {
            excelappworkbooks = excelapp.Workbooks;
            //Получаем ссылку на книгу 1 - нумерация от 1
            excelappworkbook = excelappworkbooks[1];
            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets[pageName];
            excelcells = excelworksheet.get_Range(address, address);
            return Convert.ToString(excelcells.Value2);
        }
        public void HidenRow(string pageName, int indexRow)
        {
            excelappworkbooks = excelapp.Workbooks;
            //Получаем ссылку на книгу 1 - нумерация от 1
            excelappworkbook = excelappworkbooks[1];
            excelsheets = excelappworkbook.Worksheets;
            //Получаем ссылку на лист
            excelworksheet = (Excel.Worksheet)excelsheets[pageName];

        }
        public void DisplayLine(string pageName, int indexRow)
        {
            excelappworkbooks = excelapp.Workbooks;
            //Получаем ссылку на книгу 1 - нумерация от 1
            excelappworkbook = excelappworkbooks[1];
            excelsheets = excelappworkbook.Worksheets;
            //Получаем ссылку на лист
            excelworksheet = (Excel.Worksheet)excelsheets[pageName];

        }
    }
}



