﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace CatalogPrinter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public Microsoft.Office.Interop.Excel.Application Excel { get; set; }
        public Workbook Workbook { get; set; }
        public Workbook Workbook2Print { get; set; }

        public bool ErrorOccured { get; set; } = false;

        public List<string> SheetOrder { get; set; } = new List<string>();

        private readonly string _tmpExcel = @"C:\Temp\Test.xlsx";

        public MainWindow()
        {
            InitializeComponent();
            CatalogTypeComboBox.Items.Add(new CatalogType("Dakwerker", 1));
            CatalogTypeComboBox.Items.Add(new CatalogType("Veranda", 2));
            CatalogTypeComboBox.Items.Add(new CatalogType("Aannemer", 3));
            CatalogTypeComboBox.Items.Add(new CatalogType("Particulier", 4));
            //CatalogTypeComboBox.IsDropDownOpen = true;
        }

        private void SheetOrderInput_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void WorkbookInputFile_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Print_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // check if file exists
                if(!File.Exists(WorkbookInputFile.Text))
                    throw new Exception($"Workbook " + WorkbookInputFile.Text + " not found!");
                string outputDir = new FileInfo(WorkbookInputFile.Text).Directory.FullName;

                // get sheet order input
                string sheetOrderString = SheetOrderInput.Text.Replace("\r", string.Empty);
                SheetOrder = sheetOrderString.Split('\n').Where(s => s != "").ToList();
                if (SheetOrder.Count == 0)
                    throw new Exception($"No sheet entries found!");

                //// set PDF pinter name
                //var printers = System.Drawing.Printing.PrinterSettings.InstalledPrinters;
                //string pdfPrinter = "PDF Complete";

                // start Excel 
                // open temp workbook to which the sheets of interest are copied to
                Excel = new Microsoft.Office.Interop.Excel.Application();
                Excel.DisplayAlerts = false;
                Workbook2Print = Excel.Workbooks.Add();
                Workbook2Print?.SaveAs(_tmpExcel);
                Workbook2Print?.Close(true);
                Workbook2Print = Excel.Workbooks.Open(_tmpExcel);

                // open workbook
                string file = WorkbookInputFile.Text;
                Workbook = Excel.Workbooks.Open(file);

                foreach (var shName in SheetOrder)
                {
                    if(Workbook.Sheets[shName] == null)
                        throw new Exception($"Sheet " + shName + " not found in workbook " + WorkbookInputFile.Text);
                    Workbook.Sheets[shName].Copy(After: Workbook2Print.Sheets[Workbook2Print.Sheets.Count]);
                }
                // delete default first sheet "Sheet1" on creation of workbook
                Workbook2Print.Activate();
                Workbook2Print.Worksheets[1].Delete();

                // format and print sheets
                string outputFile = outputDir + @"\Catalog.pdf";
                foreach (Worksheet sh in Workbook2Print.Worksheets)
                {
                    FormatSheet(sh);
                }
                //Workbook2Print.Worksheets.PrintOutEx(PrintToFile: true, PrToFileName: outputFile, ActivePrinter: pdfPrinter);
                Workbook2Print.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputFile);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Workbook?.Close(false);
                Workbook2Print?.Close(true);
                Excel?.Quit();
                if(Excel!=null)
                    Marshal.ReleaseComObject(Excel);
                if(Workbook!=null)
                    Marshal.ReleaseComObject(Workbook);
            }
        }

        private void FormatSheet(Worksheet sh)
        {
            string leftHeader = (sh.Cells[16, 2] as Range).Value as string ?? "null";
            string centerHeader = (sh.Cells[17, 2] as Range).Value as string ?? "null";
            var rightHeaderDate = ((sh.Cells[18, 2] as Range).Value);
            string rightHeader = "null";
            if (rightHeaderDate != null)
                rightHeader = rightHeaderDate.ToString("dd/MM/yyyy");
            string leftFooter = (sh.Cells[19, 2] as Range).Value as string ?? "null";
            string rightFooter = (sh.Cells[20, 2] as Range).Value as string ?? "null";

            sh.PageSetup.LeftHeader = leftHeader;
            sh.PageSetup.CenterHeader = centerHeader;
            sh.PageSetup.RightHeader = rightHeader;
            sh.PageSetup.LeftFooter = leftFooter;
            sh.PageSetup.RightFooter = rightFooter;

            sh.PageSetup.PrintArea = "D2:I30";

            sh.PageSetup.Zoom = false;
            sh.PageSetup.FitToPagesWide = 1;
            sh.PageSetup.FitToPagesTall = 1;
            sh.PageSetup.CenterVertically = true;
            sh.PageSetup.CenterHorizontally = true;

            sh.PageSetup.LeftMargin = Excel.InchesToPoints(0.7);
            sh.PageSetup.RightMargin = Excel.InchesToPoints(0.7);
            sh.PageSetup.TopMargin = Excel.InchesToPoints(0.75);
            sh.PageSetup.BottomMargin = Excel.InchesToPoints(0.75);
            sh.PageSetup.HeaderMargin = Excel.InchesToPoints(0.3);
            sh.PageSetup.FooterMargin = Excel.InchesToPoints(0.3);
        }
    }
}