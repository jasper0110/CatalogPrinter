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
using System.Xml;
using System.Configuration;
using Encrypter;
using ExcelUtil;

namespace CatalogPrinter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public Workbook Workbook { get; set; }
        public Workbook WorkbookSheetOrder { get; set; }
        public Workbook Workbook2Print { get; set; }

        private readonly string _tmpWorkbookDir = @"C:\temp";
        private readonly string _tmpWokbookName = @"\temp.xlsx";

        public MainWindow()
        {
            InitializeComponent();

            // add items to combobox
            //CatalogTypeComboBox.Items.Add(new CatalogType("Selecteer cataloog type", CatalogTypeEnum.NONSELECTED));
            CatalogTypeComboBox.Items.Add(new CatalogType("Dakwerker", CatalogTypeEnum.DAKWERKER));
            CatalogTypeComboBox.Items.Add(new CatalogType("Veranda", CatalogTypeEnum.VERANDA));
            CatalogTypeComboBox.Items.Add(new CatalogType("Aannemer", CatalogTypeEnum.AANNEMER));
            CatalogTypeComboBox.Items.Add(new CatalogType("Particulier", CatalogTypeEnum.PARTICULIER));

            // set inital selection of combobox
            foreach(var item in CatalogTypeComboBox.Items)
            {
                CatalogType type = item as CatalogType;
                if (type.Value == CatalogTypeEnum.DAKWERKER)
                    CatalogTypeComboBox.SelectedIndex = CatalogTypeComboBox.Items.IndexOf(item);

            }           
        }

        private void Print_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // open config
                string configPath = Directory.GetCurrentDirectory() + @"\CatalogPrinter.config";
                if (!File.Exists(configPath))
                    throw new Exception($"Config file " + configPath + " not found!");
                ExeConfigurationFileMap configMap = new ExeConfigurationFileMap();
                configMap.ExeConfigFilename = configPath;
                Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);
                var appSettings = config.GetSection("appSettings") as AppSettingsSection;
                // get config values
                string hash = appSettings.Settings["password"]?.Value;
                string masterCatalog = appSettings.Settings["masterCatalog"]?.Value;
                string clientCatalog = appSettings.Settings["clientCatalog"]?.Value;
                string outputPath = appSettings.Settings["outputPath"]?.Value;

                if(hash == null)
                    throw new Exception($"Could not find 'password' key in " + configPath + "!");
                if (masterCatalog == null)
                    throw new Exception($"Could not find 'masterCatalog' key in " + configPath + "!");
                if (clientCatalog == null)
                    throw new Exception($"Could not find 'clientCatalog' key in " + configPath + "!");
                if (outputPath == null)
                    throw new Exception($"Could not find 'outputPath' key in " + configPath + "!");

                // check if files exists
                if (!File.Exists(masterCatalog))
                    throw new Exception($"Workbook " + masterCatalog + " not found!");
                if (!File.Exists(clientCatalog))
                    throw new Exception($"Workbook " + clientCatalog + " not found!");

                // open master workbook
                string password = HashUtil.Decrypt(hash);
                Workbook = ExcelUtility.GetWorkbook(masterCatalog, password);

                // get catalog type
                string catalogType = CatalogTypeComboBox.SelectedItem.ToString();

                // open temp workbook to which the sheets of interest are copied to
                Workbook2Print = ExcelUtility.XlApp.Workbooks.Add();
                if (!Directory.Exists(_tmpWorkbookDir))
                    Directory.CreateDirectory(_tmpWorkbookDir);
                Workbook2Print?.SaveAs(_tmpWorkbookDir + _tmpWokbookName);

                // get sheet order to print
                var sheetOrder = GetSheetOrder(catalogType, clientCatalog);

                string leftHeader = "null", centerHeader = "null", rightHeader = "null", leftFooter = "null", rightFooter = "null";

                // copy necessary sheets to temp workbook and put sheets in correct order
                foreach (var shName in sheetOrder)
                {
                    if (ExcelUtility.GetWorksheetByName(Workbook, shName) == null)
                        throw new Exception($"Sheet " + shName + " not found in workbook " + masterCatalog + "!" +
                            "\nPlease check the sheet order input in " + clientCatalog + " for " + catalogType + ".");
                    // set catalog type
                    Workbook.Sheets[shName].Cells[11, 2] = catalogType;

                    leftHeader = (Workbook.Sheets[shName].Cells[16, 2] as Range).Value as string ?? "null";
                    centerHeader = (Workbook.Sheets[shName].Cells[17, 2] as Range).Value as string ?? "null";
                    var rightHeaderDate = ((Workbook.Sheets[shName].Cells[18, 2] as Range).Value);
                    rightHeader = "null";
                    if (rightHeaderDate != null)
                        rightHeader = rightHeaderDate.ToString("dd/MM/yyyy");
                    leftFooter = (Workbook.Sheets[shName].Cells[19, 2] as Range).Value as string ?? "null";
                    rightFooter = (Workbook.Sheets[shName].Cells[20, 2] as Range).Value as string ?? "null";

                    // copy sheet
                    if (catalogType == "Particulier")
                    {
                        //SetBtwField(Workbook.Sheets[shName], true);
                        Workbook.Sheets[shName].Copy(After: Workbook2Print.Sheets[Workbook2Print.Sheets.Count]);
                        Workbook2Print.Sheets[Workbook2Print.Sheets.Count].Cells[8, 2] = "ja";
                        //SetBtwField(Workbook.Sheets[shName], false);
                        Workbook.Sheets[shName].Copy(After: Workbook2Print.Sheets[Workbook2Print.Sheets.Count]);
                        Workbook2Print.Sheets[Workbook2Print.Sheets.Count].Cells[8, 2] = "neen";
                    }
                    else
                    {
                        Workbook.Sheets[shName].Copy(After: Workbook2Print.Sheets[Workbook2Print.Sheets.Count]);
                    }
                }
                // delete default first sheet on creation of workbook
                Workbook2Print.Activate();
                Workbook2Print.Worksheets[1].Delete();

                // format and print sheets
                string outputFile = outputPath + @"\catalog.pdf";
                if (ExcelUtility.IsFileInUse(outputFile))
                    throw new Exception(outputFile + " is open, please close it and press 'Print' again.");
                foreach (Worksheet sh in Workbook2Print.Worksheets)
                    FormatSheet(sh, leftHeader, centerHeader, rightHeader, leftFooter, rightFooter);
                Workbook2Print.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputFile, OpenAfterPublish: true);                

                MessageBox.Show("Done!");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ExcelUtility.CloseWorkbook(Workbook, false);
                ExcelUtility.CloseWorkbook(WorkbookSheetOrder, false);
                ExcelUtility.CloseWorkbook(Workbook2Print, true);
                File.Delete(_tmpWorkbookDir + _tmpWokbookName);

                ExcelUtility.CloseExcel();
            }
        }

        private void SetBtwField(Worksheet sh, bool btwIncl)
        {
            string value = btwIncl ? "ja" : "neen";
            Range btwCell = sh.Cells[8, 2];
            sh.Cells[8, 2] = value;
        }

        private void FormatSheet(Worksheet sh, string leftHeader, string centerHeader, string rightHeader, string leftFooter, string rightFooter)
        {    
            sh.PageSetup.LeftHeader = leftHeader;
            sh.PageSetup.CenterHeader = centerHeader;
            sh.PageSetup.RightHeader = rightHeader;
            sh.PageSetup.LeftFooter = leftFooter;
            sh.PageSetup.RightFooter = rightFooter;

            sh.PageSetup.PrintArea = "D2:N30";

            sh.PageSetup.Zoom = false;
            sh.PageSetup.FitToPagesWide = 1;
            sh.PageSetup.FitToPagesTall = 1;
            sh.PageSetup.CenterVertically = true;
            sh.PageSetup.CenterHorizontally = true;

            sh.PageSetup.LeftMargin = ExcelUtility.XlApp.InchesToPoints(0.7);
            sh.PageSetup.RightMargin = ExcelUtility.XlApp.InchesToPoints(0.7);
            sh.PageSetup.TopMargin = ExcelUtility.XlApp.InchesToPoints(0.75);
            sh.PageSetup.BottomMargin = ExcelUtility.XlApp.InchesToPoints(0.75);
            sh.PageSetup.HeaderMargin = ExcelUtility.XlApp.InchesToPoints(0.3);
            sh.PageSetup.FooterMargin = ExcelUtility.XlApp.InchesToPoints(0.3);
        }

        private List<string> GetSheetOrder(string catalogType, string fullName)
        {
            List<string> sheetOrder = new List<string>();

            // open sheet order workbook
            //WorkbookSheetOrder = Excel.Workbooks.Open(SheetOrderInputFile.Text);
            WorkbookSheetOrder = ExcelUtility.GetWorkbook(fullName);

            // get sheet order
            int startCol = 1;
            int maxCol = 100;
            int selectedCol = 0;
            for (int i = startCol; i < maxCol; i++)
            {
                Range currentRange = WorkbookSheetOrder.Worksheets[1].Cells[1, i] as Range;
                if (currentRange?.Value?.ToString() == catalogType)
                {
                    selectedCol = i;
                    break;
                }
            }
            if (selectedCol < 1)
                throw new Exception(catalogType + " not found in sheet " + WorkbookSheetOrder.Worksheets[1].Name 
                    + " of workbook " + WorkbookSheetOrder.Name);

            Range startCell = WorkbookSheetOrder.Worksheets[1].Cells[2, selectedCol];
            int lastRow = WorkbookSheetOrder.Worksheets[1].Cells[2, selectedCol].End(XlDirection.xlDown).Row;
            Range endCell = WorkbookSheetOrder.Worksheets[1].Cells[lastRow, selectedCol];
            Range sheetsToPrint = WorkbookSheetOrder.Worksheets[1].Range[startCell, endCell];
            foreach (var cell in sheetsToPrint)
            {
                Range cellRange = cell as Range;
                var cellValue = cellRange?.Value;
                string cellString = cellValue.ToString();
                sheetOrder.Add(cellString);
            }
            return sheetOrder;
        }

        private void CatalogTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CatalogTypeComboBox.Text = CatalogTypeComboBox.SelectedItem.ToString();
        }
    }
}
