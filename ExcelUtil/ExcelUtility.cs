using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUtil
{
    public static class ExcelUtility
    {
        private static Application XlApp { get; set; }

        static ExcelUtility()
        {
            try
            {
                XlApp = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                XlApp.DisplayAlerts = false;
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("0x800401E3 (MK_E_UNAVAILABLE)"))
                {
                    XlApp = new Application();
                    XlApp.DisplayAlerts = false;
                }
                else
                {
                    throw;
                }
            }
        }

        public static Workbook GetWorkbook(string fullName, string password = null)
        {
            try
            {
                if (password != null)
                    return XlApp.Workbooks.Open(fullName, Password: password);


                if (XlApp.Workbooks[fullName] == null)
                        XlApp.Workbooks.Open(fullName);

                return XlApp.Workbooks[fullName];
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

    }
}
