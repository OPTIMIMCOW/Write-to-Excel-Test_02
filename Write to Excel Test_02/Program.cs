using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Write_to_Excel_Test_02
{
    class Program
    {
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(@"C:\Users\kallan\OneDrive - Thornton Tomasetti, Inc-\RnD\Carbon Calculator\Research\Programming\GH Excel Interface\Visual Studio\Example1.xlsx"));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[3, 3] = "First Name";
                oSheet.Cells[4, 2] = "Last Name";
                oSheet.Cells[5, 2] = "Full Name";
                oSheet.Cells[3, 4] = "Salary";

                // get value of a specific cell in the spredsheet
                var cellValue = (string)(oSheet.Cells[3, 2] as Microsoft.Office.Interop.Excel.Range).Value;
                Console.Write(cellValue);
                Console.ReadLine();



                oXL.Visible = false;
                oXL.UserControl = false;
                oWB.SaveAs(@"C:\Users\kallan\OneDrive - Thornton Tomasetti, Inc-\RnD\Carbon Calculator\Research\Programming\GH Excel Interface\Visual Studio\Example1_alter.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                oWB.Close();
                oXL.Quit();

                

            }

            catch (Exception e)
            {
                
            }
     

        }
    }
}
