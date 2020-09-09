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
                oXL.Visible = false;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(@"C:\Users\kallan\OneDrive - Thornton Tomasetti, Inc-\RnD\Carbon Calculator\Research\Programming\GH Excel Interface\Visual Studio\Example5.xlsx"));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                // alternative to get named range
                var value = oSheet.get_Range("namedTable").Value;//get value of named range
                // alter value in named range
                //oSheet.get_Range("namedTable").Value2 = "Test"; // this works

                //Microsoft.Office.Interop.Excel.Range testSubSet = oSheet.get_Range("namedTable").Rows["1:1"]; // this works
                Microsoft.Office.Interop.Excel.Range testSubSet = oSheet.get_Range("namedTable").Cells[1,1]; // this also works
                testSubSet.Value2 = "Updated SubSet";

                //// get named table
                //String myTableName = "namedTable";
                //Microsoft.Office.Interop.Excel.Range myTable = oXL.Range[myTableName];

                //// get value in named table
                //Microsoft.Office.Interop.Excel.WorksheetFunction VariableObject = oXL.WorksheetFunction;

                //double rowRef = 1;
                //double colRef = 1;

                //Object displayVariable = VariableObject.Index(myTable, rowRef, colRef);


                //String myString = (String) displayVariable;

                ////Add table headers going cell by cell.
                //// oSheet.Cells[3, 2] = 3;


                //// get value of a specific cell in the spredsheet


                //Console.Write(myString);
                //Console.ReadLine();



                oXL.Visible = false;
                oXL.UserControl = false;
                oWB.SaveAs(@"C:\Users\kallan\OneDrive - Thornton Tomasetti, Inc-\RnD\Carbon Calculator\Research\Programming\GH Excel Interface\Visual Studio\Example5_alter.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
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
