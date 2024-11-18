using Autodesk.Revit.DB.Electrical;
using RVTBootcamp_Module_03;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using System.Runtime.InteropServices;
using System.Security.RightsManagement;

namespace RVTBootcamp_Module_03
{
    [Transaction(TransactionMode.Manual)]
    public class CSV_Sheets : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            //1. Declare Variables could use @ method for single \ rather than \\

            string sheetPath = "C:\\Users\\Mespadas.BNJHBK\\Desktop\\Revit-Addins\\RVT_CSV\\RAB_Bonus_Sheets.csv";

            //2. create list of string arrays for CSV data
            List<string[]> sheetData = new List<string[]>();

            //3. read text file datas
            string[] sheetArray = System.IO.File.ReadAllLines(sheetPath);

            //4. loop through file data and put into list

            foreach (string sheetString in sheetArray)
            {
                string[] rowArray = sheetString.Split(',');
                sheetData.Add(rowArray);
            }

          


            //5. remove header row
            sheetData.RemoveAt(0);

            //6. create a transaction 

            using Transaction t = new Transaction(doc);
            {
                t.Start("Create Sheets");

                //7. loop through sheet data
                int counter = 0;

                foreach (string[] currentSheet in sheetData)
                {

                    ViewSheet newSheet;

                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

                    newSheet = ViewSheet.Create(doc, collector.FirstElementId());
                    newSheet.SheetNumber = currentSheet[0];
                    newSheet.Name = currentSheet[1];

                    counter++;
                }

              
                TaskDialog.Show("Complete", "Created" + counter.ToString() + " Sheets.");
                t.Commit();



            }




            return Result.Succeeded;
        }
        

    }
}



