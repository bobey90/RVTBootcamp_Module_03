using Autodesk.Revit.DB.Electrical;
using RVTBootcamp_Module_03;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using System.Runtime.InteropServices;
using System.Security.RightsManagement;

namespace RVTBootcamp_Module_03
{
    [Transaction(TransactionMode.Manual)]
    public class CSV_Levels : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            //1. Declare Variables could use @ method for single \ rather than \\

            string levelPath = "C:\\Users\\Mespadas.BNJHBK\\Desktop\\Revit-Addins\\RVT_CSV\\RAB_Bonus_Levels.csv";

            //2. create list of string arrays for CSV data
            List<string[]> levelData = new List<string[]>();

            //3. read text file datas
            string[] levelArray = System.IO.File.ReadAllLines(levelPath);

            //4. loop through file data and put into list

            foreach (string levelString in levelArray)
            {
                string[] rowArray = levelString.Split(',');
                levelData.Add(rowArray);
            }


            //5. remove header row
            levelData.RemoveAt(0);

            //6. create a transaction 

            using Transaction t = new Transaction(doc);
            {
                t.Start("Create Levels");

                //7. loop through level data
                int counter = 0;

                foreach (string[] currentLevelData in levelData)
                {   
                    //8. create height variables
                    double heightFeet = 0;
                    double heightMeters = 0;

                    //9. get height and convert from string to double
                    bool convertFeet = double.TryParse(currentLevelData[1], out heightFeet);
                    bool convertMeters = double.TryParse(currentLevelData[2], out heightMeters);

                    //10. if using metric, convert meters to feet
                    double heightMetersConvert = heightMeters * 3.28084;
                    double heightMetersConvert2 = UnitUtils.ConvertToInternalUnits(heightMeters, UnitTypeId.Meters);

                    //11. create level and rename
                    Level currentLevel = Level.Create(doc, heightFeet);
                    currentLevel.Name = currentLevelData[0];

                    //12. increment counter
                    counter++;

                }

                //14. tell user whats happening
                TaskDialog.Show("Complete", "Created" + counter.ToString() + " Levels.");
                t.Commit();



            }




            return Result.Succeeded;
        }
        

    }
}



