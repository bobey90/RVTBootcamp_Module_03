using Autodesk.Revit.DB.Electrical;
using RVTBootcamp_Module_03;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using System.Runtime.InteropServices;
using System.Security.RightsManagement;
using System.Security.Cryptography.X509Certificates;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Windows.Controls;

namespace RVTBootcamp_Module_03
{
    [Transaction(TransactionMode.Manual)]
    public class Module03Challenge : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;
            /*
             * You just completed construction on your new Revit building. Now it's time to move in (virtually)! 
             * For this challenge, you will populate the provided Revit model with the furniture listed in the data files. 
             * Click the folder icon above to download the files.

                You have three options for reading the data into your add-in: 

                Use the hard-coded data in the two C# methods in the "RAB_Module 03_Furniture Data Methods.txt" file. 

                Read the "RAB_Module 03_Furniture Sets.csv" and "RAB_Module 03_Furniture Types.csv" files using the skills
                you learned in the "How to Read a CSV File in C#" bonus lesson. 

                Read the "RAB_Module 03_Furniture.xlsx" file using the skills from the "How to Read and Write Excel Files" bonus lesson.

                The "Furniture Set" data lists the contents for each room. All of the rooms in the model have a "Furniture Set" parameter 
                that corresponds with the sets listed in the worksheet. The "Furniture Type" data lists each piece of furniture's names, families, 
                and types. The furniture names correspond to those in the "Furniture Set" worksheet.  

                First, read the data into your add-in using one of the three methods listed above. Next, create classes to hold the furniture set
                and furniture type data. Remember, you can create a list that uses your custom class as a data type. Next, get all the rooms, read 
                their "Furniture Set" parameters, and insert the corresponding furniture families. Your job is to get the furniture to the room so 
                you can use the room's location point as the family insertion point. Someone else can move the furniture to the right spot!

                As the last step, update each room's "Furniture Count" parameter with the number of inserted furniture families. 
                Hint: you can add a method to your furniture set class that counts the number of families in the set.
            */

            // Your code goes here

            
           
            List <string[]> furSets_List = furSets_Data();
            List<string[]> furTypes_List = furTypes_Data();

            furTypes_List.RemoveAt(0);
            furSets_List.RemoveAt(0);

            List<FurnitureType> furn_Types = new List<FurnitureType>();

            foreach (string[] furniture_Types in furTypes_List)
            {
                FurnitureType newFurType = new FurnitureType(furniture_Types[0], furniture_Types[1], furniture_Types[2]);
                furn_Types.Add(newFurType);
                furn_Types.Add(new FurnitureType(furniture_Types[0], furniture_Types[1], furniture_Types[2]));
            }

            List<FurnitureSet> furn_Sets = new List<FurnitureSet>();

            foreach (string[] furniture_Sets in furSets_List)
            {
                FurnitureSet newFurSets = new FurnitureSet(furniture_Sets[0], furniture_Sets[1], furniture_Sets[2]);
                furn_Sets.Add(newFurSets);
              
            }


            FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
            roomCollector.OfCategory(BuiltInCategory.OST_Rooms);


            using (Transaction t = new Transaction(doc))
            {

                t.Start("Insert family into room");

                //9.Activate Family Symbol


                int counter = 0;

                foreach (SpatialElement curRoom in roomCollector)
                {

                    LocationPoint roomPoint = curRoom.Location as LocationPoint;
                    XYZ insPoint = roomPoint.Point;

                    string fur_Set_Lookup = curRoom.LookupParameter("Furniture Set").AsString();

                    foreach(FurnitureSet curSet in furn_Sets)
                    {

                        if(curSet.Set == fur_Set_Lookup)
                        {
                            foreach(string furnItems in curSet.Furniture)
                            {
                                FamilySymbol curSymbol = GetFurnitureByName(doc, furn_Types, furnItems);
                                if(curSymbol != null)
                                {
                                    FamilyInstance curFI = doc.Create.NewFamilyInstance(insPoint, curSymbol, StructuralType.NonStructural);

                                    counter++;
                                }

                            }
                        }
                 
                    }

                }


                TaskDialog.Show("Complete", "Created" + counter.ToString() + " Furniture.");
                t.Commit();

            }

            return Result.Succeeded;
        }

        private void SetParameterValue(SpatialElement curRoom, string parameterName, int countValue)
        {
            Parameter curParam = curRoom.LookupParameter(parameterName);
            if (parameterName != null)
            {
                curParam.Set(countValue);
            }
        }

        private void SetParameterValue(SpatialElement curRoom, string parameterName, string countValue)
        {
            Parameter curParam = curRoom.LookupParameter(parameterName);
            if (parameterName != null)
            {
                curParam.Set(countValue);
            }
        }

        private FamilySymbol GetFurnitureByName(Document doc, List<FurnitureType> furnitureTypes, string tmpfurnItem)
        {

            string furnItem_space = tmpfurnItem.Trim();
            foreach(FurnitureType curType in furnitureTypes)
            {
                if (curType.Name == furnItem_space)
                {
                    FamilySymbol curFS = GetFamilySymbolByName(doc, curType.FamilyName, curType.TypeName);

                    if(curFS != null)
                    {
                        if(curFS.IsActive == false)
                        {
                            curFS.Activate();
                        }

                    }
                    return curFS;
                }
              
            }
            return null;
        }

        private FamilySymbol GetFamilySymbolByName(Document doc, string familyName, string typeName)
        {
            FilteredElementCollector symbolCollector = new FilteredElementCollector(doc);
            symbolCollector.OfClass(typeof(FamilySymbol));

            foreach(FamilySymbol curFS in symbolCollector)
            {
                if(curFS.FamilyName == familyName && curFS.Name == typeName)
                {
                    return curFS;
                }
            }
            return null;

        }

        private List<string[]> furSets_Data()
        {
            string FurSets_Path;
            string[] FurSets_Array;
            List<string[]> furSets_Items = new List<string[]>();
            FurSets_Path = "C:\\Users\\Mespadas.BNJHBK\\Desktop\\Revit API Module 3\\RAB_Module 03_Furniture Sets.csv";
            FurSets_Array = System.IO.File.ReadAllLines(FurSets_Path);

            foreach (string FurSets_String in FurSets_Array)
            {
                String[] rowArray = FurSets_String.Split(new[] { "," }, 3, StringSplitOptions.RemoveEmptyEntries); ;
                furSets_Items.Add(rowArray);
            }
            return furSets_Items;
        }

        
        private List<string[]> furTypes_Data() {

            string FurTypes_Path;
            string[] FurTypes_Array;
            List<string[]> furTypes_Items = new List<string[]>();
            FurTypes_Path = "C:\\Users\\Mespadas.BNJHBK\\Desktop\\Revit API Module 3\\RAB_Module 03_Furniture Types.csv";
            furTypes_Items = new List<string[]>();
            FurTypes_Array = System.IO.File.ReadAllLines(FurTypes_Path);

            foreach (string FurTypes_String in FurTypes_Array)
            {
                string[] rowArray = FurTypes_String.Split(',');
                    furTypes_Items.Add(rowArray);
            }
            return furTypes_Items;
        }
    }

}

