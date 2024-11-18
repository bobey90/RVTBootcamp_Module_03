using Autodesk.Revit.DB.Electrical;
using RVTBootcamp_Module_03;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using System.Runtime.InteropServices;
using System.Security.RightsManagement;

namespace RVTBootcamp_Module_03
{
    [Transaction(TransactionMode.Manual)]
    public class Module03 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            //2. Create instances of class - v1
            Building theater = new Building("Grand Opera House", "5 Main Street", 4, 35000);
            Building hotel = new Building("Fancy Hotel", "10 Main Street", 10, 10000);
            Building office = new Building("Big Office Building", "15 Main Street", 15, 15000);

            //3. Create list of buildings
            List<Building> buildingList = new List<Building>();
            buildingList.Add(theater);
            buildingList.Add(hotel);
            buildingList.Add(office);
            buildingList.Add(new Building("Hospital", "20 Main Street", 20, 35000));

            //6. create instance of class and use method
            Neighborhood downtown = new Neighborhood("Downtown", "Middletown", "CT", buildingList);

            TaskDialog.Show("Test", $"There are {downtown.GetBuildingCount()}" +
                $" buildings in the {downtown.Name} neighborhood.");

            //7. Working with rooms
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);

            

            //8. Insert Family
            FamilySymbol curFS = GetFamilySymbolByName(doc, "Desk", "60\" x 30\"");

            using (Transaction t = new Transaction(doc))
            {

                t.Start("Insert family into room");

                //9.Activate Family Symbol
                curFS.Activate();

                foreach (SpatialElement room in collector)
                {
                    LocationPoint loc = room.Location as LocationPoint;
                    XYZ roomPoint = loc.Point as XYZ;

                    FamilyInstance curFi = doc.Create.NewFamilyInstance(roomPoint, curFS, StructuralType.NonStructural);

                    //10. Get parameter value
                    string name = GetParameterValueAsString(room, "Department");

                    //11. Set parameter values
                    SetParameterValue(room, "Ceiling Finish", "ACT");

                }
                t.Commit();

                //11. string splitting
                string myLine = "one, two, three, four, five";
                string[] splitLine = myLine.Split(',');
                TaskDialog.Show("Test ",splitLine[0].Trim());
                TaskDialog.Show("Test ", splitLine[3].Trim());
            }

            return Result.Succeeded;
        }

        internal string GetParameterValueAsString(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter myParam = paramList.First();

            return myParam.AsString();
        }

        internal double GetParameterValueAsDouble(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter myParam = paramList.First();

            return myParam.AsDouble();
        }

        internal void SetParameterValue(Element element, string paramName, string value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter param = paramList.First();

            param.Set(value);
        }

        internal void SetParameterValue(Element element, string paramName, double value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter param = paramList.First();

            param.Set(value);
        }

        internal FamilySymbol GetFamilySymbolByName(Document doc, string famName, string fsName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));

            foreach (FamilySymbol fs in collector)
            {
                if (fs.Name == fsName && fs.FamilyName == famName)
                    return fs;
            }
            return null;
        }

    }

public class Building
{
    public string Name { get; set; }
    public string Address { get; set; }
    public int NumFloors { get; set; }
    public double Area { get; set; }

    //3. Ad constructor to class
    public Building(string _name, string _address, int _numFloors, double _area)
    {
        Name = _name;
        Address = _address;
        NumFloors = _numFloors;
        Area = _area;

    }
}

    //4. Define dynamic class #2
    public class Neighborhood
    {

        public string Name { get; set; }
        public string City { get; set; }
        public string State { get; set; }

        public List<Building> BuildingList { get; set; }

        public Neighborhood(string _name, string _city, string _state, List<Building> _buildings)
        {
            Name = _name;
            City = _city;
            State = _state;
            BuildingList = _buildings;
        }

        //5. Add Method to class

        public int GetBuildingCount()
        {
            return BuildingList.Count;
        }
    }

}