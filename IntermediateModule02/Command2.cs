using Autodesk.Revit.DB;
using System.Net;

namespace IntermediateModule02
{
    [Transaction(TransactionMode.Manual)]
    public class Command2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 1a. Filtered Element Collector by view
            View curView = doc.ActiveView;
            FilteredElementCollector collector = new FilteredElementCollector(doc, curView.Id);

            // 1b. ElementMultiCategoryFilter
            List<BuiltInCategory> catList = new List<BuiltInCategory>();
            catList.Add(BuiltInCategory.OST_Areas);
            catList.Add(BuiltInCategory.OST_Walls);
            catList.Add(BuiltInCategory.OST_Doors);
            catList.Add(BuiltInCategory.OST_Furniture);
            catList.Add(BuiltInCategory.OST_LightingFixtures);
            catList.Add(BuiltInCategory.OST_Rooms);
            catList.Add(BuiltInCategory.OST_Windows);

            ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(catList);
            collector.WherePasses(catFilter).WhereElementIsNotElementType();

            // use LINQ to get family symbol by name
            FamilySymbol curDoorTag = GetFamilySymbolByName(doc, "M_Door Tag");
            FamilySymbol curRoomTag = GetFamilySymbolByName(doc, "M_Room Tag");
            FamilySymbol curWallTag = GetFamilySymbolByName(doc, "M_Wall Tag");
            FamilySymbol curCurtainWallTag = GetFamilySymbolByName(doc, "M_Curtain Wall Tag");
            FamilySymbol curFurnitureTag = GetFamilySymbolByName(doc, "M_Furniture Tag");
            FamilySymbol curLightingFixtureTag = GetFamilySymbolByName(doc, "M_Lighting Fixture Tag");
            FamilySymbol curWindowsTag = GetFamilySymbolByName(doc, "M_Window Tag");

            // create dictionary for tags
            Dictionary<string, FamilySymbol> tags = new Dictionary<string, FamilySymbol>();
            tags.Add("Doors", curDoorTag);
            tags.Add("Rooms", curRoomTag);
            tags.Add("Walls", curWallTag);
            tags.Add("Curtain Walls", curCurtainWallTag);
            tags.Add("Furnitures", curFurnitureTag);
            tags.Add("Lighting Fixtures", curLightingFixtureTag);
            tags.Add("Windows", curWindowsTag);

            // Create a list of specific tag names
            List<string> CeilingPlan = new List<string> { "Lighting Fixtures", "Rooms" };
            List<string> FloorPlan = new List<string> { "Curtain Walls", "Doors", "Furnitures", "Rooms",
            "Walls", "Windows"};


            ViewType curViewType = curView.ViewType;


            int counter = 0;
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Insert tags");

                foreach (Element curElem in collector)
                {

                    if (curElem.Location == null)
                        continue;

                    // get insertion point based on element type
                    XYZ insPoint = Utils.GetInsertPoint(curElem.Location);

                    if (insPoint == null)
                        continue;

                    switch (curViewType)
                    {
                        case ViewType.AreaPlan:
                            if (curElem.Category.Name == "Areas")
                            {
                                Area curArea = curElem as Area;

                                AreaTag curAreaTag = doc.Create.NewAreaTag(curView as ViewPlan, curArea, new UV(insPoint.X, insPoint.Y));
                                curAreaTag.TagHeadPosition = new XYZ(insPoint.X, insPoint.Y, 0);
                                curAreaTag.HasLeader = false;
                            }
                                                        counter++;
                            break;


                        case ViewType.CeilingPlan:
                            foreach (string category in CeilingPlan)
                            {
                                if (category == curElem.Category.Name)
                                {
                                    FamilySymbol curTagType = tags[curElem.Category.Name];

                                    // create reference to element
                                    Reference curRef = new Reference(curElem);

                                    // place tag
                                    IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
                                        curRef, false, TagOrientation.Horizontal, insPoint);
                                }
                            }
                            counter++;
                            break;
                        case ViewType.FloorPlan:
                            foreach (string category in FloorPlan)
                            {
                                if (category == curElem.Category.Name)
                                {
                                    FamilySymbol curTagType = tags[curElem.Category.Name];

                                    // create reference to element
                                    Reference curRef = new Reference(curElem);
                                    // place tag
                                    IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
                                        curRef, false, TagOrientation.Horizontal, insPoint);
                                    if (category == "Windows")
                                    {
                                        newTag.TagHeadPosition = insPoint.Add(new XYZ(0, 3, 0));
                                    }
                                }
                            }
                            counter++;
                            break;
                        case ViewType.Section:
                            FamilySymbol curTagSecType = tags["Rooms"];

                            // 4. create reference to element
                            Reference curRefSec = new Reference(curElem);

                            // 5a. place tag
                            IndependentTag newTag1 = IndependentTag.Create(doc, curTagSecType.Id, curView.Id,
                                curRefSec, false, TagOrientation.Horizontal, insPoint);
                            newTag1.TagHeadPosition = insPoint.Add(new XYZ(0, 0, 3));
                            counter++;
                            break;

                    }
                }
                t.Commit();
            }
            TaskDialog.Show("Complete", $"Added {counter} tags to the view");

            return Result.Succeeded;
        }

        //Method faster than the one in Bootcamp
        private static FamilySymbol GetFamilySymbolByName(Document doc, string familyName)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals(familyName))
                .FirstOrDefault();
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData.Data;
        }
    }

}
