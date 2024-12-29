namespace IntermediateModule02.Common
{
    internal static class Utils
    {
        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel;

            if (GetRibbonPanelByName(app, tabName, panelName) == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);

            else
                curPanel = GetRibbonPanelByName(app, tabName, panelName);

            return curPanel;
        }
        public static XYZ GetMidpointBetweenTwoPoints(XYZ point1, XYZ point2)
        {
            XYZ midPoint = new XYZ(
                (point1.X + point2.X) / 2,
                (point1.Y + point2.Y) / 2,
                (point1.Z + point2.Z) / 2);

            return midPoint;
        }
        public static XYZ GetInsertPoint(Location loc)
        {
            LocationPoint locPoint = loc as LocationPoint;
            XYZ point;

            if (locPoint != null)
            {
                point = locPoint.Point;
            }
            else
            {
                LocationCurve locCurve = loc as LocationCurve;
                point = GetMidpointBetweenTwoPoints(locCurve.Curve.GetEndPoint(0), locCurve.Curve.GetEndPoint(1));

            }

            return point;
        }
        public static FamilySymbol GetFamilySymbolByName(Document doc, string familyName, string familySymbolName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));

            foreach (FamilySymbol curSymbol in collector)
            {
                if (curSymbol.FamilyName == familyName)
                {
                    if (curSymbol.Name == familySymbolName)
                    {
                        return curSymbol;
                    }
                }
            }

            return null;
        }


        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }
    }
}
