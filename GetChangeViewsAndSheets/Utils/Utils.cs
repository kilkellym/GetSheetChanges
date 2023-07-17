using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetChangeViewsAndSheets
{
    internal static class Utils
    {

        public static void WriteListToTxtFile(string filePath, List<string> fileContents)
        {
            using (StreamWriter writer = File.AppendText(filePath))
            {
                foreach(string line in fileContents)
                {
                    writer.WriteLine(line);
                }
            }
        }
        public static List<Element> SelectElements(UIApplication curUIA, string prompt)
        {
            IList<Reference> curRefs = curUIA.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, prompt);

            //loop through references and add to element list
            List<Element> curElems = new List<Element>();

            foreach (Reference tmp in curRefs)
            {
                curElems.Add(curUIA.ActiveUIDocument.Document.GetElement(tmp.ElementId));
            }
            return curElems;
        }
        public static List<View> GetAllViews(Document curDoc)
        {
            FilteredElementCollector m_colviews = new FilteredElementCollector(curDoc);
            m_colviews.OfCategory(BuiltInCategory.OST_Views);

            List<View> m_views = new List<View>();
            foreach (View x in m_colviews.ToElements())
            {
                if(x.IsTemplate == false)
                    m_views.Add(x);
            }

            return m_views;
        }
        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel currentPanel = GetRibbonPanelByName(app, tabName, panelName);

            if (currentPanel == null)
                currentPanel = app.CreateRibbonPanel(tabName, panelName);

            return currentPanel;
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
