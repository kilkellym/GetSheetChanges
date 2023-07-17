#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace GetChangeViewsAndSheets
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // prompt user to select elements
            TaskDialog td1 = new TaskDialog("Calculate Sheets");
            td1.MainInstruction = "Calculate SHEETS that will need to be reissued.";
            td1.MainContent = "Please note: Drafting Views and Schedules are note part of the calculation.";
            td1.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Select Elements");
            td1.CommonButtons = TaskDialogCommonButtons.Close;
            td1.DefaultButton = TaskDialogResult.Close;

            TaskDialogResult result1 = td1.Show();

            if (TaskDialogResult.CommandLink2 == result1)
            {
                // create list of view types to check
                List<ViewType> viewTypes = new List<ViewType> { ViewType.FloorPlan, ViewType.CeilingPlan,
                ViewType.Elevation, ViewType.Section, ViewType.Detail,
                ViewType.Schedule, ViewType.AreaPlan};

                // prompt user to select elements to check
                List<Element> selectedElems = Utils.SelectElements(uiapp, "Select elements to include in the calculation:");

                // get list of all views in the model
                List<View> viewList = Utils.GetAllViews(doc);

                // create list of sheet names and numbers
                List<string> ListOfSheets = new List<string>();

                // start stopwatch to time process
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                // loop through selected elements and check if they're in any of the views
                foreach (Element curElem in selectedElems)
                {
                    // check to make sure we're only looping through model elements
                    if (curElem.Category != null && curElem.Category.CategoryType == CategoryType.Model)
                    {
                        // loop through all the selected views in the model
                        foreach (View curView in viewList)
                        {
                            // check if the current view is a selected view type
                            if (viewTypes.Contains(curView.ViewType))
                            {
                                // get the current view's sheet number and sheet name, if it has them
                                Parameter SheetNumParam = curView.LookupParameter("Sheet Number");
                                Parameter SheetNameParam = curView.LookupParameter("Sheet Name");
                            
                                // if the view has a sheet number and name, create a FEC to see if the current 
                                // element is visible in that view
                                if(SheetNumParam != null && SheetNameParam != null)
                                {
                                    string SheetNumAndName = SheetNumParam.AsString() + " - " + SheetNameParam.AsString();

                                    // create an element parameter filter to filter the FEC by element Id
                                    ParameterValueProvider provider = new ParameterValueProvider(new ElementId(BuiltInParameter.ID_PARAM));
                                    FilterNumericEquals evaluator = new FilterNumericEquals();
                                    FilterElementIdRule rule = new FilterElementIdRule(provider, evaluator, curElem.Id);
                                    ElementParameterFilter filter = new ElementParameterFilter(rule);
                            
                                    // create a FEC that filters by the element's category and applies the element param filter
                                    FilteredElementCollector elemCollector = new FilteredElementCollector(doc, curView.Id)
                                        .OfCategoryId(curElem.Category.Id)
                                        .WhereElementIsNotElementType().WherePasses(filter);

                                    // check FEC results - if greater than 0, element is visible in the view
                                    if (elemCollector.GetElementCount() > 0)
                                        ListOfSheets.Add(SheetNumAndName);
                                }
                            }
                        }
                    }
                }

                // stop stopwatch and get elapsed time
                stopwatch.Stop();
                TimeSpan ts = stopwatch.Elapsed;
                string elapsedTime = ts.ToString(@"m\:ss");

                // get list of unique sheets from sheet list
                List<string> uniqueSheets = ListOfSheets.Distinct().ToList();

                string resultString = "";
                if (uniqueSheets.Count > 0)
                    resultString = string.Join("\n", uniqueSheets);
                else
                    resultString = "No sheets to reissue.";

                // show user results
                TaskDialog td2 = new TaskDialog("Results");
                td2.MainInstruction = resultString;
                td2.MainContent = "Time taken = " + elapsedTime;
                td2.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Save list to .csv");
                td2.CommonButtons = TaskDialogCommonButtons.Close;
                td2.DefaultButton = TaskDialogResult.Close;

                TaskDialogResult result2 = td2.Show();

                if (TaskDialogResult.CommandLink2 == result2)
                {
                    if (uniqueSheets.Count > 0)
                    {
                        string SaveFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        string CurDate = DateTime.Now.ToString("yyyyMMdd_hhmm");
                        string FileName = "SheetList_" + CurDate + ".csv";
                        string SaveFile = SaveFilePath + FileName;

                        // create list for CSV export
                        List<string> resultsList = new List<string>();
                        resultsList.Add("File Name:," + doc.PathName);
                        resultsList.Add("Date:," + DateTime.Now.ToString("yyyy-MM-dd"));
                        resultsList.Add("Sheet Number-Name List:");

                        // add list of sheets to result list
                        resultsList.AddRange(uniqueSheets);

                        // write results to CSV file
                        Utils.WriteListToTxtFile(SaveFile, resultsList);

                        // open CSV file
                        Process.Start(SaveFile);
                    }
                    else
                    {
                        TaskDialog.Show("Complete", "No sheets to reissue.");
                    }
                }

            }

            return Result.Succeeded;
        }

        
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
