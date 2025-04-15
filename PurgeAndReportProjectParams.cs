using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace TheTestBed
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("AB8DBAA0-A585-49AD-B648-73499E6CCED9")]
    public partial class ThisApplication
    {
        // ... (existing startup/shutdown code)

        class ProjectParameterData
        {
            public Definition Definition = null;
            public ElementBinding Binding = null;
            public string Name = null;
            public bool IsSharedStatusKnown = false;
            public bool IsShared = false;
            public string GUID = null;
        }

        // New helper class for CSV export
        class ParameterUsage
        {
            public string ParameterName { get; set; }
            public ElementId ElementId { get; set; }
            public string Value { get; set; }
        }

        static List<ProjectParameterData> GetProjectParameterData(Document doc)
        {
            if (doc == null)
            {
                throw new ArgumentNullException("doc");
            }
            if (doc.IsFamilyDocument)
            {
                throw new Exception("doc can not be a family document.");
            }

            List<ProjectParameterData> result = new List<ProjectParameterData>();
            BindingMap map = doc.ParameterBindings;
            DefinitionBindingMapIterator it = map.ForwardIterator();
            it.Reset();

            while (it.MoveNext())
            {
                ProjectParameterData newProjectParameterData = new ProjectParameterData
                {
                    Definition = it.Key,
                    Name = it.Key.Name,
                    Binding = it.Current as ElementBinding
                };
                result.Add(newProjectParameterData);
            }

            return result;
        }

        public void CheckParamEmpty()
        {
            UIDocument uidoc = this.ActiveUIDocument;
            Document doc = uidoc.Document;

            var families = new FilteredElementCollector(doc).WhereElementIsNotElementType();
            List<ProjectParameterData> projectParametersData = GetProjectParameterData(doc);
            List<string> paramList = new List<string>();
            IList<ElementId> paramsToBeDeleted = new List<ElementId>();

            // Create a list to store parameter usage details for CSV export.
            List<ParameterUsage> usageList = new List<ParameterUsage>();

            int count = 0;
            int countInterim = 0;
            int countDeleted = 0;
            int countFams = 0;
            string checkParameter = "";

            foreach (var e in projectParametersData)
            {
                count++;
                countInterim = 0;
                checkParameter = e.Name.ToString();

                foreach (var fi in families)
                {
                    if (fi.Category != null)
                    {
                        try
                        {
                            FamilyInstance fiInst = fi as FamilyInstance;
                            // This next line will throw if fi is not a FamilyInstance.
                            Family fiFam = fiInst.Symbol.Family; 
                            Parameter p = fi.LookupParameter(checkParameter);
                            if (p != null)
                            {
                                string pValue = p.AsString();
                                if (!string.IsNullOrEmpty(pValue))
                                {
                                    // Record each usage in the list.
                                    usageList.Add(new ParameterUsage { ParameterName = checkParameter, ElementId = fi.Id, Value = pValue });
                                    countInterim++;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            IList<Parameter> systemParams = fi.GetParameters(checkParameter);
                            string pValueSystem = "";
                            foreach (var p in systemParams)
                            {
                                pValueSystem = p.AsString();
                            }
                            if (!string.IsNullOrEmpty(pValueSystem))
                            {
                                usageList.Add(new ParameterUsage { ParameterName = checkParameter, ElementId = fi.Id, Value = pValueSystem });
                                countInterim++;
                            }
                        }
                        countFams++;
                    }
                }

                if (countInterim == 0)
                {
                    // Parameter unused, mark it to be deleted.
                    paramList.Add(checkParameter);
                    countDeleted++;
                }
            }

            // Export used parameter details to a CSV file on the desktop.
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = Path.Combine(desktopPath, "ParameterUsageReport.csv");

                using (StreamWriter writer = new StreamWriter(filePath, false))
                {
                    writer.WriteLine("ParameterName,ElementId,Value");
                    foreach (var record in usageList)
                    {
                        // Wrap the value in quotes if it contains a comma.
                        string valueFormatted = record.Value.Contains(",") ? "\"" + record.Value + "\"" : record.Value;
                        writer.WriteLine($"{record.ParameterName},{record.ElementId.IntegerValue},{valueFormatted}");
                    }
                }
                TaskDialog.Show("CSV Export", "CSV file exported to:\n" + filePath);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("CSV Export Error", "An error occurred while exporting CSV: " + ex.Message);
            }

            // Existing code to collect unused parameters for deletion.
            IEnumerable<ParameterElement> _params = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(ParameterElement))
                .Cast<ParameterElement>();

            foreach (var parametername in paramList)
            {
                foreach (ParameterElement pElem in _params)
                {
                    if (pElem.GetDefinition().Name == parametername)
                    {
                        paramsToBeDeleted.Add(pElem.Id);
                    }
                }
            }

            using (Transaction t = new Transaction(doc, "Remove project parameter"))
            {
                t.Start();
                foreach (ElementId pE in paramsToBeDeleted)
                {
                    doc.Delete(pE);
                    count++;
                }
                t.Commit();
            }

            int countParam = usageList.Select(record => record.ElementId).Distinct().Count();
            if (countParam > 0)
            {
                TaskDialog.Show("Result",
                    "There are " + count.ToString() + " project parameters found, with " +
                    countFams.ToString() + " family checks. Unique elements with filled parameters = " +
                    countParam.ToString() + ". " + countDeleted.ToString() + " parameters have been marked for deletion.");
            }
            else
            {
                TaskDialog.Show("Result",
                    "There are " + countFams.ToString() + " families; the parameter has not been used. Unique element count = " +
                    countParam.ToString());
            }
        }
    }
}
