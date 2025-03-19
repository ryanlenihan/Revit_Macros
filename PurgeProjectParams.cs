/*
 * Created by SharpDevelop.
 * User: rleni
 * Date: 31/08/2021
 * Time: 10:56 AM
 *
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace TheTestBed
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("AB8DBAA0-A585-49AD-B648-73499E6CCED9")]
    public partial class ThisApplication
    {
        private void Module_Startup(object sender, EventArgs e)
        {
        }

        private void Module_Shutdown(object sender, EventArgs e)
        {
        }

        #region Revit Macros generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Module_Startup);
            this.Shutdown += new System.EventHandler(Module_Shutdown);
        }
        #endregion

        class ProjectParameterData
        {
            public Definition Definition = null;
            public ElementBinding Binding = null;
            public string Name = null;
            public bool IsSharedStatusKnown = false;
            public bool IsShared = false;
            public string GUID = null;
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
            IList<ElementId> elementsToSelect = new List<ElementId>();
            IList<ElementId> paramsToBeDeleted = new List<ElementId>();

            int count = 0;
            int countInterim = 0;
            int countDeleted = 0;
            int countParam = 0;
            int countFams = 0;
            string checkParameter = "";

            foreach (var e in projectParametersData)
            {
                count++;
                countInterim = 0;
                checkParameter = e.Name.ToString();

                foreach (var fi in families)
                {
                    countFams++;
                    if (fi.Category != null)
                    {
                        try
                        {
                            FamilyInstance fiInst = fi as FamilyInstance;
                            Family fiFam = fiInst.Symbol.Family;
                            Parameter p = fi.LookupParameter(checkParameter);
                            string pValue = p.AsString();

                            if (pValue != null)
                            {
                                countFams++;
                                if (!string.IsNullOrEmpty(pValue))
                                {
                                    elementsToSelect.Add(fi.Id);
                                    countInterim++;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            IList<Parameter> systemParams = fi.GetParameters(checkParameter);
                            string pValueSystem = "";
                            countFams++;

                            foreach (var p in systemParams)
                            {
                                pValueSystem = p.AsString();
                            }

                            if (!string.IsNullOrEmpty(pValueSystem))
                            {
                                elementsToSelect.Add(fi.Id);
                                countInterim++;
                            }
                        }
                    }
                }

                if (countInterim == 0)
                {
                    //uncomment the line below if you would like each individual parameter that is not being used to be confirmed via a task dialogue
                    //fair warning - you could be stuck with a lot of task dialogues. suggest generating a list or some other method.
                    //
                    //TaskDialog.Show("title", checkParameter + " is not being used");
                    paramList.Add(checkParameter);
                    countDeleted++;
                }
            }

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

            using (Transaction t = new Transaction(doc, "remove projectparameter"))
            {
                t.Start();
                foreach (ElementId pE in paramsToBeDeleted)
                {
                    doc.Delete(pE);
                    count++;
                }
                t.Commit();
            }

            countParam = elementsToSelect.Distinct().Count();
            if (countParam > 0)
            {
                TaskDialog.Show("result", "There are " + count.ToString() + " project parameters found, with " + countFams.ToString() + " families in the model and these parameters have been used. countParam = " + countParam.ToString() + ". " + countDeleted.ToString() + " parameters have been deleted");
            }
            else
            {
                TaskDialog.Show("result", "There are " + countFams.ToString() + " the parameter has NOT been used. countParam = " + countParam.ToString());
            }
        }
    }
}
