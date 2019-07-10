using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GroupUnitTest
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        private Document _document;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document document = commandData.Application.ActiveUIDocument.Document;
            _document = document;
            UIDocument uIDocument = commandData.Application.ActiveUIDocument;
            Element element = SelecttionSchedule(uIDocument);
            if (element is ViewSchedule viewSchedule) {
                var elementsIdCollector = new FilteredElementCollector(document, viewSchedule.Id).ToElementIds();
                using (Transaction t = new Transaction(document, "ok")) {
                    t.Start();
                    foreach (ElementId elementId in elementsIdCollector) {
                        Element ele = document.GetElement(elementId);
                        if (ele == null) {
                            continue;
                        }
                        Group gr = (Group)document.GetElement(ele.GroupId);
                        if (gr != null) {
                            UpdateGroupProcessing updateGroupProcessing = new UpdateGroupProcessing(ele);
                            updateGroupProcessing.UnGroup();
                            Parameter parameter = ele.LookupParameter("Comments");
                            parameter.Set("changed");
                            updateGroupProcessing.ReGroup();
                        }
                        else {
                            Parameter parameter = ele.LookupParameter("Comments");
                            parameter.Set("changed");
                        }
                    }
                    t.Commit();
                }
            }

            return Result.Succeeded;
        }

        public static Element SelecttionSchedule(UIDocument uIDocument)
        {
            var ids = uIDocument.Selection.GetElementIds().ToList();

            return uIDocument.Document.GetElement(ids[0]);
        }
    }
}