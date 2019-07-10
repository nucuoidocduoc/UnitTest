using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestSetParameter
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document document = commandData.Application.ActiveUIDocument.Document;
            UIDocument uIDocument = commandData.Application.ActiveUIDocument;

            var refe = uIDocument.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select element");

            if (refe == null) {
                return Result.Failed;
            }

            Element element = document.GetElement(refe.ElementId);

            Parameter baseLevel = element.LookupParameter("Base Level");

            Parameter topLevel = element.LookupParameter("Top Level");

            using (Transaction t = new Transaction(document, "set level")) {
                t.Start();
                baseLevel.SetValueString("Level 2");
                topLevel.SetValueString("Level 3");
                t.Commit();
            }

            return Result.Succeeded;
        }
    }
}