using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UnitTestGetParameter
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

            Parameter param1 = element.LookupParameter("Structural Material");
            Parameter param2 = element.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM);

            Parameter param3 = document.GetElement(element.GetTypeId()).LookupParameter("Structural Material");
            Parameter param4 = document.GetElement(element.GetTypeId()).get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM);
            int i = 1;

            return Result.Succeeded;
        }
    }
}