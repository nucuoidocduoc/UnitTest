using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

namespace UnitTest
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class CommandL : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            WorksetTable worksetTable = doc.GetWorksetTable();

            var elementIds = SelecttionSchedule(uidoc);

            Workset workset = worksetTable.GetWorkset(doc.GetElement(elementIds[0]).WorksetId);
            MessageBox.Show(workset.IsEditable.ToString());

            return Result.Succeeded;
        }

        public static List<ElementId> SelecttionSchedule(UIDocument uIDocument)
        {
            return uIDocument.Selection.GetElementIds().ToList();
        }
    }
}