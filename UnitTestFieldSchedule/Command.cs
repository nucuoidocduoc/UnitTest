using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UnitTestFieldSchedule
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            int index = 1;
            Document document = commandData.Application.ActiveUIDocument.Document;
            UIDocument uIDocument = commandData.Application.ActiveUIDocument;

            Element element = SelecttionSchedule(uIDocument);
            if (element is ViewSchedule viewSchedule) {
                var collec = new FilteredElementCollector(document, viewSchedule.Id).ToList();
                MessageBox.Show(collec.Count.ToString());
            }

            //var viewScehdule = new FilteredElementCollector(document).OfClass(typeof(ViewSchedule)).ToList();
            //string messageInfo = string.Empty;
            //foreach (ViewSchedule viewSchedule in viewScehdule) {
            //    if (index != 1) {
            //        messageInfo += "\n";
            //    }

            //    messageInfo += index.ToString() + ". " + viewSchedule.Name + " - BuiltIn: " + ((BuiltInCategory)viewSchedule.Definition.CategoryId.IntegerValue).ToString() + "\n";
            //    foreach (ScheduleFieldId id in viewSchedule.Definition.GetFieldOrder()) {
            //        ScheduleField scheduleField = viewSchedule.Definition.GetField(id);
            //        messageInfo += scheduleField.GetName() + " : " + "BuiltInParameter." + ((BuiltInParameter)scheduleField.ParameterId.IntegerValue).ToString() + " ------- ";
            //    }
            //    index++;
            //}
            //System.IO.File.WriteAllText(@"D:\WriteText.txt", messageInfo);
            return Result.Succeeded;
        }

        public static Element SelecttionSchedule(UIDocument uIDocument)
        {
            var ids = uIDocument.Selection.GetElementIds().ToList();

            return uIDocument.Document.GetElement(ids[0]);
        }
    }
}