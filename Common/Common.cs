using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    public static class Common
    {
        public static Element SelecttionSchedule(UIDocument uIDocument)
        {
            var ids = uIDocument.Selection.GetElementIds().ToList();

            return uIDocument.Document.GetElement(ids[0]);
        }
    }
}