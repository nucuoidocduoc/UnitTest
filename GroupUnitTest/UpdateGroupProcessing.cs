using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupUnitTest
{
    public class UpdateGroupProcessing
    {
        private static int _countInstanceUpdateGroup = 0;
        private Document _document;
        private Group _group;
        private GroupType _groupType;
        private string _groupName;
        private GroupSet _groupSet;
        private const string _bufferGroupName = "GroupNameAlternative";
        private Element _element;
        private ICollection<ElementId> _elementIds;

        public UpdateGroupProcessing(Element element)
        {
            _document = element.Document;
            _element = element;
            _group = (Group)_document.GetElement(element.GroupId);
            _groupType = _group.GroupType;
            _groupSet = _groupType.Groups;
            _groupName = _groupType.Name;
            _countInstanceUpdateGroup++;
        }

        public bool UnGroup()
        {
            _groupType.Name = _bufferGroupName + _countInstanceUpdateGroup.ToString();
            _elementIds = _group.UngroupMembers();
            return true;
        }

        public bool ReGroup()
        {
            var newGr = _document.Create.NewGroup(_elementIds);
            newGr.GroupType.Name = _groupName;
            if (_groupSet.Size > 1) {
                foreach (Group group in _groupSet) {
                    if (group.Id.IntegerValue != _group.Id.IntegerValue) {
                        group.GroupType = newGr.GroupType;
                    }
                }
            }
            _document.Delete(_groupType.Id);
            return true;
        }

        public static bool IsBelongGroup(Element element)
        {
            return element.Document.GetElement(element.GroupId) != null;
        }
    }
}