using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReportTest.Export
{
    public class ExportFieldItem
    {
        public int Id1 { get; set; } //CycleGroupId
        public int Id2 { get; set; } //ResultId
        public int Id3 { get; set; } //SamplesId

        private ExportField _exportField;

        public string Content { get; set; }

        private string _userFieldName;
        private string _back;

        public string AttrDig =>
            "{" 
            + (Id1 < 0 ? "" : Id1.ToString()) + (Id2 < 0 ? "" : ".")
            + (Id2 < 0 ? "" : Id2.ToString()) + (Id3 < 0 ? "" : ".")
            + (Id3 < 0 ? "" : Id3.ToString()) + (Id1 < 0 ? "" : "/")
            + _exportField.NameDig + "}";

        public string AttrText =>
            "{"
            + (Id1 < 0 ? "" : Id1.ToString()) + (Id2 < 0 ? "" : ".")
            + (Id2 < 0 ? "" : Id2.ToString()) + (Id3 < 0 ? "" : ".")
            + (Id3 < 0 ? "" : Id3.ToString()) + (Id1 < 0 ? "" : "/")
            + _exportField.Name + "}";

        public bool IsUserField => _userFieldName != null;
        public string AttrUserField => "{" + _userFieldName + "}";

        public bool HasBackground => _back != null;
        public string Background => _back;

        public ExportFieldItem(ExportField exportField, string content, int id1 = -1, int id2 = -1, int id3 = -1,
            string userFieldName = null, string back = null)
        {
            _exportField = exportField;
            Content = content;
            Id1 = id1;
            Id2 = id2;
            Id3 = id3;
            _userFieldName = userFieldName;
            _back = back;
        }

    }
}
