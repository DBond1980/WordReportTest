using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReportTest.Export
{
    public class FieldItem
    {
        private int _id1; //CycleGroupId
        private int _id2; //ResultId
        private int _id3; //SamplesId

        private Field _field;

        public string Content { get; set; } = "";

        public string AttrDig =>
            "{" 
            + (_id1 < 0 ? "" : _id1.ToString()) + (_id2 < 0 ? "" : ".")
            + (_id2 < 0 ? "" : _id2.ToString()) + (_id3 < 0 ? "" : ".")
            + (_id3 < 0 ? "" : _id3.ToString()) + (_id1 < 0 ? "" : "/")
            + _field.NameDig + "}";

        public string AttrText =>
            "{"
            + (_id1 < 0 ? "" : _id1.ToString()) + (_id2 < 0 ? "" : ".")
            + (_id2 < 0 ? "" : _id2.ToString()) + (_id3 < 0 ? "" : ".")
            + (_id3 < 0 ? "" : _id3.ToString()) + (_id1 < 0 ? "" : "/")
            + _field.Name + "}";

        public FieldItem(Field field, string content, int id1 = -1, int id2 = -1, int id3 = -1)
        {
            _field = field;
            Content = content;
            _id1 = id1;
            _id2 = id2;
            _id3 = id3;
        }

    }
}
