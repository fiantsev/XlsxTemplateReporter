using System.Collections.Generic;

namespace XlsxTemplateReporter
{
    public class WidgetData
    {
        public List<List<string>> Cols { get; set; }
        public List<List<string>> Rows { get; set; }
        public List<List<object>> Values { get; set; }
    }
}
