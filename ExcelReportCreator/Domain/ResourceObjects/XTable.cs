using System.Collections.Generic;

namespace ExcelReportCreatorProject.Domain.ResourceObjects
{
    public class XTable
    {
        public List<List<string>> Colums { get; set; }
        public List<List<string>> Rows { get; set; }
        public List<List<string>> Values { get; set; }
    }
}