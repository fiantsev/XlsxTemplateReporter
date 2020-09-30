using System.Collections.Generic;

namespace ExcelReportCreatorProject.Domain.ResourceObjects
{
    public class TableResourceObject : ResourceObject
    {
        public List<List<object>> Table { get; set; }
    }
}