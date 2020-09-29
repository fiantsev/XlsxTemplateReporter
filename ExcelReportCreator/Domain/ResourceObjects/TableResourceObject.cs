using System.Data;

namespace ExcelReportCreatorProject.Domain.ResourceObjects
{
    public class TableResourceObject : ResourceObject
    {
        public DataTable Table { get; set; }
    }
}