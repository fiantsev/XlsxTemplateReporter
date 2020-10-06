using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Domain.Markers;
using ClosedXML.Excel;

namespace TemplateCooker.Service.ResourceInjection
{
    public class InjectionContext
    {
        public MarkerRange MarkerRange { get; set; }
        public ResourceObject ResourceObject { get; set; }
        public IXLWorkbook Workbook { get; set; }
    }
}