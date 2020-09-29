using ExcelReportCreatorProject.Domain.Markers;
using System.Collections.Generic;

namespace ExcelReportCreatorProject.Service.MarkerExtraction
{
    public interface IMarkerMatcher : IEnumerable<MarkerRegion>
    {
    }
}
