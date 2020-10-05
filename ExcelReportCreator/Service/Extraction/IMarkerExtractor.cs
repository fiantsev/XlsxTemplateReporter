using System.Collections.Generic;
using ExcelReportCreatorProject.Domain.Markers;

namespace ExcelReportCreatorProject.Service.Extraction
{
    public interface IMarkerExtractor
    {
        IEnumerable<Marker> Markers();
    }
}