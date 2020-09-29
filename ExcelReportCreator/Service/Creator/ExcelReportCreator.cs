using ExcelReportCreatorProject.Domain.Markers;
using ExcelReportCreatorProject.Domain.ResourceObjects;
using ExcelReportCreatorProject.Service.Creator;
using ExcelReportCreatorProject.Service.MarkerExtraction;
using ExcelReportCreatorProject.Service.ResourceObjectProvider;
using NPOI.SS.UserModel;
using System.Linq;

namespace ExcelReportCreatorProject
{
    public class ExcelReportCreator : IExcelReportCreator
    {
        private readonly IResourceInjector _resourceInjector;
        private readonly IResourceObjectProvider _resourceObjectProvider;
        private readonly MarkerExtractorOptions _markerExtractorOptions;

        public ExcelReportCreator(ExcelReportCreatorOptions options)
        {
            _resourceInjector = options.ResourceInjector;
            _resourceObjectProvider = options.ResourceObjectProvider;
            _markerExtractorOptions = options.MarkerExtractorOptions;
        }

        public void Create(IWorkbook workbook)
        {
            foreach(var sheetIndex in Enumerable.Range(0, workbook.NumberOfSheets))
            {
                var sheet = workbook.GetSheetAt(sheetIndex);

                var markerExtractor = new MarkerExtractor(sheet, _markerExtractorOptions);
                var markerRegions = new MarkerRegionCollection(markerExtractor);

                foreach(var markerRegion in markerRegions)
                {
                    InjectResourceToSheet(sheet, markerRegion);
                }
            }
        }

        private void InjectResourceToSheet(ISheet sheet, MarkerRegion markerRegion)
        {
            var injectionContext = new InjectionContext
            {
                MarkerRegion = markerRegion,
                Workbook = sheet.Workbook,
                ResourceObject = new ResourceObject(),
            };

            _resourceInjector.Inject(injectionContext);
        }

    }
}