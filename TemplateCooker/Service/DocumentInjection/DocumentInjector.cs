using System.Linq;
using ClosedXML.Excel;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Creation;
using TemplateCooker.Service.Extraction;
using TemplateCooker.Service.ResourceInjection;
using TemplateCooker.Service.ResourceObjectProvision;

namespace TemplateCooker
{
    public class DocumentInjector : IDocumentInjector
    {
        private readonly IResourceInjector _resourceInjector;
        private readonly IResourceObjectProvider _resourceObjectProvider;
        private readonly MarkerOptions _markerOptions;

        public DocumentInjector(DocumentInjectorOptions options)
        {
            _resourceInjector = options.ResourceInjector;
            _resourceObjectProvider = options.ResourceObjectProvider;
            _markerOptions = options.MarkerOptions;
        }

        public void Inject(IXLWorkbook workbook)
        {

            foreach (var sheetIndex in Enumerable.Range(1, workbook.Worksheets.Count))
            {
                var sheet = workbook.Worksheet(sheetIndex);
                var markerExtractor = new MarkerExtractor(sheet, _markerOptions);
                var markers = markerExtractor.GetMarkers();
                var markerRegions = new MarkerRangeCollection(markers);

                foreach (var markerRegion in markerRegions)
                    InjectResourceToSheet(sheet, markerRegion);
            }
        }

        private void InjectResourceToSheet(IXLWorksheet sheet, MarkerRange markerRegion)
        {
            var resourceObject = _resourceObjectProvider.Resolve(markerRegion.StartMarker.Id);
            var injectionContext = new InjectionContext
            {
                MarkerRange = markerRegion,
                Workbook = sheet.Workbook,
                ResourceObject = resourceObject,
            };

            _resourceInjector.Inject(injectionContext);
        }
    }
}