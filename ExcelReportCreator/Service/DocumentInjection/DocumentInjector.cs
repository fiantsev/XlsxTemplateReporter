﻿using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelReportCreatorProject.Domain.Markers;
using ExcelReportCreatorProject.Service.Creation;
using ExcelReportCreatorProject.Service.Extraction;
using ExcelReportCreatorProject.Service.ResourceInjection;
using ExcelReportCreatorProject.Service.ResourceObjectProvision;

namespace ExcelReportCreatorProject
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
            var markerExtractor = new MarkerExtractor(workbook, _markerOptions);

            foreach (var sheetIndex in Enumerable.Range(1, workbook.Worksheets.Count))
            {
                var sheet = workbook.Worksheet(sheetIndex);
                var markers = markerExtractor.Markers();
                var markerRegions = new MarkerRangeCollection(markers);

                foreach(var markerRegion in markerRegions)
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