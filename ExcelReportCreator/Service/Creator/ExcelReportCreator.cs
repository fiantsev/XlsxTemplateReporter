using System.Linq;
using ClosedXML.Excel;
using ExcelReportCreatorProject.Domain.Markers;
using ExcelReportCreatorProject.Domain.Markers.ExtractorOptions;
using ExcelReportCreatorProject.Service.Creator;
using ExcelReportCreatorProject.Service.Injection;
using ExcelReportCreatorProject.Service.ResourceObjectProvider;

namespace ExcelReportCreatorProject
{
    public class ExcelReportCreator : IExcelReportCreator
    {
        private readonly IResourceInjector _resourceInjector;
        private readonly IResourceObjectProvider _resourceObjectProvider;
        private readonly MarkerExtractionOptions _markerExtractionOptions;
        private readonly FormulaEvaluationOptions _formulaEvaluationOptions;

        public ExcelReportCreator(ExcelReportCreatorOptions options)
        {
            _resourceInjector = options.ResourceInjector;
            _resourceObjectProvider = options.ResourceObjectProvider;
            _markerExtractionOptions = options.MarkerExtractionOptions;
            _formulaEvaluationOptions = options.FormulaEvaluationOptions;
        }

        public IXLWorkbook Create(IXLWorkbook workbook)
        {
            foreach(var sheetIndex in Enumerable.Range(1, workbook.Worksheets.Count))
            {
                var sheet = workbook.Worksheet(sheetIndex);

                var markerCollection = new MarkerCollection(sheet, _markerExtractionOptions);
                var markerRegions = new MarkerRegionCollection(markerCollection);

                foreach(var markerRegion in markerRegions)
                {
                    InjectResourceToSheet(sheet, markerRegion);
                }
            }

            if (_formulaEvaluationOptions.EvaluateFormulas)
                EvaluateFormulas(workbook);

            return workbook;
        }

        private void InjectResourceToSheet(IXLWorksheet sheet, MarkerRegion markerRegion)
        {
            var resourceObject = _resourceObjectProvider.Resolve(markerRegion.StartMarker.Id);
            var injectionContext = new InjectionContext
            {
                MarkerRegion = markerRegion,
                Workbook = sheet.Workbook,
                ResourceObject = resourceObject,
            };

            _resourceInjector.Inject(injectionContext);
        }

        private void EvaluateFormulas(IXLWorkbook workbook)
        {
            //var formulaEvaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            //foreach (var sheetIndex in Enumerable.Range(0, workbook.NumberOfSheets))
            //{
            //    var sheet = workbook.GetSheetAt(sheetIndex);

            //    for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; ++rowIndex)
            //    {
            //        var row = sheet.GetRow(rowIndex);
            //        if (row == null) continue;

            //        for (var cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; ++cellIndex)
            //        {
            //            var cell = row.GetCell(cellIndex);
            //            if (cell == null) continue;

            //            formulaEvaluator.EvaluateFormulaCell(cell);
            //        }
            //    }
            //}
        }

    }
}