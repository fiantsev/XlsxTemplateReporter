using ExcelReportCreatorProject.Domain.Markers;
using ExcelReportCreatorProject.Service.Creator;
using ExcelReportCreatorProject.Service.Injection;
using ExcelReportCreatorProject.Service.MarkerExtraction;
using ExcelReportCreatorProject.Service.ResourceObjectProvider;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using System.Linq;

namespace ExcelReportCreatorProject
{
    public class ExcelReportCreator : IExcelReportCreator
    {
        private readonly IResourceInjector _resourceInjector;
        private readonly IResourceObjectProvider _resourceObjectProvider;
        private readonly MarkerExtractorOptions _markerExtractorOptions;
        private readonly FormulaEvaluationOptions _formulaEvaluationOptions;

        public ExcelReportCreator(ExcelReportCreatorOptions options)
        {
            _resourceInjector = options.ResourceInjector;
            _resourceObjectProvider = options.ResourceObjectProvider;
            _markerExtractorOptions = options.MarkerExtractorOptions;
            _formulaEvaluationOptions = options.FormulaEvaluationOptions;
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

            if (_formulaEvaluationOptions.EvaluateFormulas)
                EvaluateFormulas(workbook);
        }

        private void InjectResourceToSheet(ISheet sheet, MarkerRegion markerRegion)
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

        private void EvaluateFormulas(IWorkbook workbook)
        {
            var formulaEvaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            foreach (var sheetIndex in Enumerable.Range(0, workbook.NumberOfSheets))
            {
                var sheet = workbook.GetSheetAt(sheetIndex);

                for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; ++rowIndex)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row == null) continue;

                    for (var cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; ++cellIndex)
                    {
                        var cell = row.GetCell(cellIndex);
                        if (cell == null) continue;

                        formulaEvaluator.EvaluateFormulaCell(cell);
                    }
                }
            }
        }

    }
}