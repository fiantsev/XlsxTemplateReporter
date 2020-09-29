using ExcelReportCreatorProject.Domain;
using ExcelReportCreatorProject.Domain.Data;
using ExcelReportCreatorProject.Service;
using ExcelReportCreatorProject.Service.Creator;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReportCreatorProject
{
    public class ExcelReportCreator : IExcelReportCreator
    {
        private readonly IExcelReportParser _excelReportParser;
        private readonly IResourceInjector _resourceInjector;

        public ExcelReportCreator(IExcelReportParser excelReportParser, IResourceInjector resourceInjector, IResourceObjectProvider resourceObjectProvider)
        {
            _excelReportParser = excelReportParser;
            _resourceInjector = resourceInjector;
            _resourceObjectProvider = resourceObjectProvider;
        }

        public void Create(XSSFWorkbook workbook)
        {
            foreach(var sheetIndex in Enumerable.Range(0, workbook.NumberOfSheets))
            {
                var sheet = workbook.GetSheetAt(sheetIndex);
                IEnumerable<Marker> markers;
                //переделать здесь на бесконечный IEnumerator
                while ((markers = _excelReportParser.GetMarkers(sheet)).Count() != 0)
                {
                    var firstMarker = markers.First();

                }
            }
        }

        private void InjectResourceToSheet(ISheet sheet, Marker marker)
        {
            var injectionContext = new InjectionContext
            {
                Marker = marker,
                Workbook = sheet.Workbook,
                ResourceObject = 
            };

            _resourceInjector.Inject(injectionContext);
        }

    }
}