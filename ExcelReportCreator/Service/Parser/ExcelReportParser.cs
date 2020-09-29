using ExcelReportCreatorProject.Domain;
using ExcelReportCreatorProject.LowLevelOperations;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReportCreatorProject.Service
{
    public class ExcelReportParser : IExcelReportParser
    {
        private readonly ParserOptions _parseOptions;

        public ExcelReportParser(ParserOptions parseOptions)
        {
            _parseOptions = parseOptions;
        }

        public IEnumerable<Marker> GetMarkers(IWorkbook workbook)
        {
            var result = Enumerable.Range(0, workbook.NumberOfSheets)
                .SelectMany(sheetIndex => GetMarkers(workbook.GetSheetAt(sheetIndex)));

            return result;
        }

        public IEnumerable<Marker> GetMarkers(ISheet sheet)
        {
            var result = new List<Marker>();

            for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; ++rowIndex)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null) continue;

                for (var cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; ++cellIndex)
                {
                    var cell = row.GetCell(cellIndex);
                    if (cell == null) continue;

                    if (cell.IsMarkedCell(_parseOptions.MarkerOptions))
                    {
                        var markerId = cell.ExtractMarkerValue(_parseOptions.MarkerOptions);
                        var marker = new Marker
                        {
                            Id = markerId,
                            Position = new MarkerPosition
                            {
                                SheetIndex = sheet.Workbook.GetSheetIndex(sheet),
                                RowIndex = rowIndex,
                                CellIndex = cellIndex
                            }
                        };
                        result.Add(marker);
                    }
                }
            }

            return result;
        }
    }
}