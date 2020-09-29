using ExcelReportCreatorProject.Domain;
using ExcelReportCreatorProject.LowLevelOperations;
using NPOI.SS.UserModel;
using System.Collections;
using System.Collections.Generic;

namespace ExcelReportCreatorProject.Service.MarkerExtraction
{
    public class MarkerExtractor : IMarkerExtractor
    {
        private readonly ISheet _sheet;
        private readonly MarkerExtractorOptions _markerExtractorOptions;

        public MarkerExtractor(ISheet sheet, MarkerExtractorOptions markerExtractorOptions)
        {
            _sheet = sheet;
            _markerExtractorOptions = markerExtractorOptions;
        }

        public IEnumerator<Marker> GetEnumerator()
        {
            var markerOptions = _markerExtractorOptions.MarkerOptions;
            for (var rowIndex = _sheet.FirstRowNum; rowIndex <= _sheet.LastRowNum; ++rowIndex)
            {
                var row = _sheet.GetRow(rowIndex);
                if (row == null) continue;

                for (var cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; ++cellIndex)
                {
                    var cell = row.GetCell(cellIndex);
                    if (cell == null) continue;

                    if (cell.IsMarkedCell(markerOptions))
                    {
                        var markerId = cell.ExtractMarkerValue(markerOptions);
                        var isEndMarker = markerId.Substring(0, markerOptions.Terminator.Length) == markerOptions.Terminator;
                        var marker = new Marker
                        {
                            Id = isEndMarker ? markerId.Substring(markerOptions.Terminator.Length) : markerId,
                            Position = new MarkerPosition
                            {
                                SheetIndex = _sheet.Workbook.GetSheetIndex(_sheet),
                                RowIndex = rowIndex,
                                CellIndex = cellIndex
                            },
                            MarkerType = isEndMarker ? MarkerType.End : MarkerType.Start
                        };
                        yield return marker;
                    }
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}