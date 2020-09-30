using ExcelReportCreatorProject.Domain;
using ExcelReportCreatorProject.Extensions;
using NPOI.SS.UserModel;
using System.Collections;
using System.Collections.Generic;

namespace ExcelReportCreatorProject.Service.MarkerExtraction
{
    public class MarkerExtractor : IMarkerExtractor
    {
        private readonly List<ISheet> _sheets;
        private readonly MarkerExtractorOptions _markerExtractorOptions;

        public MarkerExtractor(ISheet sheet, MarkerExtractorOptions markerExtractorOptions)
        {
            _sheets = new List<ISheet> { sheet };
            _markerExtractorOptions = markerExtractorOptions;
        }

        public MarkerExtractor(IWorkbook workbook, MarkerExtractorOptions markerExtractorOptions)
        {
            _sheets = new List<ISheet>(workbook.EnumerateSheets());
            _markerExtractorOptions = markerExtractorOptions;
        }

        public IEnumerator<Marker> GetEnumerator()
        {
            var markerOptions = _markerExtractorOptions.MarkerOptions;

            foreach(var sheet in _sheets)
            {
                for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; ++rowIndex)
                {
                    var row = sheet.GetRow(rowIndex);
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
                                    SheetIndex = sheet.Workbook.GetSheetIndex(sheet),
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
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}