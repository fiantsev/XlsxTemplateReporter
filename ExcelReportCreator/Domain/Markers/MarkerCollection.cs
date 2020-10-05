using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelReportCreatorProject.Domain.Markers.ExtractorOptions;
using ExcelReportCreatorProject.Domain.Npoi;
using ExcelReportCreatorProject.Extensions;

namespace ExcelReportCreatorProject.Domain.Markers
{
    public class MarkerCollection : IEnumerable<Marker>
    {
        private readonly List<IXLWorksheet> _sheets;
        private readonly MarkerExtractionOptions _markerExtractionOptions;

        public MarkerCollection(IXLWorksheet sheet, MarkerExtractionOptions markerExtractorOptions)
        {
            _sheets = new List<IXLWorksheet> { sheet };
            _markerExtractionOptions = markerExtractorOptions;
        }

        public MarkerCollection(XLWorkbook workbook, MarkerExtractionOptions markerExtractorOptions)
        {
            _sheets = new List<IXLWorksheet>(new SheetCollection(workbook));
            _markerExtractionOptions = markerExtractorOptions;
        }

        public IEnumerator<Marker> GetEnumerator()
        {
            var markerOptions = _markerExtractionOptions.MarkerOptions;

            foreach (var sheet in _sheets)
            {
                for (var rowIndex = sheet.FirstRowUsed().FirstCellUsed().Address.RowNumber; rowIndex <= sheet.LastRowUsed().FirstCellUsed().Address.RowNumber; ++rowIndex)
                {
                    var row = sheet.Row(rowIndex);
                    if (row == null) continue;

                    for (var cellIndex = row.FirstCellUsed().Address.ColumnNumber; cellIndex <= row.LastCellUsed().Address.ColumnNumber; ++cellIndex)
                    {
                        var cell = row.Cell(cellIndex);
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
                                    SheetIndex = sheet.Workbook.Worksheets.ToList().IndexOf(sheet) + 1,
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