using System.Collections;
using System.Collections.Generic;
using ClosedXML.Excel;
using ExcelReportCreatorProject.Domain.Markers;
using ExcelReportCreatorProject.Service.Utils;

namespace ExcelReportCreatorProject.Service.Extraction
{
    public class MarkerExtractor : IMarkerExtractor, IEnumerable<Marker>
    {
        private readonly IEnumerable<IXLWorksheet> _sheets;
        private readonly MarkerOptions _markerOptions;

        public MarkerExtractor(IXLWorkbook workbook, MarkerOptions markerOptions)
        {
            _sheets = workbook.Worksheets;
            _markerOptions = markerOptions;
        }

        public MarkerExtractor(IEnumerable<IXLWorksheet> sheets, MarkerOptions markerOptions)
        {
            _sheets = sheets;
            _markerOptions = markerOptions;
        }

        public MarkerExtractor(IXLWorksheet sheet, MarkerOptions markerOptions)
        {
            _sheets = new List<IXLWorksheet> { sheet };
            _markerOptions = markerOptions;
        }

        public IEnumerable<Marker> Markers()
        {
            return this;
        }

        IEnumerator<Marker> IEnumerable<Marker>.GetEnumerator()
        {
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

                        if (CellUtils.IsMarkedCell(cell, _markerOptions))
                        {
                            var markerId = CellUtils.ExtractMarkerValue(cell, _markerOptions);
                            var isEndMarker = markerId.Substring(0, _markerOptions.Terminator.Length) == _markerOptions.Terminator;
                            var marker = new Marker
                            {
                                Id = isEndMarker
                                    ? markerId.Substring(_markerOptions.Terminator.Length)
                                    : markerId,
                                Position = new MarkerPosition
                                {
                                    SheetIndex = SheetUtils.SheetIndex(sheet),
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
            return ((IEnumerable<Marker>)this).GetEnumerator();
        }
    }
}