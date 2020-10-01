using System.Collections;
using System.Collections.Generic;
using ExcelReportCreatorProject.Domain.Markers.ExtractorOptions;
using ExcelReportCreatorProject.Domain.Npoi;
using ExcelReportCreatorProject.Extensions;
using NPOI.SS.UserModel;

namespace ExcelReportCreatorProject.Domain.Markers
{
    public class MarkerCollection : IEnumerable<Marker>
    {
        private readonly List<ISheet> _sheets;
        private readonly MarkerExtractionOptions _markerExtractionOptions;

        public MarkerCollection(ISheet sheet, MarkerExtractionOptions markerExtractorOptions)
        {
            _sheets = new List<ISheet> { sheet };
            _markerExtractionOptions = markerExtractorOptions;
        }

        public MarkerCollection(IWorkbook workbook, MarkerExtractionOptions markerExtractorOptions)
        {
            _sheets = new List<ISheet>(new SheetCollection(workbook));
            _markerExtractionOptions = markerExtractorOptions;
        }

        public IEnumerator<Marker> GetEnumerator()
        {
            var markerOptions = _markerExtractionOptions.MarkerOptions;

            foreach (var sheet in _sheets)
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