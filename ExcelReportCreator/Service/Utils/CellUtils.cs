using System;
using ClosedXML.Excel;
using ExcelReportCreatorProject.Domain.Markers;

namespace ExcelReportCreatorProject.Service.Utils
{
    public class CellUtils
    {
        public static bool IsMarkedCell(IXLCell cell, MarkerOptions markerOptions)
        {
            if (cell.DataType == XLDataType.Text)
            {
                var stringCellValue = cell.GetString().Trim();
                if (stringCellValue.Length < 4)
                    return false;
                var isPrefixMatch = stringCellValue.Substring(0, markerOptions.Prefix.Length) == markerOptions.Prefix;
                var isSuffixMatch = stringCellValue.Substring(stringCellValue.Length - markerOptions.Suffix.Length, markerOptions.Suffix.Length) == markerOptions.Suffix;
                if (isPrefixMatch && isSuffixMatch)
                    return true;
            }
            return false;
        }

        public static string ExtractMarkerValue(IXLCell cell, MarkerOptions markerOptions)
        {
            var stringCellValue = cell.GetString().Trim();
            return stringCellValue.Substring(markerOptions.Prefix.Length, stringCellValue.Length - (markerOptions.Prefix.Length + markerOptions.Suffix.Length));
        }

        public static void SetDynamicCellValue(IXLCell cell, object value)
        {
            switch (value)
            {
                case string stringValue:
                    cell.SetDataType(XLDataType.Text);
                    cell.SetValue(stringValue);
                    break;
                case int intValue:
                    cell.SetDataType(XLDataType.Number);
                    cell.SetValue(intValue);
                    break;
                case double doubleValue:
                    cell.SetDataType(XLDataType.Number);
                    cell.SetValue(doubleValue);
                    break;
                default:
                    throw new Exception($"Неизвестный тип: {value?.GetType().Name}");
            }
        }
    }
}