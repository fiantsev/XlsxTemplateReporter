using System;
using System.Threading;
using ExcelReportCreatorProject.Domain;
using NPOI.SS.UserModel;

namespace ExcelReportCreatorProject.Extensions
{
    public static class CellExtensions
    {
        public static bool IsMarkedCell(this ICell cell, MarkerOptions markerOptions)
        {
            if (cell.CellType == CellType.String)
            {
                var stringCellValue = cell.StringCellValue.Trim();
                if (stringCellValue.Length < 4)
                    return false;
                var isPrefixMatch = stringCellValue.Substring(0, markerOptions.Prefix.Length) == markerOptions.Prefix;
                var isSuffixMatch = stringCellValue.Substring(stringCellValue.Length - markerOptions.Suffix.Length, markerOptions.Suffix.Length) == markerOptions.Suffix;
                if (isPrefixMatch && isSuffixMatch)
                    return true;
            }
            return false;
        }

        public static string ExtractMarkerValue(this ICell cell, MarkerOptions markerOptions)
        {
            var stringCellValue = cell.StringCellValue.Trim();
            return stringCellValue.Substring(markerOptions.Prefix.Length, cell.StringCellValue.Length - (markerOptions.Prefix.Length + markerOptions.Suffix.Length));
        }

        public static void SetDynamicCellValue(this ICell cell, object value)
        {
            switch (value)
            {
                case string stringValue:
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(stringValue);
                    break;
                case int intValue:
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(intValue);
                    break;
                case double doubleValue:
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(doubleValue);
                    break;
                default:
                    throw new Exception($"Неизвестный тип: {value?.GetType().Name}");
            }
        }
    }
}