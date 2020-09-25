using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Test_NpoiDotnet.Operators;

namespace Test_NpoiDotnet
{
    public static class Util
    {
        public static void CopyRange(XSSFSheet sheet, CellRangeAddress from, CellRangeAddress to)
        {
            var rangeCopier = new RangeCopier();
            rangeCopier.CopyRange(sheet, sheet, from, to);
        }
    }
}