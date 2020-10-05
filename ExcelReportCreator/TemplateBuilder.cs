using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelReportCreatorProject.Domain.Markers;
using ExcelReportCreatorProject.Service.Creation;
using ExcelReportCreatorProject.Service.Extraction;
using ExcelReportCreatorProject.Service.FormulaCalculation;
using NPOI.XSSF.UserModel;

namespace ExcelReportCreatorProject
{
    public class TemplateBuilder
    {
        private readonly NonClosingMemoryStream _template;

        public TemplateBuilder(Stream template)
        {
            _template = new NonClosingMemoryStream((int)template.Length);
            template.CopyTo(_template);
            _template.Seek(0, SeekOrigin.Begin);
        }

        public List<Marker> ReadMarkers(MarkerOptions markerOptions)
        {
            var workbook = new XLWorkbook(_template);
            var markerExtractor = new MarkerExtractor(workbook, markerOptions);
            return markerExtractor.Markers().ToList();
        }

        public TemplateBuilder InjectData(DocumentInjectorOptions options)
        {
            var workbook = new XLWorkbook(_template);
            var documentInjector = new DocumentInjector(options);
            documentInjector.Inject(workbook);

            workbook.SaveAs(_template);
            //_template.Flush();
            //_template.Seek(0, SeekOrigin.Begin);

            return this;
        }

        public TemplateBuilder RecalculateFormulas(FormulaCalculatorOptions options)
        {
            _template.Seek(0, SeekOrigin.Begin);
            var workbook = new XSSFWorkbook(_template);
            var formulaEvaluator = new FormulaCalculator(options);
            formulaEvaluator.Recalculate(workbook);

            //workbook.Write(_template, leaveOpen:true);
            //_template.Flush();
            //_template.Seek(0, SeekOrigin.Begin);

            return this;
        }

        public Stream Build()
        {
            _template.Flush();
            _template.Close();
            _template.Seek(0, SeekOrigin.Begin);
            return _template;
        }
    }

    public class NonClosingMemoryStream : MemoryStream
    {
        public NonClosingMemoryStream(int capacity) : base(capacity)
        {

        }

        public override void Close()
        {
            //base.Close();
        }
    }
}