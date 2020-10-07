using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Creation;
using TemplateCooker.Service.Extraction;
using TemplateCooker.Service.FormulaCalculation;
using NPOI.XSSF.UserModel;

namespace TemplateCooker
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
            return markerExtractor.GetMarkers().ToList();
        }

        public TemplateBuilder InjectData(DocumentInjectorOptions options)
        {
            var workbook = new XLWorkbook(_template);

            var documentInjector = new DocumentInjector(options);
            documentInjector.Inject(workbook);

            workbook.SaveAs(_template);

            return this;
        }

        public TemplateBuilder RecalculateFormulas(FormulaCalculatorOptions options)
        {
            _template.Seek(0, SeekOrigin.Begin);
            var workbook = new XSSFWorkbook(_template);

            var formulaEvaluator = new FormulaCalculator(options);
            formulaEvaluator.Recalculate(workbook);

            return this;
        }

        public MemoryStream Build()
        {
            _template.Flush();
            _template.Seek(0, SeekOrigin.Begin);
            return _template;
        }

        /// <summary>
        /// NPOI закрывает поток в момент открытия файла, нам это не нужно
        /// </summary>
        private class NonClosingMemoryStream : MemoryStream
        {
            public NonClosingMemoryStream(int capacity) : base(capacity)
            {
            }

            public override void Close()
            {
            }

            public void RealClose()
            {
                base.Close();
            }
        }
    }
}