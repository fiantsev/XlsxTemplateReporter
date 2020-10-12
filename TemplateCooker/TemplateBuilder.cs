using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Creation;
using TemplateCooker.Service.Extraction;
using TemplateCooker.Service.FormulaCalculation;

namespace TemplateCooker
{
    public class TemplateBuilder
    {
        private MemoryStream _workbookStream;

        public TemplateBuilder(Stream template)
        {
            _workbookStream = new MemoryStream((int)template.Length);
            template.CopyTo(_workbookStream);
            _workbookStream.Seek(0, SeekOrigin.Begin);
        }

        public List<Marker> ReadMarkers(MarkerOptions markerOptions)
        {
            var workbook = new XLWorkbook(_workbookStream);
            var markerExtractor = new MarkerExtractor(workbook, markerOptions);
            return markerExtractor.GetMarkers().ToList();
        }

        public TemplateBuilder InjectData(DocumentInjectorOptions options)
        {
            var newStream = new MemoryStream((int)_workbookStream.Length);
            _workbookStream.Position = 0;

            using (_workbookStream)
            {
                var workbook = new XLWorkbook(_workbookStream);

                var documentInjector = new DocumentInjector(options);
                documentInjector.Inject(workbook);

                workbook.SaveAs(newStream);
            }

            _workbookStream = newStream;
            _workbookStream.Position = 0;

            return this;
        }

        public TemplateBuilder RecalculateFormulas(FormulaCalculatorOptions options)
        {
            var newStream = new MemoryStream((int)_workbookStream.Length);
            _workbookStream.Position = 0;

            using (_workbookStream)
            {
                var workbook = new XLWorkbook(_workbookStream);

                workbook.ForceFullCalculation = true;
                workbook.FullCalculationOnLoad = true;

                workbook.SaveAs(newStream, new SaveOptions { EvaluateFormulasBeforeSaving = true });
            }

            _workbookStream = newStream;
            _workbookStream.Position = 0;

            return this;
        }

        public MemoryStream Build()
        {
            _workbookStream.Seek(0, SeekOrigin.Begin);
            return _workbookStream;
        }
    }
}