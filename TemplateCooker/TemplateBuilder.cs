using System.IO;
using ExcelReportCreatorProject;
using ExcelReportCreatorProject.Service.Creation;
using ExcelReportCreatorProject.Service.FormulaCalculation;

namespace TemplateCooker
{
    public class TemplateBuilder
    {
        private readonly Stream _template;

        public TemplateBuilder(Stream template)
        {
            template.CopyTo(_template);
        }

        public void InjectData(DocumentInjectorOptions options) {
            var documentInjector = new DocumentInjector(options);
            documentInjector.Inject(_template);
        }

        public void RecalculateFormulas(FormulaCalculatorOptions options)
        {
            var formulaEvaluator = new FormulaCalculator(options);
            formulaEvaluator.Recalculate(_template);
        }

        public Stream Build()
        {
            return _template;
        }
    }
}