using System;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelReportCreatorProject;
using ExcelReportCreatorProject.Domain.Markers;
using ExcelReportCreatorProject.Service.Creation;
using ExcelReportCreatorProject.Service.Extraction;

namespace XlsxTemplateReporter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var templates = new[]
            {
                //"insert image",
                "insert table and image",
            };
            var files = templates
                .Select(x => new
                {
                    In = $"./Templates/{x}.xlsx",
                    Out = $"./Output/{x}.out.xlsx"
                })
                .ToList();

            files.ForEach(file =>
            {
                Console.WriteLine($"workbook: {file}");
                using var fileStream = File.Open(file.In, FileMode.Open, FileAccess.ReadWrite);
                var workbook = new XLWorkbook(fileStream);

                var markerOptions = new MarkerOptions("{{", "}}", ".");
                var markerExtractor = new MarkerExtractor(workbook, markerOptions);
                //при реальном использование есть необходимость извлечь все маркеры прежде чем двигаться дальше
                //маркеры необходимы для того что бы отправить запрос за данными
                var allMarkers = markerExtractor.Markers().ToList();
                Console.WriteLine($"Found {allMarkers.Count}: {string.Join(',', allMarkers.Select(x => x.Id))}");

                var resourceInjector = new ResourceInjector();
                var resourceObjectProvider = new ObjectProvider();
                var excelReportCreator = new ExcelReportUpdator(new ExcelReportUpdatorOptions
                {
                    ResourceInjector = resourceInjector,
                    ResourceObjectProvider = resourceObjectProvider,
                    MarkerExtractor = markerExtractor,
                    FormulaEvaluationOptions = new FormulaEvaluationOptions
                    {
                        EvaluateFormulas = true
                    }
                });
                excelReportCreator.Update(workbook);

                using (var outputFileStream = File.Open(file.Out, FileMode.Create, FileAccess.ReadWrite))
                    workbook.SaveAs(outputFileStream);
            });

            //Console.ReadKey();
        }
    }
}