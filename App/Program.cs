﻿using System;
using System.Data;
using System.IO;
using System.Linq;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service;
using TemplateCooker.Service.Creation;
using TemplateCooker.Service.FormulaCalculation;

namespace XlsxTemplateReporter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var templates = new[]
            {
                //"template1",
                "Данные для расчета КП 1 (1)",
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
                using var fileStream = File.Open(file.In, FileMode.Open, FileAccess.Read);

                var templateBuilder = new TemplateBuilder(fileStream);
                var markerOptions = new MarkerOptions("{{", ".", "}}");

                //при реальном использование есть необходимость извлечь все маркеры прежде чем двигаться дальше
                //маркеры необходимы для того что бы отправить запрос за данными
                var allMarkers = templateBuilder.ReadMarkers(markerOptions);
                Console.WriteLine($"Found {allMarkers.Count}: {string.Join(',', allMarkers.Select(x => x.Id))}");

                var resourceInjector = new ResourceInjector();
                var injectionProvider = new InjectionProvider();
                var documentInjectorOptions = new DocumentInjectorOptions
                {
                    ResourceInjector = resourceInjector,
                    InjectionProvider = injectionProvider,
                    MarkerOptions = markerOptions,
                };

                var documentStream = templateBuilder
                    .InjectData(documentInjectorOptions)
                    .SetupFormulaCalculations(new FormulaCalculationOptions { ForceFullCalculation = true, FullCalculationOnLoad = true })
                    .RecalculateFormulasOnBuild(false)
                    .Build(false);

                using (var outputFileStream = File.Open(file.Out, FileMode.Create, FileAccess.ReadWrite))
                    documentStream.CopyTo(outputFileStream);
            });

            //Console.ReadKey();
        }
    }
}