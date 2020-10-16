﻿using System;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Service.ResourceInjection;
using TemplateCooker.Service.ResourceInjection.Injectors;

namespace XlsxTemplateReporter
{
    public class ResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => context =>
        {
            var region = context.MarkerRange;
            var sheet = context.Workbook.Worksheet(region.StartMarker.Position.SheetIndex);
            var injection = context.Injection;

            Console.WriteLine($"sheet: {sheet.Name}");
            Console.WriteLine($"region: marker {{{{{region.StartMarker.Id}}}}} from [{region.StartMarker.Position.RowIndex};{region.StartMarker.Position.CellIndex}] to [{region.EndMarker.Position.RowIndex};{region.EndMarker.Position.RowIndex}]");
            Console.WriteLine($"resourceObject: {injection.GetType().Name}");


            switch (injection)
            {
                case TableInjection _:
                    new TableResourceInjector().Inject(context);
                    break;
                case ImageInjection _:
                    new ImageResourceInjector().Inject(context);
                    break;
                case TextInjection _:
                    new TextResourceInjector().Inject(context);
                    break;
                default:
                    throw new Exception();
            }

        };
    }
}
