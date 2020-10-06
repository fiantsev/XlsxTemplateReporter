﻿using System;
using TemplateCooker.Domain.ResourceObjects;
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
            var resourceObject = context.ResourceObject;

            Console.WriteLine($"sheet: {sheet.Name}");
            Console.WriteLine($"region: marker {{{{{region.StartMarker.Id}}}}} from [{region.StartMarker.Position.RowIndex};{region.StartMarker.Position.CellIndex}] to [{region.EndMarker.Position.RowIndex};{region.EndMarker.Position.RowIndex}]");
            Console.WriteLine($"resourceObject: {resourceObject.GetType().Name}");


            switch (resourceObject)
            {
                case TableResourceObject table:
                    new TableResourceInjector().Inject(context);
                    break;
                case ImageResourceObject image:
                    new ImageResourceInjector().Inject(context);
                    break;
                case TextResourceObject text:
                    new TextResourceInjector().Inject(context);
                    break;
                default:
                    throw new Exception();
            }

        };
    }
}
