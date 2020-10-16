using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.InjectionProviders;

namespace XlsxTemplateReporter
{
    public class InjectionProvider : IInjectionProvider
    {
        public Injection Resolve(string markerId)
        {
            switch (markerId)
            {
                case "table1":
                    {
                        var widgetData = PrepareData()["table1"];
                        var table = WidgetDataToListOfList(widgetData);
                        return new TableInjection { Resource = new TableResourceObject(table), LayoutShift = LayoutShiftType.MoveRows };

                    }
                case "table2":
                    {
                        var widgetData = PrepareData()["table1"];
                        var table = WidgetDataToListOfList(widgetData, true);
                        return new TableInjection { Resource = new TableResourceObject(table), LayoutShift = LayoutShiftType.MoveCells };

                    }
                case "table3":
                    {
                        var widgetData = new WidgetData() { 
                            Cols = new List<List<string>> { new List<string> { "column1" } } ,
                            Rows = new List<List<string>> { new List<string> { "row1" } } ,
                            Values = new List<List<object>> { new List<object>() },
                        };
                        var table = WidgetDataToListOfList(widgetData, false, false);
                        return new TableInjection { Resource = new TableResourceObject(table), LayoutShift = LayoutShiftType.MoveRows };

                    }
                case "image1":
                    {
                        var imageBytes = File.ReadAllBytes("./Templates/image1.jpg");
                        return new ImageInjection { Resource = new ImageResourceObject(imageBytes) };
                    }
                case "image2":
                    {
                        var imageBytes = File.ReadAllBytes("./Templates/image2_884x2392.png");
                        return new ImageInjection { Resource = new ImageResourceObject(imageBytes) };
                    }
                case "text1":
                    return new TextInjection { Resource = new TextResourceObject("www.google.com") };
                default:
                    throw new Exception("По маркеру нет данных");
            }
        }

        private List<List<object>> WidgetDataToListOfList(WidgetData widgetDataFrame, bool withColumns = false, bool withRows = false)
        {
            var result = new List<List<object>>();

            var columns = Invert(widgetDataFrame.Cols)
                .Select(x => x.Cast<object>().ToList())
                .ToList();
            var values = Invert(widgetDataFrame.Values);
            var rows = widgetDataFrame.Rows
                .Select(x => x.Cast<object>().ToList())
                .ToList();

            if (withColumns)
                result.AddRange(columns);

            result.AddRange(values);

            if (withRows)
            {
                var rowDifference = result.Count - rows.Count;
                var itemCountInFirstRow = rows.FirstOrDefault()?.Count() ?? 0;

                result = result
                    .Select((resultRow, index) => index < rowDifference
                        ? new List<object>(Enumerable.Repeat((object)string.Empty, itemCountInFirstRow).Concat(resultRow))
                        : new List<object>(rows[index - rowDifference].Concat(resultRow))
                    )
                    .ToList();
            }

            return result;
        }

        static Dictionary<string, WidgetData> PrepareData()
        {
            var files = new[]
            {
                "table1",
            };

            var widgets = files
                .ToList()
                .Select(file => new
                {
                    Name = file,
                    Data = JsonConvert.DeserializeObject<WidgetData>(File.ReadAllText($"./Data/{file}.json"))
                });

            var result = widgets.ToDictionary(x => x.Name, x => x.Data);
            return result;
        }

        static List<List<T>> Invert<T>(List<List<T>> array)
        {
            var result = new List<List<T>>();

            if (array.Count == 0)
                return result;

            var rowCount = array.Count;
            var colCount = array[0].Count;

            for (var col = 0; col < colCount; ++col)
            {
                result.Add(new List<T>());
                for (var row = 0; row < rowCount; ++row)
                {
                    result[col].Add(array[row][col]);
                }
            }

            return result;
        }
    }
}
