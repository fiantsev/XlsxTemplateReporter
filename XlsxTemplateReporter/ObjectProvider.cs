using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelReportCreatorProject.Domain.ResourceObjects;
using ExcelReportCreatorProject.Service.ResourceObjectProvision;
using Newtonsoft.Json;

namespace XlsxTemplateReporter
{
    public class ObjectProvider : IResourceObjectProvider
    {
        public Func<string, ResourceObject> Resolve => markerId =>
        {
            switch (markerId)
            {
                case "table1":
                    {
                        var widgetData = PrepareData()["table1"];
                        var table = WidgetDataToListOfList(widgetData);
                        var resource = new TableResourceObject
                        {
                            Table = table
                        };
                        return resource;

                    }
                case "image1":
                    {
                        var imageBytes = File.ReadAllBytes("./Templates/image1.jpg");
                        var resource = new ImageResourceObject
                        {
                            Image = imageBytes
                        };
                        return resource;
                    }
                case "image2":
                    {
                        var imageBytes = File.ReadAllBytes("./Templates/image2_884x2392.png");
                        var resource = new ImageResourceObject
                        {
                            Image = imageBytes
                        };
                        return resource;
                    }
                case "text1":
                    return new TextResourceObject { Text = "www.google.com" };
                default:
                    throw new Exception("По маркеру нет данных");
            }
        };


        static List<List<object>> WidgetDataToListOfList(WidgetData widgetData)
        {

            var result = new List<List<object>>();
            var columns = Invert(widgetData.Cols);
            var values = Invert(widgetData.Values);
            var rows = widgetData.Rows;

            var itemCountInFirstColumn = columns.FirstOrDefault()?.Count ?? 0;
            var itemCountInFirstRow = rows.FirstOrDefault()?.Count ?? 0;

            var combinedRowCount = rows.Count + columns.Count;

            for (var rowIndex = 0; rowIndex < combinedRowCount; ++rowIndex)
            {
                result.Add(new List<object>());
                for (var columnIndex = 0; columnIndex < itemCountInFirstColumn; ++columnIndex)
                {
                    if (rowIndex < columns.Count)
                        if (columnIndex < itemCountInFirstRow)
                            result[rowIndex].Add("");
                        else
                            result[rowIndex].Add(columns[rowIndex][columnIndex - itemCountInFirstRow]);
                    else
                        if (columnIndex < itemCountInFirstRow)
                        result[rowIndex].Add(rows[rowIndex - columns.Count][columnIndex]);
                    else
                        result[rowIndex].Add(values[rowIndex - columns.Count][columnIndex - itemCountInFirstRow]);
                }
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
                })
                //.Select(x=> {
                //    x.Data.Values = Invert(x.Data.Values);
                //    return x;
                //})
                ;

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
