using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using NPOI.XSSF.UserModel;
using Test;

namespace XlsxTemplateReporter
{
    class Program
    {
        static void Main(string[] args)
        {
            var templates = new[]
            {
                //"template1",
                //"template2",
                "template3",
            };
            var files = templates.Select(x => new { 
                In = $"./Templates/{x}.xlsx",
                Out = $"./Output/{x}.out.xlsx"
            });

            files.ToList().ForEach(file =>
            {
                Console.WriteLine($"workbook: {file}");
                using var fileStream = File.Open(file.In, FileMode.Open, FileAccess.ReadWrite);
                var workbook = new XSSFWorkbook(fileStream);
                TemplateParser.PrintFullInfo(workbook);
                var service = new TemplateDataInjectorService();
                service.InjectData(workbook, PrepareData());

                using (var outputFileStream = File.Open(file.Out, FileMode.Create, FileAccess.ReadWrite))
                {
                    workbook.Write(outputFileStream);
                }
            });
            
            //Console.WriteLine("Press any key to exit");
            //Console.ReadKey();
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
                .Select(x=> {
                    x.Data.Values = Invert(x.Data.Values);
                    return x;
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
                for(var row=0;row<rowCount;++row)
                {
                    result[col].Add(array[row][col]);
                }
            }

            return result;
        }
    }
}
