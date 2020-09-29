using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using ClosedXML.Excel;
using ExcelReportCreatorProject;
using ExcelReportCreatorProject.Domain;
using ExcelReportCreatorProject.Domain.ResourceObjects;
using ExcelReportCreatorProject.Service.Creator;
using ExcelReportCreatorProject.Service.Injection;
using ExcelReportCreatorProject.Service.MarkerExtraction;
using ExcelReportCreatorProject.Service.ResourceObjectProvider;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Extractor;
using NPOI.XSSF.UserModel;
using Test;
using Test_ClosedXML;
using Test_NpoiDotnet;

namespace XlsxTemplateReporter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var templates = new[]
            {
                "test_1",
            };
            var files = templates
                .Select(x => new {
                    In = $"./Templates/{x}.xlsx",
                    Out = $"./Output/{x}.out.xlsx"
                })
                .ToList();

            files.ForEach(file =>
            {
                Console.WriteLine($"workbook: {file}");
                using var fileStream = File.Open(file.In, FileMode.Open, FileAccess.ReadWrite);
                var workbook = new XSSFWorkbook(fileStream);

                var resourceInjector = new ResourceInjector(ctx => {
                    var region = ctx.MarkerRegion;
                    var sheet = ctx.Workbook.GetSheetAt(region.StartMarker.Position.SheetIndex);
                    var resourceObject = ctx.ResourceObject;

                    Console.WriteLine($"sheet: {sheet.SheetName}");
                    Console.WriteLine($"region: marker {{{{{region.StartMarker.Id}}}}} from [{region.StartMarker.Position.RowIndex};{region.StartMarker.Position.CellIndex}] to [{region.EndMarker.Position.RowIndex};{region.EndMarker.Position.RowIndex}]");
                    Console.WriteLine($"resourceObject: {resourceObject.GetType().Name}");

                    switch (resourceObject)
                    {
                        case TableResourceObject table: ctx.InjectListOfList(table.Table.Values); break;
                        default: throw new Exception();
                    }
                    
                });
                var resourceObjectProvider = new ResourceObjectProvider(markerId =>
                {
                    var widgetData = PrepareData()["table1"];
                    var table = new XTable
                    {
                        Colums = Invert(widgetData.Cols),
                        Rows = widgetData.Rows,
                        Values = Invert(widgetData.Values)
                    };
                    var resource = new TableResourceObject
                    {
                        Table = table
                    };
                    return resource;
                });
                var excelReportCreator = new ExcelReportCreator(new ExcelReportCreatorOptions
                {
                    ResourceInjector = resourceInjector,
                    ResourceObjectProvider = resourceObjectProvider,
                    MarkerExtractorOptions = new MarkerExtractorOptions
                    {
                        MarkerOptions = "{{.}}"
                    }
                });
                excelReportCreator.Create(workbook);

                using(var outputFileStream = File.Open(file.Out, FileMode.Create, FileAccess.ReadWrite))
                    workbook.Write(outputFileStream);
            });

            Console.ReadKey();
        }

        static void Main1(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var templates = new[]
            {
                //"template1",
                //"template2",
                //"template3",
                //"template4",
                //"template5",

                //"0503151_fss",
                //"164604082011",
                //"meropriyatiya_po_snizheniyu_zadolzhennosti_za_yanvar_noyabr_2019_goda_1",
                //"pub_32016",
                //"Бухгалтерская+отчетность+за+6+месяцев+2019г",
                //"Отчет о работе департамента за 1 полугодие 2011",
                //"Отчет-об-использовании-сумм-страховых-взносов-ОТОТ",
                //"с48-2-ноябрь",


                //"test_mergedRegion",
                //"test_rangeCopy",
                "test_totalRow",
            };
            var files = templates
                //.Where(x=>false)
                .Select(x => new { 
                    In = $"./Templates/{x}.xlsx",
                    Out = $"./Output/{x}.out.xlsx"
                })
                .ToList();

            //files.Add(new
            //{
            //    In = $"./Templates/0503151_fss.xls",
            //    Out = $"./Output/0503151_fss.out.xls"
            //});

            //files.ForEach(file =>
            //{
            //    Console.WriteLine($"workbook: {file}");
            //    using var fileStream = File.Open(file.In, FileMode.Open, FileAccess.ReadWrite);
            //    var workbook = new XSSFWorkbook(fileStream);
            //    TemplateParser.PrintFullInfo(workbook);
            //    Util.CopyRange((XSSFSheet)workbook.GetSheetAt(0), CellRangeAddress.ValueOf("B4:D4"), CellRangeAddress.ValueOf("B7:D7"));
            //    (new TemplateDataInjectorService()).InjectData(workbook, PrepareData());
            //    using (var outputFileStream = File.Open(file.Out, FileMode.Create, FileAccess.ReadWrite))
            //        workbook.Write(outputFileStream);
            //});

            //files.ForEach(file =>
            //{
            //    Console.WriteLine($"workbook: {file}");
            //    using var fileStream = File.Open(file.In, FileMode.Open, FileAccess.ReadWrite);
            //    var workbook = new XLWorkbook(fileStream);
            //    var sheet = workbook.Worksheets.First();
            //    (new CopyRange()).Copy(sheet.Range("C2:E4"), sheet.Range("C6:E8"));
            //    (new CopyRange()).Copy(sheet.Range("G2:I5"), sheet.Range("G6:G6"));
            //    using (var outputFileStream = File.Open(file.Out, FileMode.Create, FileAccess.ReadWrite))
            //        workbook.SaveAs(outputFileStream);
            //});

            //Console.WriteLine("Press any key to exit");
            var dataTable = new DataTable("table1");
            dataTable.Columns.Add("header1", typeof(string));
            dataTable.Columns.Add("header2", typeof(string));
            dataTable.Columns.Add("header3", typeof(DateTime));
            dataTable.Rows.Add("r1", "11", DateTime.Now);
            dataTable.Rows.Add("r2", "21", DateTime.UtcNow);
            dataTable.WriteXml(Console.OpenStandardOutput());
            foreach(DataRow row in dataTable.Rows)
                foreach(var el in row.ItemArray)
                    Console.WriteLine(el);
            Console.ReadKey();
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
