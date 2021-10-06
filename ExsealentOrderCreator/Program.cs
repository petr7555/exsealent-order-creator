using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace ExsealentOrderCreator
{
    internal class InputQuestion
    {
        public string Question { get; }
        public string DefaultValue { get; }
        public Action<string> ConfigurationSetter { get; }

        public InputQuestion(string question, string defaultValue, Action<string> configurationSetter)
        {
            Question = question;
            DefaultValue = defaultValue;
            ConfigurationSetter = configurationSetter;
        }
    }

    internal static class Program
    {
        private static void Main(string[] args)
        {
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(PascalCaseNamingConvention.Instance)
                .Build();

            var config = deserializer.Deserialize<Configuration>(File.ReadAllText(Configuration.ConfigurationFilePath));

            PrintLogo();
            PrintName();

            GetInputs(config);

            Console.WriteLine("Creating order...");
            CreateOrder(config);
            Console.WriteLine("Order successfully created.");
        }

        private static void CreateOrder(Configuration config)
        {
            var inWb = new XLWorkbook(config.InputWorkbookPath);
            var inWs = inWb.Worksheet(config.InputWorksheetName);

            var outWb = new XLWorkbook();
            var outWs = outWb.Worksheets.Add(config.OutputWorksheetName);

            var inputRange = inWs.RangeUsed();
            var inputTable = inputRange.AsTable();

            var groupedRows = inputTable.DataRange.Rows()
                .GroupBy(row => new
                {
                    ColInProduct = row.Field(config.ColInProduct).GetString(),
                    ColInColor = row.Field(config.ColInColor).GetString()
                })
                .ToList();

            var maxSizes = groupedRows.Max(row => row.Count());
            config.NumSizes = maxSizes;

            InsertHeader(outWs, config);

            var rowIdx = config.HeaderRowIndex + 1;
            foreach (var rowGroup in groupedRows)
            {
                var row = rowGroup.First();
                var rows = rowGroup.ToList();

                InsertImage(outWs, row, rowIdx, 1, config);
                InsertCollection(outWs, row, rowIdx, 2, config);
                InsertProduct(outWs, row, rowIdx, 3, config);
                InsertDescription(outWs, row, rowIdx, 4, config);
                InsertCategory(outWs, row, rowIdx, 5, config);
                var priceNoVatColumnNumber = 6;
                InsertPriceNoVat(outWs, row, rowIdx, priceNoVatColumnNumber, config);
                InsertRecommendedPriceWithVat(outWs, row, rowIdx, 7, config, inputTable);
                InsertColor(outWs, row, rowIdx, 8, config);
                var sizeInStockOrderColumnNumber = 9;
                InsertSizeInStockOrder(outWs, rows, rowIdx, sizeInStockOrderColumnNumber, config);
                var totalPcsColumnNumber = 9 + config.NumSizes;
                InsertTotalPcs(outWs, row, rowIdx, totalPcsColumnNumber, config, sizeInStockOrderColumnNumber);
                InsertTotalPriceNoVat(outWs, row, rowIdx, 10 + config.NumSizes, config, priceNoVatColumnNumber,
                    totalPcsColumnNumber);

                rowIdx += 3;
            }

            InsertTotalPriceBox(outWs, config, 9 + config.NumSizes - 3);

            // fit columns to contents excluding the image column
            foreach (var col in outWs.Columns().Skip(1))
            {
                col.AdjustToContents();
            }

            // Set workbook calculation mode
            // When you make any change to the document, all affected parts of the document are recalculated.
            outWb.CalculateMode = XLCalculateMode.Auto;

            outWb.SaveAs(config.OutputWorkbookPath);
        }

        private static void InsertTotalPriceBox(IXLWorksheet ws, Configuration config, int columnNumber)
        {
            var rowsCount = ws.Rows().Count();
            var leftLabelCell = ws.Cell(config.HeaderRowIndex + rowsCount + 1, columnNumber);
            var rightLabelCell = leftLabelCell.CellRight().CellRight();
            var labelRange = ws.Range(leftLabelCell, rightLabelCell);
            labelRange.Merge();
            labelRange.Value = "Celkem:";

            var totalPcsCell = rightLabelCell.CellRight();
            totalPcsCell.FormulaA1 =
                $"SUM({ws.Cell(config.HeaderRowIndex + 1, totalPcsCell.Address.ColumnNumber)}:{ws.Cell(config.HeaderRowIndex + rowsCount - 1, totalPcsCell.Address.ColumnNumber)})";
            // conditional formatting
            totalPcsCell
                .AddConditionalFormat()
                .WhenGreaterThan(0)
                .Fill.SetBackgroundColor(config.Yellow);

            var totalPriceCell = totalPcsCell.CellRight();
            totalPriceCell.FormulaA1 =
                $"SUM({ws.Cell(config.HeaderRowIndex + 1, totalPriceCell.Address.ColumnNumber)}:{ws.Cell(config.HeaderRowIndex + rowsCount - 1, totalPriceCell.Address.ColumnNumber)})";
            // styling
            totalPriceCell.Style.NumberFormat.Format = config.EurFormat;

            var boxRange = ws.Range(leftLabelCell, totalPriceCell);
            // styling
            boxRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            boxRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            boxRange.Style.Font.Bold = true;
            boxRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
        }

        private static void InsertHeader(IXLWorksheet ws, Configuration config)
        {
            var idx = 1;
            foreach (var column in config.OutputHeaderColumns)
            {
                if (column == config.ColOutSizeInStockOrder)
                {
                    var range = ws.Range(ws.Cell(config.HeaderRowIndex, idx),
                        ws.Cell(config.HeaderRowIndex, idx + config.NumSizes - 1));
                    range.Merge();
                    range.Value = column;
                    idx += config.NumSizes;

                    // styling
                    range.Style.Fill.BackgroundColor = config.Yellow;
                }
                else
                {
                    ws.Cell(config.HeaderRowIndex, idx).Value = column;
                    idx += 1;
                }
            }

            // styling
            var headerRange = ws.Row(config.HeaderRowIndex).RowUsed();
            headerRange.Style.Alignment.WrapText = true;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.Row(config.HeaderRowIndex).Height = 40;
        }

        private static void InsertImage(IXLWorksheet ws, IXLTableRow row, int rowIdx, int columnNumber,
            Configuration config)
        {
            var imgName = $"{row.Field(config.ColInProduct).GetString()}-{row.Field(config.ColInColor).GetString()}";
            var cell = ws.Cell(rowIdx, columnNumber);

            if (FindImagePath(config.ImageFolderPath, imgName, out var imgPath))
            {
                var image = ws.AddPicture(imgPath)
                    .MoveTo(cell, config.ImageXOffset, config.ImageYOffset, ws.Cell(rowIdx + 1, columnNumber + 1),
                        -config.ImageXOffset, -config.ImageYOffset);
            }

            // styling
            ws.Row(rowIdx).Height = config.RowHeight;
            ws.Column(columnNumber).Width = config.ImageRowWidth;

            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Fill.BackgroundColor = config.LightBlue;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private static void InsertCollection(IXLWorksheet ws, IXLTableRow row, int rowIdx, int columnNumber,
            Configuration config)
        {
            var cell = ws.Cell(rowIdx, columnNumber);
            cell.Value = row.Field(config.ColInClassification).GetString();
            // styling
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Fill.BackgroundColor = config.LightBlue;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private static void InsertProduct(IXLWorksheet ws, IXLTableRow row, int rowIdx, int columnNumber,
            Configuration config)
        {
            var cell = ws.Cell(rowIdx, columnNumber);
            cell.Value = row.Field(config.ColInProduct).GetString();
            // styling
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Fill.BackgroundColor = config.LightBlue;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private static void InsertDescription(IXLWorksheet ws, IXLTableRow row, int rowIdx, int columnNumber,
            Configuration config)
        {
            var cell = ws.Cell(rowIdx, columnNumber);
            cell.Value = row.Field(config.ColInDescription).GetString();
            // styling
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Fill.BackgroundColor = config.LightBlue;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private static void InsertCategory(IXLWorksheet ws, IXLTableRow row, int rowIdx, int columnNumber,
            Configuration config)
        {
            var cell = ws.Cell(rowIdx, columnNumber);
            cell.Value = row.Field(config.ColInDetailTwo).GetString();
            // styling
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Fill.BackgroundColor = config.LightBlue;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private static void InsertPriceNoVat(IXLWorksheet ws, IXLTableRow row, int rowIdx, int columnNumber,
            Configuration config)
        {
            var cell = ws.Cell(rowIdx, columnNumber);
            cell.Value = row.Field(config.ColInDetailOne).GetString();
            // styling
            cell.Style.NumberFormat.Format = config.EurFormat;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Fill.BackgroundColor = config.LightBlue;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private static void InsertRecommendedPriceWithVat(IXLWorksheet ws, IXLTableRow row, int rowIdx,
            int columnNumber,
            Configuration config, IXLTable table)
        {
            var priceColumn = table.Fields.First(field => field.Name.StartsWith(config.ColInPrice)).Name;

            var cell = ws.Cell(rowIdx, columnNumber);
            cell.Value = row.Field(priceColumn).GetString();
            // styling
            cell.Style.NumberFormat.Format =
                priceColumn.ToUpper().Contains(config.Czk) ? config.CzkFormat : config.EurFormat;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Fill.BackgroundColor = config.LightBlue;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private static void InsertColor(IXLWorksheet ws, IXLTableRow row, int rowIdx, int columnNumber,
            Configuration config)
        {
            var cell = ws.Cell(rowIdx, columnNumber);
            cell.Value = row.Field(config.ColInColor).GetString();
            // styling
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Fill.BackgroundColor = config.LightBlue;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private static void InsertSizeInStockOrder(IXLWorksheet ws, IReadOnlyList<IXLTableRow> rows, int rowIdx,
            int columnNumber,
            Configuration config)
        {
            string GetSize(IXLTableRow row)
            {
                return row.Field(config.ColInSize).GetString();
            }

            // Order from smallest size to largest
            // Compares numbers, then strings
            // e.g. 86, 86/92, 104/110, 104, XS, S, M, L, XL, XXL, ONE
            rows = rows.OrderBy(GetSize,
                new SemiNumericComparer()).ToList();

            for (var i = 0; i < rows.Count; i++)
            {
                // Size
                var sizeCell = ws.Cell(rowIdx, columnNumber + i);
                sizeCell.Value = GetSize(rows[i]);
                // styling
                sizeCell.Style.Font.Bold = true;
                sizeCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                sizeCell.Style.Fill.BackgroundColor = config.LightBlue;
                sizeCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sizeCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                // Available
                var pcsAvailableCell = ws.Cell(rowIdx + 1, columnNumber + i);
                pcsAvailableCell.Value = rows[i].Field(config.ColInPcsAvailable).GetString();
                // styling
                pcsAvailableCell.Style.Font.Bold = true;
                pcsAvailableCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                pcsAvailableCell.Style.Fill.BackgroundColor = config.LightBlue;
                pcsAvailableCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                pcsAvailableCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                // Field for noting order
                var pcsOrderedCell = ws.Cell(rowIdx + 2, columnNumber + i);
                // styling
                pcsOrderedCell.Style.Font.Bold = true;
                pcsOrderedCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                pcsOrderedCell.Style.Fill.BackgroundColor = config.Yellow;
                pcsOrderedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                pcsOrderedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }
        }

        private static void InsertTotalPcs(IXLWorksheet ws, IXLTableRow row, int rowIdx, int columnNumber,
            Configuration config, int sizeInStockOrderColumnNumber)
        {
            var cell = ws.Cell(rowIdx + 2, columnNumber);
            var sizeInStockOrderStartingCellAddress = ws.Cell(rowIdx + 2, sizeInStockOrderColumnNumber).Address;
            var sizeInStockOrderEndingCellAddress =
                ws.Cell(rowIdx + 2, sizeInStockOrderColumnNumber + config.NumSizes - 1).Address;
            cell.FormulaA1 =
                $"SUM({sizeInStockOrderStartingCellAddress}:{sizeInStockOrderEndingCellAddress})";
            // styling
            cell.Style.Font.Bold = true;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // conditional formatting 
            ws.Range(cell.CellAbove().CellAbove(), cell)
                .AddConditionalFormat()
                .WhenIsTrue($"{cell.Address.ToStringFixed()}>0")
                .Fill.SetBackgroundColor(config.Yellow);
        }

        private static void InsertTotalPriceNoVat(IXLWorksheet ws, IXLTableRow row, int rowIdx, int columnNumber,
            Configuration config, int priceNoVatColumnNumber, int totalPcsColumnNumber)
        {
            var cell = ws.Cell(rowIdx + 2, columnNumber);
            var priceNoVatCellAddress = ws.Cell(rowIdx, priceNoVatColumnNumber).Address;
            var totalPcsCellAddress = ws.Cell(rowIdx + 2, totalPcsColumnNumber).Address;
            cell.FormulaA1 = $"{priceNoVatCellAddress}*{totalPcsCellAddress}";
            // styling
            cell.Style.NumberFormat.Format = config.EurFormat;
            cell.Style.Font.Bold = true;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // conditional formatting 
            ws.Range(cell.CellAbove().CellAbove(), cell)
                .AddConditionalFormat()
                .WhenIsTrue($"{cell.Address.ToStringFixed()}>0")
                .Fill.SetBackgroundColor(config.Yellow);
        }

        private static void GetInputs(Configuration config)
        {
            var inputQuestions = new[]
            {
                new InputQuestion("Input workbook path", config.InputWorkbookPath,
                    input => config.InputWorkbookPath = input),
                new InputQuestion("Input worksheet name", config.InputWorksheetName,
                    input => config.InputWorksheetName = input),
                new InputQuestion("Output workbook path", config.OutputWorkbookPath,
                    input => config.OutputWorkbookPath = input),
                new InputQuestion("Output worksheet name", config.OutputWorksheetName,
                    input => config.OutputWorksheetName = input),
                new InputQuestion("Image folder path", config.ImageFolderPath,
                    input => config.ImageFolderPath = input),
            };

            foreach (var question in inputQuestions)
            {
                Console.Write($"{question.Question} ({question.DefaultValue}): ");
                var input = Console.ReadLine() ?? question.DefaultValue;
                if (string.IsNullOrEmpty(input))
                {
                    input = question.DefaultValue;
                }

                question.ConfigurationSetter(input);
                Console.WriteLine($"Entered value: {input}");
            }
        }

        /**
         * Returns true if image is found and sets imgPath.
         * Returns false if the image is not found, imgPath is set to empty string and should not be used.
         */
        private static bool FindImagePath(string imgFolder, string imgName, out string imgPath)
        {
            var reg = new Regex($"^{imgName}");
            var files = Directory.GetFiles(imgFolder, "*")
                .Where(path => reg.IsMatch(Path.GetFileName(path)));

            imgPath = files.FirstOrDefault();
            return !string.IsNullOrEmpty(imgPath);
        }

        private static void PrintName()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(@"
  ___                      _            _      ___           _              ___                  _             
 | __|__ __ ___ ___  __ _ | | ___  _ _ | |_   / _ \  _ _  __| | ___  _ _   / __| _ _  ___  __ _ | |_  ___  _ _ 
 | _| \ \ /(_-</ -_)/ _` || |/ -_)| ' \|  _| | (_) || '_|/ _` |/ -_)| '_| | (__ | '_|/ -_)/ _` ||  _|/ _ \| '_|
 |___|/_\_\/__/\___|\__,_||_|\___||_||_|\__|  \___/ |_|  \__,_|\___||_|    \___||_|  \___|\__,_| \__|\___/|_|  
                                                                                                               
");
            Console.ResetColor();
        }

        private static void PrintLogo()
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(@"                                                                                
                                               &@@@@@%/   /#@@@@@,              
                                          /@@/                    .@@&          
                                       .@@                            @@*       
                                      @@                                #@/     
                                    @@                                    @@    
                                   @@                                      &@   
                                  @@        &@@@                            @@  
                                 @@         @@@@@              @@@@@         @% 
                                #@                             #@@@@         %@ 
                      &@@@@@    @(    @@@@*         #@@@@@@.                 .@ 
                     @/     *@@%@                     ,@#*         @@@@@      @.
                     @.        @@     #&*.     @@    ,@@@.    @/   /&/       .@ 
                     @@         @@                ,*,     &@@/               #@ 
                      %@.                                                    @& 
                        @@                                                  (@  
                          /@@                                               @.  
                          @@                                               @(   
                        %@,                                              ,@.    
      (@@@@@@@@@*     #@(                                               @@      
   &@*            &@@@.                                               /@/       
   @@   ,##/                                                        #@& %@@@(   
  (@@%                                                           (@@         %@ 
 @*                                                          &@@&            @@ 
 %@/           #@@@* @@@@@/                           .&@@@&      %@@@@@@@&*    
    ,&@@@@@&/                 *#@@@@@@@@@@@@@@@@@@&(                            
");
            Console.ResetColor();
        }
    }
}
