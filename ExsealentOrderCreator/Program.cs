using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using System.Linq;
using System.Text.RegularExpressions;

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
            Configuration config = new();

            // PrintLogo();
            // PrintName();

            // GetInputs(config);

            CreateOrder(config);
        }

        private static void CreateOrder(Configuration config)
        {
            var inWb = new XLWorkbook(config.InputWorkbookName);
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
                InsertRecommendedPriceWithVat(outWs, row, rowIdx, 7, config);
                InsertColor(outWs, row, rowIdx, 8, config);
                var sizeInStockOrderColumnNumber = 9;
                InsertSizeInStockOrder(outWs, rows, rowIdx, sizeInStockOrderColumnNumber, config);
                var totalPcsColumnNumber = 9 + config.NumSizes;
                InsertTotalPcs(outWs, row, rowIdx, totalPcsColumnNumber, config, sizeInStockOrderColumnNumber);
                InsertTotalPriceNoVat(outWs, row, rowIdx, 10 + config.NumSizes, config, priceNoVatColumnNumber,
                    totalPcsColumnNumber);

                rowIdx += 3;
            }

            // fit columns to contents excluding the image column
            foreach (var col in outWs.Columns().Skip(1))
            {
                col.AdjustToContents();
            }

            // TODO might not be needed
            // Set workbook calculation mode
            // When you make any change to the document, all affected parts of the document are recalculated.
            outWb.CalculateMode = XLCalculateMode.Auto;

            outWb.SaveAs(config.OutputWorkbookName);
        }

        private static void InsertHeader(IXLWorksheet ws, Configuration config)
        {
            var idx = 1;
            foreach (var column in config.OutputHeaderColumns)
            {
                if (column == config.ColOutSizeInStockOrder)
                {
                    var range = ws.Range(ws.Cell(1, idx), ws.Cell(1, idx + config.NumSizes - 1));
                    range.Merge();
                    range.Value = column;
                    idx += config.NumSizes;

                    // styling
                    range.Style.Fill.BackgroundColor = config.Yellow;
                }
                else
                {
                    ws.Cell(1, idx).Value = column;
                    idx += 1;
                }
            }

            // styling
            var headerRange = ws.Row(config.HeaderRowIndex).RowUsed();
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

            if (FindImagePath(config.ImgFolder, imgName, out var imgPath))
            {
                var image = ws.AddPicture(imgPath)
                    .MoveTo(cell);
                image.Scale(config.RowHeight / (image.OriginalHeight * 0.75d));
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
            Configuration config)
        {
            var cell = ws.Cell(rowIdx, columnNumber);
            cell.Value = row.Field(config.ColInPrice).GetString();
            // styling
            cell.Style.NumberFormat.Format = config.CzkFormat;
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
                var variant = row.Field(config.ColInVariant).GetString();
                var size = variant.Substring(variant.IndexOf("-", StringComparison.Ordinal) + 1);
                return size;
            }

            // Order from smallest size to largest
            rows = rows.OrderBy(row =>
            {
                var sizeStr = GetSize(row);
                var size = int.Parse(sizeStr.Split('/').First());
                return size;
            }).ToList();

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
                new InputQuestion("Input workbook name", config.InputWorkbookName,
                    input => config.InputWorkbookName = input),
                new InputQuestion("Input worksheet name", config.InputWorksheetName,
                    input => config.InputWorksheetName = input),
                new InputQuestion("Output workbook name", config.OutputWorkbookName,
                    input => config.OutputWorkbookName = input),
                new InputQuestion("Output worksheet name", config.OutputWorksheetName,
                    input => config.OutputWorksheetName = input),
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

            var extensions = new[] {"jpg", "png", "jpeg"};

            foreach (var extension in extensions)
            {
                var possiblePath = Path.Combine(imgFolder, $"{imgName}.{extension}");
                if (!File.Exists(possiblePath)) continue;
                imgPath = possiblePath;
                return true;
            }

            imgPath = "";
            return false;
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
