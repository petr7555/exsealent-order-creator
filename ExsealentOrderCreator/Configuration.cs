using ClosedXML.Excel;

namespace ExsealentOrderCreator
{
    public class Configuration
    {
        public Configuration()
        {
            InputHeaderColumns = new[]
            {
                ColInClassification, ColInPcsAvailable, ColInProduct, ColInColor, ColInSize, ColInPrice,
                ColInDetailOne, ColInDescription, ColInDetailTwo
            };

            OutputHeaderColumns = new[]
            {
                ColOutPicture, ColOutCollection, ColOutProduct, ColOutDescription, ColOutCategory, ColOutPriceNoVat,
                ColOutRecommendedPriceWithVat, ColOutColor, ColOutSizeInStockOrder, ColOutTotalPcs,
                ColOutTotalPriceNoVat
            };

            ImageRowWidth = RowHeight / RowHeightImageWidthRatio;
        }

        public const string ConfigurationFilePath = "configuration.yaml";

        /** Also set from configuration.yaml **/
        public string InputWorkbookPath { get; set; } = "Data.xlsx";
        public string InputWorksheetName { get; set; } = "DATA CZK";

        public string OutputWorkbookPath { get; set; } = "Nabidka.xlsx";
        public string OutputWorksheetName { get; set; } = "Nabídka";

        public string ImageFolderPath { get; set; } = "Fotky";

        /** End of Also set from configuration.yaml **/

        public string CompressedImagesDirectoryName { get; } = "_compressed";
        public int ResizeRatio { get; } = 2;

        public int NumSizes { get; set; }

        public int HeaderRowIndex { get; } = 1;


        public string ColInClassification { get; } = "Zařazení";
        public string ColInPcsAvailable { get; } = "K dispozici";
        public string ColInProduct { get; } = "Produkt";
        public string ColInColor { get; } = "Barva";
        public string ColInSize { get; } = "Velikost";
        public string ColInPrice { get; } = "Cena";
        public string ColInDetailOne { get; } = "Údaj 1";
        public string ColInDescription { get; } = "Popis";
        public string ColInDetailTwo { get; } = "Údaj 2";

        public string[] InputHeaderColumns { get; }


        public string ColOutPicture { get; } = "Obrázek";
        public string ColOutCollection { get; } = "Kolekce";
        public string ColOutProduct { get; } = "Produkt";
        public string ColOutDescription { get; } = "Popis";
        public string ColOutCategory { get; } = "Kategorie";
        public string ColOutPriceNoVat { get; } = "Cena nákup bez DPH";
        public string ColOutRecommendedPriceWithVat { get; } = "DMOC s DPH";
        public string ColOutColor { get; } = "Barva";
        public string ColOutSizeInStockOrder { get; } = "VELIKOST/SKLADEM/OBJEDNÁVKA";
        public string ColOutTotalPcs { get; } = "CELKEM ks";
        public string ColOutTotalPriceNoVat { get; } = "CELKEM bez DPH";

        public string[] OutputHeaderColumns { get; }

        public double ColOutSizeInStockOrderWidth { get; } = 7;
        
        public string EurFormat { get; } = "#,##0.00 €";
        public string CzkFormat { get; } = "#,##0 Kč";
        public string Czk { get; } = "CZK";

        public XLColor Yellow { get; } = XLColor.Yellow;
        public XLColor LightBlue { get; } = XLColor.FromArgb(230, 255, 255);

        public double RowHeight { get; } = 100;
        // Width is right in Excel. In WPS office it looks different (too wide).
        public double RowHeightImageWidthRatio { get; } = 5.556;
        public double ImageRowWidth { get; }
        
        // In order not to hide borders with an image
        public int ImageXOffset { get; } = 2;
        public int ImageYOffset { get; } = 2;
    }
}
