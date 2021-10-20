using ClosedXML.Excel;

namespace ExsealentOrderCreator
{
    /**
     * NOTE: Setters are necessary for YAML configuration to work.
     */
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
        public double ResizeRatio { get; set; } = 2;

        /** End of Also set from configuration.yaml **/

        public string CompressedImagesDirectoryName { get; set; } = "_compressed";

        public int NumSizes { get; set; }

        public int HeaderRowIndex { get; set; } = 1;


        public string ColInClassification { get; set; } = "Zařazení";
        public string ColInPcsAvailable { get; set; } = "K dispozici";
        public string ColInProduct { get; set; } = "Produkt";
        public string ColInColor { get; set; } = "Barva";
        public string ColInSize { get; set; } = "Velikost";
        public string ColInPrice { get; set; } = "Cena";
        public string ColInDetailOne { get; set; } = "Údaj 1";
        public string ColInDescription { get; set; } = "Popis";
        public string ColInDetailTwo { get; set; } = "Údaj 2";

        public string[] InputHeaderColumns { get; set; }


        public string ColOutPicture { get; set; } = "Obrázek";
        public string ColOutCollection { get; set; } = "Kolekce";
        public string ColOutProduct { get; set; } = "Produkt";
        public string ColOutDescription { get; set; } = "Popis";
        public string ColOutCategory { get; set; } = "Kategorie";
        public string ColOutPriceNoVat { get; set; } = "Cena nákup bez DPH";
        public string ColOutRecommendedPriceWithVat { get; set; } = "DMOC s DPH";
        public string ColOutColor { get; set; } = "Barva";
        public string ColOutSizeInStockOrder { get; set; } = "VELIKOST/SKLADEM/OBJEDNÁVKA";
        public string ColOutTotalPcs { get; set; } = "CELKEM ks";
        public string ColOutTotalPriceNoVat { get; set; } = "CELKEM bez DPH";

        public string[] OutputHeaderColumns { get; set; }

        public double ColOutSizeInStockOrderWidth { get; set; } = 7;

        public string EurFormat { get; set; } = "#,##0.00 €";
        public string CzkFormat { get; set; } = "#,##0 Kč";
        public string Czk { get; set; } = "CZK";

        public XLColor Yellow { get; set; } = XLColor.Yellow;
        public XLColor LightBlue { get; set; } = XLColor.FromArgb(230, 255, 255);

        public double RowHeight { get; set; } = 100;

        // Width is right in Excel. In WPS office it looks different (too wide).
        public double RowHeightImageWidthRatio { get; set; } = 5.556;
        public double ImageRowWidth { get; set; }

        // In order not to hide borders with an image
        public int ImageXOffset { get; set; } = 2;
        public int ImageYOffset { get; set; } = 2;
    }
}
