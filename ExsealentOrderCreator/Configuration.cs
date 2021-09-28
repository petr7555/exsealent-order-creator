using ClosedXML.Excel;

namespace ExsealentOrderCreator
{
    public class Configuration
    {
        public Configuration()
        {
            InputHeaderColumns = new[]
            {
                ColInEan, ColInClassification, ColInPcsAvailable, ColInProduct, ColInColor, ColInSize, ColInPrice,
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

        public string InputWorkbookName { get; set; } = "Data.xlsx";
        public string OutputWorkbookName { get; set; } = "Nabidka.xlsx";

        public string InputWorksheetName { get; set; } = "Data new";
        public string OutputWorksheetName { get; set; } = "Nabídka";

        public string ImgFolder { get; set; } = "/Users/petr.janik/Desktop/Excel_Makro/Fotky";
        public int NumSizes { get; set; }

        public int HeaderRowIndex { get; set; } = 1;


        public string ColInEan { get; } = "EAN kode";
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

        public string EurFormat { get; } = "#,##0.00 €";
        public string CzkFormat { get; } = "#,##0 Kč";

        public XLColor Yellow { get; } = XLColor.Yellow;
        public XLColor LightBlue { get; } = XLColor.LightBlue;

        public double RowHeight { get; } = 150;
        public double RowHeightImageWidthRatio { get; } = 5.8;
        public double ImageRowWidth { get; }
    }
}
