# Exsealent Order Creator

Created with [ClosedXML](https://github.com/ClosedXML/ClosedXML).

## How to run

1. [Install .NET 5.0](https://dotnet.microsoft.com/download/dotnet/5.0)
2. In terminal, run `dotnet ExsealentOrderCreator.dll`

## Configuration

- `Configuration.cs` contains default values for all properties.
- User can configure those properties using `configuration.yaml`.
- The key must be written in PascalCase and must be the same as its equivalent inside `Configuration.cs`.
- The most import properties (input file name, output file name etc.) will be re-asked during program execution. If you
  are happy with the default value, press **Enter**.
- The basic `configuration.yaml` looks the following:
  ```yaml
  InputWorkbookPath: ./Data.xlsx
  InputWorksheetName: DATA CZK
  OutputWorkbookPath: ./Nabidka.xlsx
  OutputWorksheetName: Nabídka
  ImageFolderPath: ./Fotky
  ResizeRatio: 2 # image width and height are divided by ResizeRatio
  
  ```
