# Exsealent Order Creator

Created with [ClosedXML](https://github.com/ClosedXML/ClosedXML).

## How to run

1. [Install .NET 5.0](https://dotnet.microsoft.com/download/dotnet/5.0)
2. In terminal, run `dotnet ExsealentOrderCreator.dll`

On Windows, you can create `Run.ps1` file in the same folder as `ExsealentOrderCreator.dll` is. Put the following
contents inside it:

```powershell
dotnet .\ExsealentOrderCreator.dll
Read-Host -Prompt "Press Enter to exit"
```

You can then right-click on it and select `Run with PowerShell`.

## Configuration

- `Configuration.cs` contains default values for all properties.
- User can configure those properties using `configuration.yaml`.
- The key must be written in PascalCase and must be the same as its equivalent inside `Configuration.cs`.
- The basic `configuration.yaml` looks the following:
  ```yaml
  InputWorkbookPath: ./Data.xlsx
  InputWorksheetName: DATA CZK
  OutputWorkbookPath: ./Nabidka.xlsx
  OutputWorksheetName: Nabídka
  ImageFolderPath: ./Fotky
  ResizeRatio: 2 # image width and height are divided by ResizeRatio
  ```

## Features

- If the input worksheet contains column `Cena CZK`, `DMOC s DPH` column in the output worksheet will use CZK format.
  Otherwise, EUR format will be used.

## Troubleshooting

- Make sure the input worksheet contains the following columns:
    - `Zařazení`
    - `K dispozici`
    - `Produkt`
    - `Barva`
    - `Velikost`
    - `Cena`
    - `Údaj 1`
    - `Popis`
    - `Údaj 2`
- Make sure the prices in input worksheet are formatted as number, text or general – **NOT** currency.
- Make sure the paths in `configuration.yaml` are right.
- Make sure `configuration.yaml` is properly formatted.
