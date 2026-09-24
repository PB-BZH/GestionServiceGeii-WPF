/*
╔════════════════════════════════════════════════════════════════════════════════╗
║                                                                                ║
║                    ───────────────────────────────────                         ║
║                      © Copyright PB-BZH Concept 2026                           ║
║                    ───────────────────────────────────                         ║
║                                                                                ║
║                 contact : mailto:admin@pb-bzh-concept.fr                       ║
╚════════════════════════════════════════════════════════════════════════════════╝

╔════════════════════════════════════════════════════════════════════════════════╗
║  Auteur : Patrick Bourges - PB-BZH Concept                                     ║
║  Le 24/9/2026 - 00:10
╟────────────────────────────────────────────────────────────────────────────────║
║     Projet Visual Studio Professional 2026 : GestionServiceGeii
╟────────────────────────────────────────────────────────────────────────────────║
║     Version : 1.0.0
╟────────────────────────────────────────────────────────────────────────────────║
║                Visual Studio Professional 2026 - Insiders                      ║
║                ──────────────────────────────────────────                      ║
║  Langage     : C# 14                                                           ║
║  Technologie : .NET 10 / WPF                                                   ║
║  Plateforme  : Windows 10 / Windows 11                                         ║
║  Encodage    : UTF-8                                                           ║
╟────────────────────────────────────────────────────────────────────────────────║
║  Nom de fichier : ExcelReferenceResolver.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace GestionServiceGeii.Core.Excel;

internal sealed class CellReferenceTarget {
  public string WorksheetName { get; set; } = string.Empty;

  public string Address { get; set; } = string.Empty;

  public List<string> Chain { get; } =
      new List<string>();
}

internal static class ExcelReferenceResolver {
  private static readonly Regex DirectCellReferenceRegex =
      new Regex(
          @"^(?:(?:'(?<sheetq>[^']+)'|(?<sheet>[^'!]+))!)?\$?(?<col>[A-Z]{1,3})\$?(?<row>[0-9]+)$",
          RegexOptions.IgnoreCase | RegexOptions.Compiled);

  public static CellReferenceTarget ResolveFinalReferencedCell(
      ExcelWorkbook workbook,
      string startWorksheetName,
      string startAddress
  ) {
    if (workbook == null)
      throw new ArgumentNullException(nameof(workbook));

    if (string.IsNullOrWhiteSpace(startWorksheetName))
      throw new ArgumentException("Nom de feuille vide.",nameof(startWorksheetName));

    if (string.IsNullOrWhiteSpace(startAddress))
      throw new ArgumentException("Adresse de cellule vide.",nameof(startAddress));

    string worksheetName =
        startWorksheetName.Trim();

    string address =
        startAddress.Replace("$",string.Empty).Trim();

    CellReferenceTarget result =
        new CellReferenceTarget();

    HashSet<string> visited =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    for (int depth = 0;depth < 30;depth++) {
      string key =
          worksheetName + "!" + address;

      if (!visited.Add(key)) {
        throw new InvalidOperationException(
            "Référence circulaire détectée : " + key);
      }

      result.Chain.Add(key);

      ExcelWorksheet worksheet =
          FindWorksheet(workbook,worksheetName);

      if (worksheet == null) {
        throw new InvalidOperationException(
            "Feuille introuvable : " + worksheetName);
      }

      var cell =
          worksheet.Cells[address];

      string formula =
          cell.Formula?.Trim() ?? string.Empty;

      if (string.IsNullOrWhiteSpace(formula)) {
        result.WorksheetName =
            worksheet.Name;

        result.Address =
            address;

        return result;
      }

      if (!TryParseDirectCellReference(
              formula,
              worksheet.Name,
              out string nextWorksheetName,
              out string nextAddress)) {
        throw new InvalidOperationException(
            "La formule " + key + " n'est pas une référence directe : " + formula);
      }

      worksheetName =
          nextWorksheetName;

      address =
          nextAddress;
    }

    throw new InvalidOperationException(
        "Trop de niveaux de références Excel.");
  }

  private static bool TryParseDirectCellReference(
      string formula,
      string currentWorksheetName,
      out string worksheetName,
      out string address
  ) {
    worksheetName =
        currentWorksheetName;

    address =
        string.Empty;

    if (string.IsNullOrWhiteSpace(formula))
      return false;

    string cleanedFormula =
        formula.Trim();

    if (cleanedFormula.StartsWith("=",StringComparison.Ordinal)) {
      cleanedFormula =
          cleanedFormula.Substring(1).Trim();
    }

    Match match =
        DirectCellReferenceRegex.Match(cleanedFormula);

    if (!match.Success)
      return false;

    if (match.Groups["sheetq"].Success) {
      worksheetName =
          match.Groups["sheetq"].Value.Trim();
    }
    else if (match.Groups["sheet"].Success) {
      worksheetName =
          match.Groups["sheet"].Value.Trim();
    }

    address =
        match.Groups["col"].Value.ToUpperInvariant() +
        match.Groups["row"].Value;

    return true;
  }

  private static ExcelWorksheet FindWorksheet(
      ExcelWorkbook workbook,
      string worksheetName
  ) {
    foreach (ExcelWorksheet worksheet in workbook.Worksheets) {
      if (
          string.Equals(
              worksheet.Name,
              worksheetName,
              StringComparison.OrdinalIgnoreCase)
      ) {
        return worksheet;
      }
    }

    return null!;
  }
}