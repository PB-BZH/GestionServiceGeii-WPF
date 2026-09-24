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
║  Nom de fichier : ServiceWorkbookSheetCatalog.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
namespace GestionServiceGeii.Core.Service;

internal sealed class ServiceWorkbookSheetInfo {
  internal string SheetName { get; set; } = string.Empty;

  internal string[] LogicalSemesters { get; set; } =
      Array.Empty<string>();
  internal bool UseForGroupEnrichment { get; set; }

  internal bool IsSemesterSheet {
    get {
      return LogicalSemesters.Length > 0;
    }
  }
}

internal static class ServiceWorkbookSheetCatalog {
  internal const string Global =
      "Global";

  internal const string S1 =
      "S1";

  internal const string S2 =
      "S2";

  internal const string S3Fifa =
      "S3 FIFA";

  internal const string S4Fifa =
      "S4 FIFA";

  internal const string S5EtS6 =
      "S5 & S6";

  internal const string S5S6HpPn =
      "S5S6 HP PN";

  internal const string ModuleComplCadets =
      "Module Compl Cadets";

  internal static readonly ServiceWorkbookSheetInfo[] SheetsToCache = {
  new ServiceWorkbookSheetInfo {
    SheetName = Global
  },

  new ServiceWorkbookSheetInfo {
    SheetName = S1,
    LogicalSemesters = new[] { "S1" },
    UseForGroupEnrichment = true
  },

  new ServiceWorkbookSheetInfo {
    SheetName = S2,
    LogicalSemesters = new[] { "S2" },
    UseForGroupEnrichment = true
  },

  new ServiceWorkbookSheetInfo {
    SheetName = S3Fifa,
    LogicalSemesters = new[] { "S3" },
    UseForGroupEnrichment = false
  },

  new ServiceWorkbookSheetInfo {
    SheetName = S4Fifa,
    LogicalSemesters = new[] { "S4" },
    UseForGroupEnrichment = false
  },

  new ServiceWorkbookSheetInfo {
    SheetName = S5EtS6,
    LogicalSemesters = new[] { "S5","S6" },
    UseForGroupEnrichment = false
  },

  new ServiceWorkbookSheetInfo {
    SheetName = S5S6HpPn,
    LogicalSemesters = new[] { "S5","S6" },
    UseForGroupEnrichment = false
  },

  new ServiceWorkbookSheetInfo {
    SheetName = ModuleComplCadets
  }
};
}