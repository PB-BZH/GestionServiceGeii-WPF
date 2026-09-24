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
║  Nom de fichier : ExcelQueryBuilder.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
namespace GestionServiceGeii.Shared.Database {
  internal static class ExcelQueryBuilder {
    internal static string SelectAllFromSheet(string sheetName) {
      return string.Format(
          "SELECT * FROM [{0}]",
          sheetName
      );
    }

    /// <summary>
    /// Échappe les apostrophes dans une valeur utilisée dans un filtre DataView / BindingSource.
    /// </summary>
    /// <param name="value">Valeur à sécuriser.</param>
    /// <returns>Valeur échappée, ou chaîne vide si la valeur est null.</returns>
    internal static string EscapeFilterValue(string value) {
      if (value == null)
        return string.Empty;

      return value.Replace("'","''");
    }

    /// <summary>
    /// Construit un filtre d'égalité au format [Colonne] = 'Valeur'.
    /// </summary>
    /// <param name="columnName">Nom de la colonne à filtrer.</param>
    /// <param name="value">Valeur attendue.</param>
    /// <returns>Expression de filtre BindingSource.</returns>
    internal static string BuildEqualsFilter(
        string columnName,
        string value
    ) {
      return string.Format(
          "[{0}] = '{1}'",
          columnName,
          EscapeFilterValue(value)
      );
    }
  }
}