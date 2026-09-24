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
║  Nom de fichier : ExcelSchemaNames.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
namespace GestionServiceGeii.Shared.Database {
  /// <summary>
  /// Centralise les noms des feuilles, tables et colonnes utilisés
  /// dans les classeurs ExcelApp de l'application.
  /// </summary>
  internal static class ExcelSchemaNames {
    /// <summary>
    /// Noms des feuilles ou tables ExcelApp utilisées par l'application.
    /// </summary>
    internal static class Tables {
      /// <summary>
      /// Feuille contenant les informations de formation.
      /// </summary>
      internal const string Formation = "Formation";
      internal const string Cours = "Cours";

      /// <summary>
      /// Feuille contenant les informations des modules.
      /// </summary>
      internal const string Module = "Module";

      /// <summary>
      /// Feuille principale contenant les services détaillés.
      /// </summary>
      internal const string TableComplete = "Table_complete";
      internal const string NomTableGlobal = "Global";
    }

    /// <summary>
    /// Noms des colonnes ExcelApp utilisées dans les différentes feuilles.
    /// </summary>
    internal static class Columns {
      internal const string ID = "ID";
      internal const string Semestre = "Semestre";
      internal const string Formation = "Formation";
      internal const string Module = "Module";
      internal const string LibellePpn = "Libellé_PPN";
      internal const string Cours = "Cours";
      internal const string Noms = "Noms";
      internal const string Duree = "Durée";
      internal const string Nombre = "Nombre";
      internal const string Groupe = "Groupe";
      internal const string TotalType = "Total_type";
      internal const string Commentaires = "Commentaires";
      internal const string Libelle = "Libellé";
      internal const string OSE = "OSE";
      internal const string CM = "CM";
      internal const string TD = "TD";
      internal const string TP = "TP";
      internal const string Parcours = "Parcours";
      internal const string Total = "Total";
      internal const string Salle = "Salle";
      internal const string PPN = "PPN";
      internal const string Apogee = "Apogée";
      internal const string LibelleCourt = "LIBELLE COURT";
      internal const string Infos = "INFOS";
      internal const string StatutIntervenant = "STATUT INTERVENANT";
      internal const string SourceColumn = "SourceColumn";
      internal const string SourceRow = "SourceRow";
      internal const string SourceSheet = "SourceSheet";
      internal const string PN = "PN";
    }
  }
}