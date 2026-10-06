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
║  Nom de fichier : FicheServiceHelper.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Data;
using GestionServiceGeii.Shared.Database;

namespace GestionServiceGeii.Core.Services;

internal static class FicheServiceHelper {
  internal static IEnumerable<string> ExtraireIntervenants(string valeur) {
    return valeur
        .Split('/',StringSplitOptions.RemoveEmptyEntries)
        .Select(nom => nom.Trim())
        .Where(nom => !string.IsNullOrWhiteSpace(nom));
  }

  internal static List<string> ExtraireIntervenantsDepuisTable(DataTable tableComplete) {
    return tableComplete
        .AsEnumerable()
        .Select(row => LireValeur(row,ExcelSchemaNames.Columns.Noms))
        .SelectMany(ExtraireIntervenants)
        .Where(EstIntervenantValide)
        .Distinct(StringComparer.CurrentCultureIgnoreCase)
        .OrderBy(nom => nom)
        .ToList();
  }

  private static bool EstIntervenantValide(string? nom) {
    if (string.IsNullOrWhiteSpace(nom))
      return false;

    nom = nom.Trim();

    if (nom == "0")
      return false;

    if (
      nom.StartsWith("BUT",StringComparison.OrdinalIgnoreCase) ||
      nom.StartsWith("#",StringComparison.OrdinalIgnoreCase) ||
      nom.StartsWith("?",StringComparison.OrdinalIgnoreCase))
      return false;

    return true;
  }

  internal static List<string> GetIntervenantsDepuisTableComplete(DataTable tableComplete) {
    List<string> intervenants =
        new();

    if (tableComplete == null)
      return intervenants;

    if (!tableComplete.Columns.Contains(ExcelSchemaNames.Columns.Noms))
      return intervenants;

    foreach (DataRow row in tableComplete.Rows) {
      string noms =
          row[ExcelSchemaNames.Columns.Noms]?.ToString()?.Trim() ?? string.Empty;

      if (string.IsNullOrWhiteSpace(noms))
        continue;

      string[] nomsSepares =
          noms.Split(
              '/',
              StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

      foreach (string nom in nomsSepares) {
        string intervenant =
            nom.Trim().ToUpperInvariant();

        if (string.IsNullOrWhiteSpace(intervenant))
          continue;

        if (!intervenants.Contains(intervenant,StringComparer.OrdinalIgnoreCase)) {
          intervenants.Add(intervenant);
        }
      }
    }

    return intervenants;
  }

  internal static RegleStatutService ObtenirRegleStatutService(string statut) {
    return statut switch {
      "PR" => new RegleStatutService(1.5m,1.0m,1.0m,192m),
      "MCF" => new RegleStatutService(1.5m,1.0m,1.0m,192m),

      "PRAG" => new RegleStatutService(1.5m,1.0m,1.0m,384m),
      "PRCE" => new RegleStatutService(1.5m,1.0m,1.0m,384m),

      "Vacataire" => new RegleStatutService(1.0m,1.0m,1.0m,0m),

      // Compatibilité avec les anciennes valeurs éventuelles
      "PR / MCF" => new RegleStatutService(1.5m,1.0m,1.0m,192m),
      "PRAG / PRCE" => new RegleStatutService(1.5m,1.0m,1.0m,384m),

      _ => new RegleStatutService(1.5m,1.0m,1.0m,0m)
    };
  }

  internal static Dictionary<string,LigneServicePrevisionnel> ConstruireLignesServicePrevisionnel(
      DataTable tableComplete,
      string intervenant,
      RegleStatutService regle) {

    Dictionary<string,LigneServicePrevisionnel> lignesParModule =
        new(StringComparer.CurrentCultureIgnoreCase);

    foreach (DataRow ligne in tableComplete.Rows) {
      string noms =
          LireValeur(ligne,ExcelSchemaNames.Columns.Noms);

      bool concerneIntervenant =
          ExtraireIntervenants(noms)
              .Any(nom => string.Equals(
                  nom,
                  intervenant,
                  StringComparison.CurrentCultureIgnoreCase));

      if (!concerneIntervenant)
        continue;

      string ose =
        LireValeur(ligne,ExcelSchemaNames.Columns.OSE);

      string semestre =
        DeterminerSemestreDepuisOse(ose);

      string module =
          LireValeur(ligne,ExcelSchemaNames.Columns.Module);

      string cours =
          LireValeur(ligne,ExcelSchemaNames.Columns.Cours)
              .ToUpperInvariant();

      if (string.IsNullOrWhiteSpace(module))
        continue;

      decimal heures =
          LireHeuresReelles(ligne);

      string cleRegroupement =
          $"{ose}|{module}";

      if (!lignesParModule.TryGetValue(cleRegroupement,out LigneServicePrevisionnel? ligneService)) {
        ligneService =
            new LigneServicePrevisionnel {
              Semestre = semestre,
              Ose = ose,
              Module = module
            };

        lignesParModule.Add(cleRegroupement,ligneService);
      }

      switch (cours) {
        case "CM":
          ligneService.Cm += heures;
          break;

        case "TD":
          ligneService.Td += heures;
          break;

        case "TP":
          ligneService.Tp += heures;
          break;
      }
    }

    foreach (LigneServicePrevisionnel ligne in lignesParModule.Values) {
      ligne.TotalRetenu =
          (ligne.Cm * regle.CmCoefficient) +
          (ligne.Td * regle.TdCoefficient) +
          (ligne.Tp * regle.TpCoefficient);
    }

    return lignesParModule;
  }

  internal static string DeterminerSemestreDepuisOse(string ose) {
    string code =
        (ose ?? string.Empty)
            .Trim()
            .ToUpperInvariant();

    if (code.StartsWith("F2U",StringComparison.OrdinalIgnoreCase) &&
        code.Length >= 4 &&
        char.IsDigit(code[3])) {
      int numeroSemestre =
          code[3] - '0';

      if (numeroSemestre is >= 1 and <= 6) {
        return $"S{numeroSemestre}";
      }
    }

    return "Non renseigné";
  }

  internal static int ObtenirOrdreSemestre(string semestre) {
    string valeur =
        (semestre ?? string.Empty)
            .Trim()
            .ToUpperInvariant();

    if (valeur.StartsWith("S",StringComparison.OrdinalIgnoreCase) &&
        int.TryParse(valeur[1..],out int numero)) {
      return numero;
    }

    return 99;
  }

  internal static string ConstruireLibelleSemestre(string semestre) {
    if (string.IsNullOrWhiteSpace(semestre) ||
        semestre.Equals("Non renseigné",StringComparison.CurrentCultureIgnoreCase)) {
      return "Semestre non renseigné";
    }

    return $"Semestre {semestre}";
  }

  private static string LireValeur(DataRow ligne,string nomColonne) {
    if (!ligne.Table.Columns.Contains(nomColonne))
      return string.Empty;

    return ligne[nomColonne]?.ToString()?.Trim() ?? string.Empty;
  }

  private static decimal LireHeuresReelles(DataRow ligne) {
    decimal duree =
        LireDecimal(
            LireValeur(ligne,ExcelSchemaNames.Columns.Duree));

    decimal nombre =
        LireDecimal(
            LireValeur(ligne,ExcelSchemaNames.Columns.Nombre));

    if (nombre == 0m)
      nombre = 1m;

    return duree * nombre;
  }

  private static decimal LireDecimal(string valeur) {
    if (decimal.TryParse(
        valeur,
        System.Globalization.NumberStyles.Any,
        System.Globalization.CultureInfo.CurrentCulture,
        out decimal resultat)) {
      return resultat;
    }

    if (decimal.TryParse(
        valeur,
        System.Globalization.NumberStyles.Any,
        System.Globalization.CultureInfo.InvariantCulture,
        out resultat)) {
      return resultat;
    }

    return 0m;
  }
}

internal sealed class LigneServicePrevisionnel {
  internal string Semestre { get; set; } = "Non renseigné";
  internal string Ose { get; set; } = string.Empty;
  internal string Module { get; set; } = string.Empty;
  internal decimal Cm { get; set; }
  internal decimal Td { get; set; }
  internal decimal Tp { get; set; }
  internal decimal TotalRetenu { get; set; }
}

internal readonly record struct RegleStatutService(
    decimal CmCoefficient,
    decimal TdCoefficient,
    decimal TpCoefficient,
    decimal ServiceReference);