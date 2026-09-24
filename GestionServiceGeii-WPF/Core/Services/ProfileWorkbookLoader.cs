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
║  Nom de fichier : ProfileWorkbookLoader.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using GestionServiceGeii.Core.Models;
using GestionServiceGeii.Core.Profiles;
using GestionServiceGeii.Shared.Librairie_Fichier;

namespace GestionServiceGeii.Core.Services {
  internal static class ProfileWorkbookLoader {
    internal static ProfileWorkbookLoadResult Load(ServiceManagerProfile profile) {
      ProfileWorkbookLoadResult result = new();

      if (profile == null) {
        result.Success = false;
        result.ErrorMessage = "Le profil n'est pas initialisé.";
        return result;
      }

      if (profile.Files == null) {
        result.Success = false;
        result.ErrorMessage = "La section Files du profil est manquante.";
        return result;
      }

      ClasseExcel selectionWorkbook = new();

      ClasseExcel serviceWorkbook = new();

      if (!selectionWorkbook.ChargerDepuisChemin(profile.Files.ListsWorkbookPath)) {
        result.Success = false;
        result.ErrorMessage =
            "Impossible de charger le fichier de sélection depuis le profil.";

        return result;
      }

      if (!serviceWorkbook.ChargerDepuisChemin(profile.Files.ServiceWorkbookPath)) {
        result.Success = false;
        result.ErrorMessage =
            "Impossible de charger le fichier de service depuis le profil.";

        return result;
      }

      result.SelectionWorkbook = selectionWorkbook;

      result.ServiceWorkbook = serviceWorkbook;

      result.Success = true;

      return result;
    }
  }
}