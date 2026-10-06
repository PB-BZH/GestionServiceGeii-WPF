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
║  Nom de fichier : ProfileFileValidator.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.IO;
using GestionServiceGeii.Core.Profiles;

namespace GestionServiceGeii.Core.Services {
  internal static class ProfileFileValidator {
    internal static bool AreProfileFilesAvailable(ServiceManagerProfile profile) {
      if (profile == null)
        return false;

      if (profile.Files == null)
        return false;

      return
          !string.IsNullOrWhiteSpace(profile.Files.ServiceWorkbookPath) &&
          File.Exists(profile.Files.ServiceWorkbookPath) &&
          !string.IsNullOrWhiteSpace(profile.Files.ListsWorkbookPath) &&
          File.Exists(profile.Files.ListsWorkbookPath);
    }

    internal static string BuildStatusText(ServiceManagerProfile profile,string currentProfilePath) {
      if (profile == null || profile.Files == null)
        return "Aucun profil valide n'est chargé.";

      string servicePath = profile.Files.ServiceWorkbookPath;
      string listsPath = profile.Files.ListsWorkbookPath;
      bool serviceConfigured = !string.IsNullOrWhiteSpace(servicePath);
      bool listsConfigured = !string.IsNullOrWhiteSpace(listsPath);
      bool serviceExists = serviceConfigured && File.Exists(servicePath);
      bool listsExists = listsConfigured && File.Exists(listsPath);
      string profilePath = string.IsNullOrWhiteSpace(currentProfilePath)
              ? "(aucun profil chargé)"
              : currentProfilePath;

      return
          "Profile" +
          System.Environment.NewLine +
          "-------" +
          System.Environment.NewLine +
          "Path   : " + profilePath +
          System.Environment.NewLine +
          System.Environment.NewLine +
          "Profile information" +
          System.Environment.NewLine +
          "-------------------" +
          System.Environment.NewLine +
          "Name          : " + profile.ProfileName +
          System.Environment.NewLine +
          "Academic year : " + profile.AcademicYear +
          System.Environment.NewLine +
          System.Environment.NewLine +
          "Files" +
          System.Environment.NewLine +
          "-----" +
          System.Environment.NewLine +
          "Service workbook : " + GetStatusText(serviceExists,serviceConfigured) +
          System.Environment.NewLine +
          servicePath +
          System.Environment.NewLine +
          System.Environment.NewLine +
          "Lists workbook   : " + GetStatusText(listsExists,listsConfigured) +
          System.Environment.NewLine +
          listsPath;
    }

    private static string GetStatusText(bool exists,bool configured) {
      if (!configured)
        return "Not configured";

      if (!exists)
        return "Missing";

      return "OK";
    }
  }
}