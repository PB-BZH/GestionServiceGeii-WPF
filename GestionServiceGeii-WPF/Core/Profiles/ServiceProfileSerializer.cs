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
║  Nom de fichier : ServiceProfileSerializer.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.IO;
using System.Text;
using Newtonsoft.Json;

namespace GestionServiceGeii.Core.Profiles {
  public static class ServiceProfileSerializer {
    private static readonly Encoding Utf8NoBom =
        new UTF8Encoding(false);

    public static void Save(string path,ServiceManagerProfile profile) {
      if (string.IsNullOrWhiteSpace(path))
        throw new ArgumentException("Profile path is empty.",nameof(path));

      ArgumentNullException.ThrowIfNull(profile);

      string json =
          JsonConvert.SerializeObject(
              profile,
              Formatting.Indented
          );

      File.WriteAllText(
          path,
          json,
          Utf8NoBom
      );
    }

    public static ServiceManagerProfile Load(string path) {
      if (string.IsNullOrWhiteSpace(path))
        throw new ArgumentException("Profile path is empty.",nameof(path));

      if (!File.Exists(path))
        throw new FileNotFoundException("Profile file not found.",nameof(path));

      string json =
          File.ReadAllText(path);

      ServiceManagerProfile profile =
          JsonConvert.DeserializeObject<ServiceManagerProfile>(json)
          ?? throw new InvalidOperationException("Invalid service profile file.");

      EnsureDefaults(profile);

      return profile;
    }

    private static void EnsureDefaults(ServiceManagerProfile profile) {
      profile.Files ??= new ServiceFilesOptions();
      profile.Excel ??= new ServiceExcelOptions();
      profile.Display ??= new ServiceDisplayOptions();
      profile.Safety ??= new ServiceSafetyOptions();
    }
  }
}