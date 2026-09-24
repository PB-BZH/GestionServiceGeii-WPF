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
║  Nom de fichier : ApplicationSettingsService.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.IO;
using System.Reflection;
using System.Text;
using Newtonsoft.Json;

namespace GestionServiceGeii.Core.Profiles {

  internal static class ApplicationSettingsService {
    private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

    internal static string GetSettingsDirectory() {
      string appData =
          Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

      return Path.Combine(
          appData,
          "PB-BZH Concept",
          "GestionServiceGeii",
          GetProductVersionFolderName());
    }

    internal static string GetSettingsPath() {
      return Path.Combine(
          GetSettingsDirectory(),
          "ApplicationSettings.json"
      );
    }

    internal static ApplicationSettings Load() {
      string path =
          GetSettingsPath();

      if (!File.Exists(path)) {
        TryMigrateLegacySettingsFile(path);

        if (!File.Exists(path))
          return new ApplicationSettings();
      }

      string json =
          File.ReadAllText(path);

      ApplicationSettings settings =
          JsonConvert.DeserializeObject<ApplicationSettings>(json)!;

      if (settings == null)
        return new ApplicationSettings();

      return settings;
    }

    internal static void Save(ApplicationSettings settings) {
      settings ??= new ApplicationSettings();

      string directory =
          GetSettingsDirectory();

      Directory.CreateDirectory(directory);

      string json =
          JsonConvert.SerializeObject(settings,Formatting.Indented);

      File.WriteAllText(
          GetSettingsPath(),
          json,
          Utf8NoBom
      );
    }

    private static string GetProductVersionFolderName() {
      string version = GetProductVersion();

      if (string.IsNullOrWhiteSpace(version))
        version = "UnknownVersion";

      version =
          SanitizePathSegment(version);

      return "v" + version;
    }

    private static string GetProductVersion() {
      Assembly assembly = Assembly.GetEntryAssembly() ?? typeof(ApplicationSettingsService).Assembly;
      string informationalVersion =
        assembly
        .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
        ?.InformationalVersion
        ?? string.Empty;
      if (!string.IsNullOrWhiteSpace(informationalVersion)) {
        int plusIndex = informationalVersion.IndexOf('+');
        if (plusIndex > 0)
          informationalVersion = informationalVersion.Substring(0,plusIndex);
        return informationalVersion.Trim();
      }
      Version? assemblyVersion = assembly.GetName().Version;
      if (assemblyVersion == null)
        return string.Empty;
      return
          assemblyVersion.Major + "." +
          assemblyVersion.Minor + "." +
          assemblyVersion.Build;
    }

    private static string SanitizePathSegment(string value) {
      if (string.IsNullOrWhiteSpace(value))
        return "UnknownVersion";

      string result =
          value.Trim();

      foreach (char invalidChar in Path.GetInvalidFileNameChars()) {
        result =
            result.Replace(invalidChar,'_');
      }

      return result;
    }

    private static string GetLegacySettingsDirectory() {
      string appData =
          Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

      return Path.Combine(
          appData,
          "PB-BZH Concept",
          "GestionServiceGeii");
    }

    private static string GetLegacySettingsPath() {
      return Path.Combine(
          GetLegacySettingsDirectory(),
          "ApplicationSettings.json");
    }

    private static void TryMigrateLegacySettingsFile(string newSettingsPath) {
      try {
        string legacyPath =
            GetLegacySettingsPath();

        if (!File.Exists(legacyPath))
          return;

        string? newDirectory =
            Path.GetDirectoryName(newSettingsPath);

        if (string.IsNullOrWhiteSpace(newDirectory))
          return;

        Directory.CreateDirectory(newDirectory);

        File.Copy(
            legacyPath,
            newSettingsPath,
            overwrite: false);
      }
      catch {
        // La migration des anciens settings ne doit jamais bloquer le démarrage.
      }
    }
  }
}