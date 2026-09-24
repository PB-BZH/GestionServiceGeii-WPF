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
║  Nom de fichier : UpdateChecker.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.IO;
using System.Reflection;
using System.Text.Json;

namespace GestionServicesGeii.Core.Services;

/// <summary>
///  
/// </summary>
public static class UpdateChecker {
  private static readonly JsonSerializerOptions JsonOptions =
      new() {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
      };

  /// <summary>
  /// 
  /// </summary>
  /// <param name="log"></param>
  /// <returns></returns>
  internal static async Task<UpdateCheckResult> GetUpdateStatusAsync(
        Action<string>? log = null
    ) {
    try {
      UpdateSettings? settings =
          ReadUpdateSettings();

      if (settings == null ||
          string.IsNullOrWhiteSpace(settings.UpdateManifestUrl) ||
          string.IsNullOrWhiteSpace(settings.ApplicationId)) {
        return new UpdateCheckResult {
          Success = false,
          Title = "Recherche de mise à jour",
          Message =
              "Les paramètres de mise à jour ne peuvent pas être lus." +
              Environment.NewLine +
              "Le fichier UpdateSettings.json est introuvable ou incomplet.",
          ErrorMessage = "UpdateSettings.json est introuvable ou incomplet."
        };
      }

      using HttpClient client = new();

      string json =
          await client.GetStringAsync(settings.UpdateManifestUrl);

      UpdateManifest? manifest =
          JsonSerializer.Deserialize<UpdateManifest>(
              json,
              JsonOptions
          );

      if (manifest == null ||
          string.IsNullOrWhiteSpace(manifest.Version)) {
        return new UpdateCheckResult {
          Success = false,
          Title = "Recherche de mise à jour",
          Message = "Les informations de mise à jour ne peuvent pas être lues.",
          ErrorMessage = "Le manifest de mise à jour est indisponible."
        };
      }

      if (!string.Equals(
              manifest.ApplicationId,
              settings.ApplicationId,
              StringComparison.OrdinalIgnoreCase)) {
        return new UpdateCheckResult {
          Success = false,
          Title = "Recherche de mise à jour",
          Message =
              "Le manifest de mmise à jour ne correspond pas à cette application." +
              Environment.NewLine +
              "Local application ID : " + settings.ApplicationId +
              Environment.NewLine +
              "Remote application ID: " + manifest.ApplicationId +
              Environment.NewLine +
              "Remote product name  : " + manifest.ProductName,
          ErrorMessage = "ApplicationId mismatch.",
          DownloadPage = manifest.DownloadPage,
          RemoteVersion = manifest.Version
        };
      }

      string localVersionText =
          GetApplicationDisplayVersion();

      Version localVersion =
          NormalizeVersionForCompare(localVersionText);

      Version remoteVersion =
          NormalizeVersionForCompare(manifest.Version);

      if (remoteVersion > localVersion) {
        return new UpdateCheckResult {
          Success = true,
          UpdateAvailable = true,
          Title = "Recherche de mise à  jour",
          Message =
              "Une nouvelle version est disponible.",
          LocalVersion = localVersionText,
          RemoteVersion = manifest.Version,
          DownloadPage = manifest.DownloadPage,
          ProductName = manifest.ProductName
        };
      }

      return new UpdateCheckResult {
        Success = true,
        UpdateAvailable = false,
        Title = "Recherche de mise à jour",
        Message =
            "Votre application est à jour.",
        LocalVersion = localVersionText,
        RemoteVersion = manifest.Version,
        DownloadPage = manifest.DownloadPage,
        ProductName = manifest.ProductName
      };
    }
    catch (Exception ex) {
      log?.Invoke("[ERROR] Echec de recherche de mise à jour : " + ex.Message);

      return new UpdateCheckResult {
        Success = false,
        Title = "Recheche de mise à jour",
        Message =
            "Imossible de mettre à jour." +
            Environment.NewLine +
            Environment.NewLine +
            ex.Message,
        ErrorMessage = ex.Message
      };
    }
  }

  private static UpdateSettings? ReadUpdateSettings() {
    string settingsPath =
        Path.Combine(
            AppContext.BaseDirectory,
            "UpdateSettings.json"
        );

    if (!File.Exists(settingsPath))
      return null;

    string json =
        File.ReadAllText(settingsPath);

    return JsonSerializer.Deserialize<UpdateSettings>(
        json,
        JsonOptions
    );
  }

  private static string GetApplicationDisplayVersion() {
    Assembly? assembly =
        Assembly.GetEntryAssembly();

    if (assembly == null)
      return "0.0.0";

    string? informationalVersion =
        assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion;

    if (!string.IsNullOrWhiteSpace(informationalVersion))
      return StripBuildMetadata(informationalVersion);

    Version? version =
        assembly.GetName().Version;

    if (version == null)
      return "0.0.0";

    return version.ToString();
  }

  private static string StripBuildMetadata(string version) {
    int plusIndex =
        version.IndexOf('+');

    if (plusIndex >= 0)
      version = version[..plusIndex];

    return version.Trim();
  }

  private static Version NormalizeVersionForCompare(string versionText) {
    if (string.IsNullOrWhiteSpace(versionText))
      return new Version(0,0,0,0);

    string cleanVersion =
        StripBuildMetadata(versionText)
            .Trim();

    if (!Version.TryParse(cleanVersion,out Version? version))
      return new Version(0,0,0,0);

    return new Version(
        version.Major,
        version.Minor,
        Math.Max(version.Build,0),
        Math.Max(version.Revision,0)
    );
  }

  private sealed class UpdateSettings {
    public string ApplicationId { get; set; } = "";
    public string UpdateManifestUrl { get; set; } = "";
  }

  private sealed class UpdateManifest {
    public string ApplicationId { get; set; } = "";
    public string ProductName { get; set; } = "";
    public string Version { get; set; } = "";
    public string Publisher { get; set; } = "";
    public string DownloadPage { get; set; } = "";
    public string MsiUrl { get; set; } = "";
    public string WebSetupUrl { get; set; } = "";
    public string UpdateManifestUrl { get; set; } = "";
    public string ReleaseDate { get; set; } = "";
  }


  internal sealed class UpdateCheckResult {
    public bool Success { get; set; }
    public bool UpdateAvailable { get; set; }
    public string Title { get; set; } = "Check for updates";
    public string Message { get; set; } = string.Empty;
    public string ProductName { get; set; } = string.Empty;
    public string LocalVersion { get; set; } = string.Empty;
    public string RemoteVersion { get; set; } = string.Empty;
    public string DownloadPage { get; set; } = string.Empty;
    public string ErrorMessage { get; set; } = string.Empty;
  }
}
