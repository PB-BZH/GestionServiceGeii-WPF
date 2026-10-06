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
║  Nom de fichier : ServiceWorkbookDiskCache.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Data;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using GestionServiceGeii.Shared.Database;


namespace GestionServiceGeii.Core.Services;

internal static class ServiceWorkbookDiskCache {
  private const string CacheVersion =
      "GlobalServiceCache-v1";

  internal static bool TryLoadGlobalTable(
      string serviceWorkbookPath,
      out DataTable? tableGlobal
  ) {
    tableGlobal = null;

    try {
      string cachePath = BuildGlobalCachePath(serviceWorkbookPath);
      if (string.IsNullOrWhiteSpace(cachePath))
        return false;
      if (!File.Exists(cachePath))
        return false;
      DataSet dataSet = new DataSet();
      dataSet.ReadXml(cachePath,XmlReadMode.ReadSchema);
      if (!dataSet.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal))
        return false;
      tableGlobal = dataSet.Tables[ExcelSchemaNames.Tables.NomTableGlobal]!.Copy();
      tableGlobal.TableName = ExcelSchemaNames.Tables.NomTableGlobal;
      return true;
    }
    catch {
      tableGlobal = null;
      return false;
    }
  }

  internal static bool TryLoadServiceDataSet(string serviceWorkbookPath,out DataSet? serviceDataSet) {
    serviceDataSet = null;

    try {
      string cachePath = BuildServiceDataSetCachePath(serviceWorkbookPath);
      if (string.IsNullOrWhiteSpace(cachePath)) {
        return false;
      }
      if (!File.Exists(cachePath)) {
        return false;
      }
      DataSet dataSet = new();
      dataSet.ReadXml(cachePath,XmlReadMode.ReadSchema);
      if (!dataSet.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal)) {
        return false;
      }
      serviceDataSet = dataSet;
      return true;
    }
    catch (Exception ex) {
      serviceDataSet = null;
      return false;
    }
  }

  internal static void SaveServiceDataSet(
    string serviceWorkbookPath,
    DataSet serviceDataSet
) {
    if (string.IsNullOrWhiteSpace(serviceWorkbookPath))
      return;

    if (serviceDataSet == null)
      return;

    try {
      string cachePath = BuildServiceDataSetCachePath(serviceWorkbookPath);
      if (string.IsNullOrWhiteSpace(cachePath))
        return;
      string? cacheDirectory = Path.GetDirectoryName(cachePath);
      if (string.IsNullOrWhiteSpace(cacheDirectory))
        return;
      Directory.CreateDirectory(cacheDirectory);
      DataSet copy = serviceDataSet.Copy();
      copy.DataSetName = "GestionServiceGeii_ServiceWorkbookCache";
      copy.WriteXml(cachePath,XmlWriteMode.WriteSchema);
    }
    catch {
      // Le cache ne doit jamais empêcher le chargement normal.
    }
  }

  private static string BuildServiceDataSetCachePath(string serviceWorkbookPath) {
    if (string.IsNullOrWhiteSpace(serviceWorkbookPath))
      return string.Empty;
    FileInfo fileInfo = new FileInfo(serviceWorkbookPath);
    if (!fileInfo.Exists)
      return string.Empty;
    string identity =
        "ServiceWorkbookDataSetCache-v1" +
        "|" +
        fileInfo.FullName.ToUpperInvariant() +
        "|" +
        fileInfo.LastWriteTimeUtc.Ticks +
        "|" +
        fileInfo.Length;
    string hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity)));
    string cacheDirectory =
      Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "GestionServiceGeii",
            "Cache",
            "ServiceWorkbook");

    return Path.Combine(cacheDirectory,"ServiceDataSet_" + hash + ".xml");
  }

  internal static void SaveGlobalTable(string serviceWorkbookPath,DataTable tableGlobal) {
    if (string.IsNullOrWhiteSpace(serviceWorkbookPath))
      return;
    if (tableGlobal == null)
      return;
    try {
      string cachePath = BuildGlobalCachePath(serviceWorkbookPath);
      if (string.IsNullOrWhiteSpace(cachePath))
        return;
      string? cacheDirectory = Path.GetDirectoryName(cachePath);
      if (string.IsNullOrWhiteSpace(cacheDirectory))
        return;
      Directory.CreateDirectory(cacheDirectory);
      DataSet dataSet = new DataSet("GestionServiceGeii_ServiceCache");
      DataTable copy = tableGlobal.Copy();
      copy.TableName = ExcelSchemaNames.Tables.NomTableGlobal;
      dataSet.Tables.Add(copy);
      dataSet.WriteXml(cachePath,XmlWriteMode.WriteSchema);
    }
    catch {
      // Le cache ne doit jamais empêcher le chargement normal.
    }
  }

  internal static void DeleteGlobalCache(string serviceWorkbookPath) {
    try {
      string cachePath = BuildGlobalCachePath(serviceWorkbookPath);
      if (!string.IsNullOrWhiteSpace(cachePath) && File.Exists(cachePath))
        File.Delete(cachePath);
    }
    catch {
      // Suppression non critique.
    }
  }

  private static string BuildGlobalCachePath(string serviceWorkbookPath) {
    if (string.IsNullOrWhiteSpace(serviceWorkbookPath))
      return string.Empty;
    FileInfo fileInfo = new FileInfo(serviceWorkbookPath);
    if (!fileInfo.Exists)
      return string.Empty;
    string identity =
        CacheVersion +
        "|" +
        fileInfo.FullName.ToUpperInvariant() +
        "|" +
        fileInfo.LastWriteTimeUtc.Ticks +
        "|" +
        fileInfo.Length +
        "|" +
        ExcelSchemaNames.Tables.NomTableGlobal;
    string hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity)));
    string cacheDirectory =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "GestionServiceGeii",
            "Cache",
            "ServiceWorkbook");
    return Path.Combine(cacheDirectory,"Global_" + hash + ".xml");
  }



  internal static int ClearServiceWorkbookCache() {
    int deletedFiles = 0;

    try {
      string cacheDirectory = GetServiceWorkbookCacheDirectory();

      if (string.IsNullOrWhiteSpace(cacheDirectory))
        return deletedFiles;

      if (!Directory.Exists(cacheDirectory))
        return deletedFiles;

      string[] cacheFiles = Directory.GetFiles(cacheDirectory,"*.xml",SearchOption.TopDirectoryOnly);

      foreach (string cacheFile in cacheFiles) {
        try {
          File.Delete(cacheFile);
          deletedFiles++;
        }
        catch {
          // Un fichier verrouillé ne doit pas bloquer la suppression des autres.
        }
      }
    }
    catch {
      // Le vidage du cache ne doit jamais faire planter l'application.
    }

    return deletedFiles;
  }


  private static string GetServiceWorkbookCacheDirectory() {
    return Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "GestionServiceGeii",
        "Cache",
        "ServiceWorkbook");
  }

  internal static int InvalidateServiceWorkbookCache() {
    return ClearServiceWorkbookCache();
  }
}

