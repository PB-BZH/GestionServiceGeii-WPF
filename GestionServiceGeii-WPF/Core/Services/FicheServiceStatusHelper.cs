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
║  Nom de fichier : FicheServiceStatusHelper.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;

namespace GestionServiceGeii.Core.Services;

internal static class FicheServiceStatusHelper {
  private const string TableTitulaires = "Titulaires";
  private const string ColonneTitulaires = "Titulaires";
  private const string ColonneStatut = "Statut";

  private const string TableVacataires = "Vacataires";
  private const string ColonneVacataires = "Vacataires";
  private const string ColonnePrenom = "Prenom";

  private static readonly JsonSerializerOptions JsonOptions =
      new() {
        WriteIndented = true
      };

  internal static string GetDefaultJsonPath() {
    string directory =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PB-BZH Concept",
            "GestionServiceGeii");

    Directory.CreateDirectory(directory);

    return Path.Combine(
        directory,
        "teacher-statuses.json");
  }

  internal static TeacherStatusesFile BuildFromSelectionDataSet(DataSet dataSetSelection,string sourceFile) {
    Dictionary<string,TeacherStatusEntry> entries = new(StringComparer.OrdinalIgnoreCase);
    AjouterTitulaires(dataSetSelection,entries);
    AjouterVacataires(dataSetSelection,entries);
    return new TeacherStatusesFile {
      GeneratedAt = DateTime.Now,
      SourceFile = sourceFile,
      Teachers =
          entries
              .Values
              .OrderBy(entry => entry.Name)
              .ToList()
    };
  }

  internal static void Save(string jsonPath,TeacherStatusesFile statusFile) {
    string? directory = Path.GetDirectoryName(jsonPath);
    if (!string.IsNullOrWhiteSpace(directory)) {
      Directory.CreateDirectory(directory);
    }
    string json = JsonSerializer.Serialize(statusFile,JsonOptions);
    File.WriteAllText(jsonPath,json,Encoding.UTF8);
  }

  internal static TeacherStatusesFile Load(string jsonPath) {
    if (!File.Exists(jsonPath))
      return new TeacherStatusesFile();
    string json = File.ReadAllText(jsonPath,Encoding.UTF8);
    return JsonSerializer.Deserialize<TeacherStatusesFile>(json,JsonOptions) ?? new TeacherStatusesFile();
  }

  internal static bool TryGetStatus(TeacherStatusesFile statusFile,string teacherName,out string status) {
    status = string.Empty;
    string searchedKey = NormalizeNameKey(teacherName);
    if (string.IsNullOrWhiteSpace(searchedKey))
      return false;
    TeacherStatusEntry? entry = statusFile
      .Teachers
      .FirstOrDefault(item => NormalizeNameKey(item.Name) == searchedKey);
    if (entry == null || string.IsNullOrWhiteSpace(entry.Status)) {
      return false;
    }
    status = entry.Status.Trim();
    return true;
  }

  private static void AjouterTitulaires(DataSet dataSetSelection,Dictionary<string,TeacherStatusEntry> entries) {
    if (!dataSetSelection.Tables.Contains(TableTitulaires)) {
      MessageBox.Show($"Table '{TableTitulaires}' absente.");
      return;
    }

    DataTable? table = dataSetSelection.Tables[TableTitulaires];

    if (table == null)
      return;

    string colonnes = string.Join(Environment.NewLine,table.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
    MessageBox.Show($"Table : {table.TableName}\nLignes : {table.Rows.Count}\n\nColonnes :\n{colonnes}");

    if (!table.Columns.Contains(ColonneTitulaires) || !table.Columns.Contains(ColonneStatut)) {
      MessageBox.Show($"Colonne recherchée : '{ColonneTitulaires}' = {table.Columns.Contains(ColonneTitulaires)}\nColonne recherchée : '{ColonneStatut}' = {table.Columns.Contains(ColonneStatut)}");
      return;
    }

    foreach (DataRow row in table.Rows) {
      string name = ReadValue(row,ColonneTitulaires);
      string firstName = ReadValue(row,ColonnePrenom);
      string status = NormalizeStatus(ReadValue(row,ColonneStatut));
      AddEntry(entries,name,firstName,status,TableTitulaires,overwriteExisting: true);
    }
  }

  private static void AjouterVacataires(DataSet dataSetSelection,Dictionary<string,TeacherStatusEntry> entries) {
    if (!dataSetSelection.Tables.Contains(TableVacataires))
      return;
    DataTable? table = dataSetSelection.Tables[TableVacataires];
    if (table == null)
      return;
    if (!table.Columns.Contains(ColonneVacataires))
      return;
    foreach (DataRow row in table.Rows) {
      string name = ReadValue(row,ColonneVacataires);
      string firstName = ReadValue(row,ColonnePrenom);
      AddEntry(entries,name,firstName,"Vacataire",TableVacataires,overwriteExisting: false);
    }
  }

  private static void AddEntry(Dictionary<string,TeacherStatusEntry> entries,string name,string firstName,string status,string source,bool overwriteExisting) {
    name = NormalizeDisplayName(name);
    firstName = NormalizeFirstName(firstName);
    status = NormalizeStatus(status);
    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(status)) {
      return;
    }
    string key = NormalizeNameKey(name);
    if (string.IsNullOrWhiteSpace(key))
      return;
    if (!overwriteExisting && entries.ContainsKey(key)) {
      return;
    }
    entries[key] = new TeacherStatusEntry { Name = name,FirstName = firstName,DisplayName = BuildDisplayName(firstName,name),Status = status,Source = source };
  }

  internal static string GetDisplayName(TeacherStatusesFile statusFile,string teacherName) {
    string searchedKey = NormalizeNameKey(teacherName);
    if (string.IsNullOrWhiteSpace(searchedKey))
      return teacherName;
    TeacherStatusEntry? entry = statusFile.Teachers.FirstOrDefault(item => NormalizeNameKey(item.Name) == searchedKey);
    if (entry == null)
      return teacherName;
    if (!string.IsNullOrWhiteSpace(entry.DisplayName))
      return entry.DisplayName;
    if (!string.IsNullOrWhiteSpace(entry.FirstName))
      return $"{entry.FirstName.Trim()} {entry.Name.Trim()}";
    return entry.Name;
  }

  private static string NormalizeFirstName(string value) {
    return (value ?? string.Empty).Trim();
  }

  private static string BuildDisplayName(string firstName,string name) {
    firstName =
        (firstName ?? string.Empty).Trim();

    name =
        (name ?? string.Empty).Trim();

    if (string.IsNullOrWhiteSpace(firstName))
      return name;

    return $"{firstName} {name}";
  }

  private static string ReadValue(DataRow row,string columnName) {
    if (!row.Table.Columns.Contains(columnName))
      return string.Empty;

    return row[columnName]?.ToString()?.Trim() ?? string.Empty;
  }

  private static string NormalizeDisplayName(string value) {
    return (value ?? string.Empty).Trim().ToUpperInvariant();
  }

  private static string NormalizeStatus(string value) {
    string status = (value ?? string.Empty).Trim();
    return status switch {
      "PR" => "PR",
      "MCF" => "MCF",
      "PRAG" => "PRAG",
      "PRCE" => "PRCE",
      "Vacataire" => "Vacataire",
      "VACATAIRE" => "Vacataire",
      "PR / MCF" => "PR",
      "PRAG / PRCE" => "PRAG",
      _ => status
    };
  }

  private static string NormalizeNameKey(string value) {
    string normalized = (value ?? string.Empty).Trim().ToUpperInvariant().Normalize(NormalizationForm.FormD);
    StringBuilder builder = new();
    foreach (char character in normalized) {
      UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(character);
      if (category != UnicodeCategory.NonSpacingMark) {
        builder.Append(character);
      }
    }
    return builder.ToString().Normalize(NormalizationForm.FormC);
  }
}

internal sealed class TeacherStatusesFile {
  public DateTime GeneratedAt { get; set; }

  public string SourceFile { get; set; } = string.Empty;

  public List<TeacherStatusEntry> Teachers { get; set; } = [];
}

internal sealed class TeacherStatusEntry {
  public string Name { get; set; } = string.Empty;
  public string FirstName { get; set; } = string.Empty;
  public string DisplayName { get; set; } = string.Empty;
  public string Status { get; set; } = string.Empty;
  public string Source { get; set; } = string.Empty;
}