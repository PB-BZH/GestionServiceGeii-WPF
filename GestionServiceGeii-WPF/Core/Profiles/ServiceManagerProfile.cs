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
║  Nom de fichier : ServiceManagerProfile.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Drawing;
using GestionServiceGeii.Core.Enums;
using Newtonsoft.Json;

namespace GestionServiceGeii.Core.Profiles {
  public sealed class ServiceManagerProfile {
    public string ProfileName { get; set; } = "GEII";
    public string AcademicYear { get; set; } = "2026-2027";
    public ServiceGroupsOptions Groups { get; set; } = new ServiceGroupsOptions();
    public ServiceFilesOptions Files { get; set; } = new ServiceFilesOptions();
    public ServiceExcelOptions Excel { get; set; } = new ServiceExcelOptions();
    public ServiceDisplayOptions Display { get; set; } = new ServiceDisplayOptions();
    public ServiceSafetyOptions Safety { get; set; } = new ServiceSafetyOptions();
    public ProductOptions Product { get; set; } = new();
    public UpdateManifest UpdateManifest { get; set; } = new();
    public SaveState SaveState { get; set; } = new();
    public SaveWindowState Window { get; set; } = new();
  }

  public sealed class SaveState {
    public string CurrentSemester { get; set; } = string.Empty;
    public string CurrentFormation { get; set; } = string.Empty;
    public string CurrentModule { get; set; } = string.Empty;
    public string CurrentCourse { get; set; } = string.Empty;
  }

  public sealed class UpdateManifest {
    public string ProductName { get; set; } = "";
    public string Version { get; set; } = "";
    public string Publisher { get; set; } = "";
    public string DownloadPage { get; set; } = "https://www.pb-bzh-concept.fr";
    public string PrivacyPage { get; set; } = "https://www.pb-bzh-concept.fr/privacy.php";
    public string MsiUrl { get; set; } = "";
    public string WebSetupUrl { get; set; } = "";
    public string UpdateManifestUrl { get; set; } = "";
    public string ReleaseDate { get; set; } = "";
    public string ApplicationId { get; set; } = "";
  }

  public sealed class ProductOptions {
    public string ProductName { get; set; } = "Gestion Service GEII";
    public string ProductId { get; set; } = "GestionServiceGEII";
    public string Manufacturer { get; set; } = "PB BZH Concept";
    public string Version { get; set; } = "1.4.0";
    public string Description { get; set; } = "Gestion du service des intervenants pour les étudiants du GEII";
    public string UpgradeCode { get; set; } = "";
    public string IconPath { get; set; } = "";
    [JsonIgnore]
    public Image? LogoImage { get; set; } = GestionServiceGeii.Properties.Resources.Application;
    public string DownloadPageUrl { get; set; } = "https://www.pb-bzh-concept.fr";
    public string PrivacyPageUrl { get; set; } = "https://www.pb-bzh-concept.fr/privacy.php";
    public string Copyright { get; set; } = "© Copyright PB BZH Concept 2026";
    public string EmailContact { get; set; } = "admin@pb-bzh-concept.fr";
    public string PoductId { get; set; } = "GestionServiceGEII";
    public string DownloadCategory { get; set; } = "msi-software-packager";
  }

  public sealed class ServiceGroupsOptions {
    public bool IsConfigured { get; set; }
    public string FormationName { get; set; } = "GEII 1";
    public int Geii1_StudentCount { get; set; }
    public int Geii1_TdGroupCount { get; set; }
    public int Geii1_TpGroupCount { get; set; }
    public int MaxStudentsPerTpGroup { get; set; }
    public string[] Geii1_TdGroupNames { get; set; } = [];
    public string[] Geii1_TpGroupNames { get; set; } = [];
    public int Geii2_StudentCount { get; set; }
    public int Geii2_TdGroupCount { get; set; }
    public int Geii2_TpGroupCount { get; set; }
    public string[] Geii2_TdGroupNames { get; set; } = [];
    public string[] Geii2_TpGroupNames { get; set; } = [];
  }

  public sealed class ServiceFilesOptions {
    public string ServiceWorkbookPath { get; set; } = string.Empty;
    public string ListsWorkbookPath { get; set; } = string.Empty;
    public string BackupDirectory { get; set; } = string.Empty;
    public string ExportDirectory { get; set; } = string.Empty;
    public string ExportPdfDirectory { get; set; } = string.Empty;
  }

  public sealed class ServiceExcelOptions {
    public string ServiceSheetName { get; set; } = "Services";
    public string TeachersSheetName { get; set; } = "Enseignants";
    public string ModulesSheetName { get; set; } = "Modules";
    public string GroupsSheetName { get; set; } = "Groupes";
  }

  public sealed class ServiceDisplayOptions {
    public bool RememberLastFiles { get; set; } = true;
    public bool AutoResizeColumns { get; set; } = true;
    public string LastSelectedTeacher { get; set; } = string.Empty;
  }

  public sealed class ServiceSafetyOptions {
    public bool CreateBackupBeforeSave { get; set; } = true;
    public int BackupRetentionCount { get; set; } = 20;
    public bool ReadOnlyIfExcelFileLocked { get; set; } = true;
  }
}