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
║  Le 17/6/2026 - 11:30
╟────────────────────────────────────────────────────────────────────────────────║
║     Projet Visual Studio Professional 2026 : GestionServicesGEII
╟────────────────────────────────────────────────────────────────────────────────║
║     Version : 1.5.3
╟────────────────────────────────────────────────────────────────────────────────║
║                Visual Studio Professional 2026 - Insiders                      ║
║                ──────────────────────────────────────────                      ║
║  Langage     : C# 14                                                           ║
║  Technologie : .NET 10 / WinForms                                              ║
║  Plateforme  : Windows 10 / Windows 11                                         ║
║  Encodage    : UTF-8                                                           ║
╟────────────────────────────────────────────────────────────────────────────────║
║  Nom de fichier : GenerateServiceSheetsForm.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using GestionServiceGeii.Core.Profiles;
using GestionServiceGeii.Core.Services;
using GestionServiceGeii.Shared.Database;
using GestionServiceGeii.Shared.Librairie_Fichier;
using Microsoft.Win32;
using PB.BZH.Theme.Theming;

namespace GestionServiceGeii.UI.Windows;

public partial class GenerateServiceSheets: Window {
  private readonly DataSet _dataSetExcelSelection;
  private readonly DataSet _dataSetExcelService;
  private readonly ClasseExcel? _selectionWorkbook;
  private readonly ClasseExcel? _serviceWorkbook;
  private readonly ServiceManagerProfile _profile;
  private readonly string _academicYear;

  internal GenerateServiceSheets(
    DataSet dataSetExcelSelection,
    DataSet dataSetExcelService,
    ClasseExcel? selectionWorkbook,
    ClasseExcel? serviceWorkbook,
    ServiceManagerProfile profile) {

    InitializeComponent();
    ArgumentNullException.ThrowIfNull(profile);
    _profile = profile;

    ThemeManager.ApplyTheme(this);

    _dataSetExcelSelection = dataSetExcelSelection;
    _dataSetExcelService = dataSetExcelService;
    _academicYear = profile.AcademicYear;
    _selectionWorkbook = selectionWorkbook;
    _serviceWorkbook = serviceWorkbook;
    ApplyProfileToUi();
  }

  private void SetMessage(string message,Brush color) {
    txtMessage.Content = message;
    txtMessage.Foreground = color;
  }

  private void ApplyProfileToUi() {
    txtTargetDirectory.Text = _profile.Files.ExportPdfDirectory;
  }

  private void ApplyUiToProfile() {
    _profile.Files.ExportPdfDirectory = txtTargetDirectory.Text.Trim();
  }

  private void SelectDirectory(TextBox targetTextBox,string description) {
    OpenFolderDialog dialog = new();
    dialog.Title = description;

    if (!string.IsNullOrWhiteSpace(targetTextBox.Text) &&
        Directory.Exists(targetTextBox.Text)) {
      dialog.FolderName = targetTextBox.Text;
    }

    if (dialog.ShowDialog(this) != true)
      return;

    targetTextBox.Text = dialog.FolderName;
  }

  private void GenererToutesLesFichesPdf() {
    string dossierDestination = txtTargetDirectory.Text.Trim();
    if (string.IsNullOrWhiteSpace(dossierDestination)) {
      SetMessage("Veuillez sélectionner un dossier de destination.",Brushes.IndianRed);
      return;
    }
    Directory.CreateDirectory(dossierDestination);
    NettoyerAnciennesFichesPdf(dossierDestination);
    string jsonPath =
        FicheServiceStatusHelper.GetDefaultJsonPath();
    if (!File.Exists(jsonPath)) {
      SetMessage("Le fichier des statuts n'existe pas encore. Lancez d'abord la mise à jour des statuts.",Brushes.IndianRed);
      return;
    }
    TeacherStatusesFile statusFile = FicheServiceStatusHelper.Load(jsonPath);
    if (statusFile.Teachers.Count == 0) {
      SetMessage("Le fichier des statuts est vide ou invalide. Appeler Generation des fiches",Brushes.Goldenrod);
    }
    if (!_dataSetExcelService.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal)) {
      SetMessage("La table complète n'est pas disponible.",Brushes.IndianRed);
      return;
    }
    DataTable nomTableGlobal = _dataSetExcelService.Tables[ExcelSchemaNames.Tables.NomTableGlobal]!;
    List<string> intervenants =
        FicheServiceHelper
            .GetIntervenantsDepuisTableComplete(nomTableGlobal)
            .OrderBy(nom => nom)
            .ToList();

    int fichiersGeneres = 0;
    List<string> intervenantsSansStatut = [];
    List<string> erreurs = [];

    foreach (string intervenant in intervenants) {
      try {
        if (!FicheServiceStatusHelper.TryGetStatus(
                statusFile,
                intervenant,
                out string statut)) {

          intervenantsSansStatut.Add(intervenant);
          continue;
        }
        RegleStatutService regle = FicheServiceHelper.ObtenirRegleStatutService(statut);
        List<LigneServicePrevisionnel> lignes = FicheServiceHelper.ConstruireLignesServicePrevisionnel(nomTableGlobal,intervenant,regle)
                .Values
                .OrderBy(ligne => FicheServiceHelper.ObtenirOrdreSemestre(ligne.Semestre))
                .ThenBy(ligne => ligne.Ose)
                .ThenBy(ligne => ligne.Module)
                .ToList();
        if (lignes.Count == 0)
          continue;
        string nomFichier = ConstruireNomFichierPdf(intervenant);
        string cheminPdf = Path.Combine(dossierDestination,nomFichier);
        string titre = $"SERVICE PRÉVISIONNEL {_academicYear}";
        string intervenantAffiche = FicheServiceStatusHelper.GetDisplayName(statusFile,intervenant);
        FicheServicePdfHelper.ExporterPdf(cheminPdf,titre,intervenantAffiche,statut,lignes,regle);
        fichiersGeneres++;
      }
      catch (Exception ex) {
        MessageBox.Show(
            this,
            $"Erreur lors de la génération de la fiche de {intervenant} :\n\n{ex}",
            "Erreur QuestPDF",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        return;
      }
    }
    AfficherBilanGenerationPdf(fichiersGeneres,intervenantsSansStatut,erreurs);
  }

  private static string ConstruireNomFichierPdf(string intervenant) {
    string nom =
        intervenant
            .Trim()
            .ToUpperInvariant()
            .Replace(" ","_");
    foreach (char invalidChar in Path.GetInvalidFileNameChars()) {
      nom = nom.Replace(invalidChar,'_');
    }
    return $"Service_previsionnel_{nom}.pdf";
  }

  private static void NettoyerAnciennesFichesPdf(string dossierDestination) {
    if (string.IsNullOrWhiteSpace(dossierDestination))
      return;

    if (!Directory.Exists(dossierDestination))
      return;

    foreach (string fichierPdf in Directory.GetFiles(
                 dossierDestination,
                 "Service_previsionnel_*.pdf")) {
      File.Delete(fichierPdf);
    }
  }

  private void GenererStatutsDepuisSelection() {
    try {
      string sourceFile = _selectionWorkbook?.CheminFichier ?? string.Empty;
      if (string.IsNullOrWhiteSpace(sourceFile))
        sourceFile = "Fichier de sélection non renseigné";
      TeacherStatusesFile statusFile = FicheServiceStatusHelper.BuildFromSelectionDataSet(_dataSetExcelSelection,sourceFile);
      string jsonPath = FicheServiceStatusHelper.GetDefaultJsonPath();
      FicheServiceStatusHelper.Save(jsonPath,statusFile);
      SetMessage($"{statusFile.Teachers.Count} statut(s) intervenant(s) généré(s).",Brushes.GreenYellow);
    }
    catch (Exception ex) {
      SetMessage($"Erreur lors de la génération des statuts : {ex.Message}",Brushes.IndianRed);
    }
  }

  private void AfficherBilanGenerationPdf(int fichiersGeneres,List<string> intervenantsSansStatut,List<string> erreurs) {
    StringBuilder message = new();
    message.AppendLine($"{fichiersGeneres} fiche(s) PDF générée(s)");
    SetMessage($"{fichiersGeneres} fiche(s) PDF générée(s)",Brushes.GreenYellow);
    if (intervenantsSansStatut.Count > 0) {
      message.AppendLine();
      SetMessage("Intervenant(s) ignoré(s), statut introuvable :",Brushes.IndianRed);
      foreach (string intervenant in intervenantsSansStatut) {
        message.AppendLine($"- {intervenant}");
      }
    }
    if (erreurs.Count > 0) {
      message.AppendLine();
      message.AppendLine("Erreur(s) :");
      foreach (string erreur in erreurs) {
        message.AppendLine($"- {erreur}");
      }
    }
  }

  private void btnOk_Click(object sender,RoutedEventArgs e) {
    ApplyUiToProfile();
    Close();
  }

  private void btnOpenTargetDirectory_Click(object sender,RoutedEventArgs e) {
    if (txtTargetDirectory.Text == string.Empty) {
      SetMessage("Veuillez définir le chemin de destination",Brushes.Red);
    }
    else
      Process.Start(new ProcessStartInfo {
        FileName = txtTargetDirectory.Text,
        UseShellExecute = true
      });
  }

  private void btnBrowseTargetDirectory_Click(object sender,RoutedEventArgs e) {
    SelectDirectory(txtTargetDirectory,"Sélectionner le dossier des fiches de services des intervenants");
  }

  private void btnStartUpdate_Click(object sender,RoutedEventArgs e) {
    GenererStatutsDepuisSelection();
  }

  private void btnGenerateServiceSheets_Click(object sender,RoutedEventArgs e) {
    GenererToutesLesFichesPdf();
    SetMessage(txtMessage.Content + " - " + txtTargetDirectory.Text,Brushes.GreenYellow);
  }
}
