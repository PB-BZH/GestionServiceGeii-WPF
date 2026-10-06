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
║  Nom de fichier : GestionIntervenantsForm.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using GestionServiceGeii.Core.Services;
using GestionServiceGeii.Shared.Database;
using GestionServiceGeii.Shared.Librairie_Fichier;
using PB.BZH.Theme.Theming;

namespace GestionServiceGeii.UI.Windows;

public partial class GestionIntervenants: Window {
  private const string TableTitulaires = "Titulaires";
  private const string ColonneTitulaires = "Titulaires";
  private const string ColonnePrenom = "Prenom";
  private const string ColonneStatut = "Statut";

  private const string TableVacataires = "Vacataires";
  private const string ColonneVacataires = "Vacataires";
  private static readonly IntervenantTableConfig TitulairesConfig = new(TableTitulaires,ColonneTitulaires,"Titulaire");
  private static readonly IntervenantTableConfig VacatairesConfig = new(TableVacataires,ColonneVacataires,"Vacataire");


  private readonly DataSet _dataSetExcelSelection;
  private readonly string _selectionWorkbookPath;

  private bool _updatingChecks;

  internal GestionIntervenants(DataSet dataSetExcelSelection,string selectionWorkbookPath) {

    InitializeComponent();

    _dataSetExcelSelection = dataSetExcelSelection;
    _selectionWorkbookPath = selectionWorkbookPath;

    ThemeManager.ApplyTheme(this);

    ChargerIntervenants();

    SetMessage("Prêt.",Brushes.LightGray);
  }

  private void ChargerIntervenants() {
    ChargerIntervenants(cmbTitulaires,TitulairesConfig);
    ChargerIntervenants(cmbVacataires,VacatairesConfig);
  }

  private void ChargerIntervenants(ComboBox comboBox,IntervenantTableConfig config) {
    comboBox.Items.Clear();
    DataTable? table = GetTable(config.TableName);
    if (table == null)
      return;
    if (!table.Columns.Contains(config.NameColumn))
      return;
    foreach (DataRow row in table.Rows) {
      string nom = ReadValue(row,config.NameColumn);
      if (string.IsNullOrWhiteSpace(nom))
        continue;
      string prenom = ReadValue(row,ColonnePrenom);
      string statut = ReadValue(row,ColonneStatut);
      if (string.IsNullOrWhiteSpace(statut) && config == VacatairesConfig) {
        statut = "Vacataire";
      }
      comboBox.Items.Add(new IntervenantListItem(nom,prenom,statut));
    }
  }

  private void chkTitulaire_CheckedChanged(object? sender,RoutedEventArgs e) {
    if (_updatingChecks)
      return;

    if (chkTitulaire.IsChecked != true)
      return;

    _updatingChecks = true;
    chkVacataire.IsChecked = false;
    cmbVacataires.SelectedItem = null;
    _updatingChecks = false;

    ChargerIntervenantSelectionne();
  }

  private void chkVacataire_CheckedChanged(object? sender,RoutedEventArgs e) {
    if (_updatingChecks)
      return;

    if (chkVacataire.IsChecked != true)
      return;

    _updatingChecks = true;
    chkTitulaire.IsChecked = false;
    cmbTitulaires.SelectedItem = null;
    _updatingChecks = false;

    ChargerIntervenantSelectionne();
  }

  private void cmbTitulaires_SelectionChanged(object? sender,SelectionChangedEventArgs e) {
    if (cmbTitulaires.SelectedItem == null)
      return;

    _updatingChecks = true;
    cmbVacataires.SelectedItem = null;
    chkTitulaire.IsChecked = true;
    chkVacataire.IsChecked = false;
    _updatingChecks = false;

    ChargerIntervenantSelectionne();
  }

  private void cmbVacataires_SelectionChanged(object? sender,SelectionChangedEventArgs e) {
    if (cmbVacataires.SelectedItem == null)
      return;

    _updatingChecks = true;
    cmbTitulaires.SelectedItem = null;
    chkTitulaire.IsChecked = false;
    chkVacataire.IsChecked = true;
    _updatingChecks = false;

    ChargerIntervenantSelectionne();
  }
  private void ChargerIntervenantSelectionne() {
    if (chkTitulaire.IsChecked == true) {
      if (cmbTitulaires.SelectedItem is not IntervenantListItem titulaire)
        return;
      txtName.Text = titulaire.Name;
      txtFirstName.Text = titulaire.FirstName;
      cmbStatus.Text = NormalizeStatus(titulaire.Status);
      SetMessage($"Titulaire chargé : {titulaire.DisplayName}",Brushes.LightSkyBlue);
      return;
    }

    if (chkVacataire.IsChecked == true) {
      if (cmbVacataires.SelectedItem is not IntervenantListItem vacataire)
        return;
      txtName.Text = vacataire.Name;
      txtFirstName.Text = vacataire.FirstName;
      cmbStatus.Text = "Vacataire";
      SetMessage($"Vacataire chargé : {vacataire.DisplayName}",Brushes.LightSkyBlue);
      return;
    }
  }

  private void btnAddUpdate_Click(object? sender,RoutedEventArgs e) {
    AjouterOuMettreAJourIntervenant();
  }

  private void AjouterOuMettreAJourIntervenant() {
    string nom = NormalizeName(txtName.Text);
    string prenom = txtFirstName.Text.Trim();
    string statut = NormalizeStatus(cmbStatus.Text);
    if (string.IsNullOrWhiteSpace(nom)) {
      SetMessage("Veuillez renseigner le nom de l'intervenant.",Brushes.Goldenrod);
      return;
    }

    if (string.IsNullOrWhiteSpace(statut)) {
      SetMessage("Veuillez sélectionner un statut.",Brushes.Goldenrod);
      return;
    }
    IntervenantTableConfig config = statut == "Vacataire" ? VacatairesConfig : TitulairesConfig;
    IntervenantTableConfig autreConfig = statut == "Vacataire" ? TitulairesConfig : VacatairesConfig;
    AjouterOuMettreAJourIntervenantDataSet(config,nom,prenom,statut);
    SupprimerIntervenantDansTable(autreConfig.TableName,autreConfig.NameColumn,nom);
    if (!SauvegarderIntervenantDansExcel(config,nom,prenom,statut,out bool ajoute)) {
      return;
    }
    SupprimerIntervenantDansExcel(autreConfig,nom);
    ChargerIntervenants();
    RegenererStatutsDepuisSelection(false);
    SetMessage($"{config.TypeLabel} {BuildDisplayName(prenom,nom)} {(ajoute ? "ajouté" : "mis à jour")}.",Brushes.LightGreen);
  }

  private bool SupprimerIntervenantDansExcel(IntervenantTableConfig config,string nom) {
    string cheminFichier = GetSelectionWorkbookPath();
    if (string.IsNullOrWhiteSpace(cheminFichier)) {
      SetMessage("Fichier de sélection indisponible : suppression impossible.",Brushes.IndianRed);
      return false;
    }
    try {
      ClasseEpplus.SupprimeLigneExcel_Epplus(cheminFichier: cheminFichier,nomFeuille: config.TableName,
          condition: (feuille,ligne) =>
              string.Equals(ClasseEpplus.GetStringByColumnName_Epplus(feuille,ligne,config.NameColumn).Trim(),nom,StringComparison.OrdinalIgnoreCase));
      return true;
    }
    catch (Exception ex) {
      SetMessage(
          $"Erreur suppression {config.TypeLabel.ToLower()} : {ex.Message}",
          Brushes.IndianRed);
      return false;
    }
  }

  private void AjouterOuMettreAJourIntervenantDataSet(IntervenantTableConfig config,string nom,string prenom,string statut) {
    DataTable? table = GetTable(config.TableName);
    if (table == null)
      return;
    if (!table.Columns.Contains(config.NameColumn))
      return;
    if (!table.Columns.Contains(ColonnePrenom))
      table.Columns.Add(ColonnePrenom);
    if (!table.Columns.Contains(ColonneStatut))
      table.Columns.Add(ColonneStatut);
    DataRow? row = FindRowByName(table,config.NameColumn,nom);
    if (row == null) {
      row = table.NewRow();
      table.Rows.Add(row);
    }
    row[config.NameColumn] = nom;
    row[ColonnePrenom] = prenom;
    row[ColonneStatut] = statut;
  }

  private bool SauvegarderIntervenantDansExcel(IntervenantTableConfig config,string nom,string prenom,string statut,out bool ajoute) {
    ajoute = false;
    string cheminFichier = GetSelectionWorkbookPath();
    if (string.IsNullOrWhiteSpace(cheminFichier)) {
      SetMessage("Fichier de sélection indisponible : sauvegarde impossible.",Brushes.IndianRed);
      return false;
    }
    try {
      int lignesModifiees = ClasseEpplus.MiseAJourCelluleExcel_Epplus(cheminFichier: cheminFichier,nomFeuille: config.TableName,
              condition: (feuille,ligne) =>
                  string.Equals(
                      ClasseEpplus.GetStringByColumnName_Epplus(
                          feuille,
                          ligne,
                          config.NameColumn).Trim(),
                      nom,
                      StringComparison.OrdinalIgnoreCase),
              actionMiseAJour: (feuille,ligne) => {
                ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligne,config.NameColumn,nom);
                ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligne,ColonnePrenom,prenom);
                ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligne,ColonneStatut,statut);
              });
      if (lignesModifiees == 0) {
        ClasseEpplus.AjouterLigneExcel_Epplus(
            cheminFichier,
            config.TableName,
            (feuille,ligne) => {
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligne,config.NameColumn,nom);
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligne,ColonnePrenom,prenom);
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligne,ColonneStatut,statut);
            });
        ajoute = true;
      }
      return true;
    }
    catch (Exception ex) {
      SetMessage($"Erreur sauvegarde {config.TypeLabel.ToLower()} : {ex.Message}",Brushes.IndianRed);
      return false;
    }
  }

  private void btnDeleteIntervenant_Click(object? sender,RoutedEventArgs e) {
    SupprimerIntervenantSelectionne();
  }

  private void SupprimerIntervenantSelectionne() {
    if (!chkTitulaire.IsChecked == true && !chkVacataire.IsChecked == true) {
      SetMessage("Choisir un type d'intervenant à supprimer.",Brushes.Goldenrod);
      return;
    }

    if (chkTitulaire.IsChecked == true) {
      if (cmbTitulaires.SelectedItem is not IntervenantListItem titulaire) {
        SetMessage("Aucun titulaire sélectionné.",Brushes.Goldenrod);
        return;
      }
      SupprimerIntervenantDansTable(TableTitulaires,ColonneTitulaires,titulaire.Name);
      if (!SupprimerTitulaireDansExcel(titulaire.Name))
        return;
      ChargerIntervenants();
      RegenererStatutsDepuisSelection(false);
      SetMessage($"Titulaire {titulaire.DisplayName} supprimé.",Brushes.LightGreen);
      return;
    }
    if (chkVacataire.IsChecked == true) {
      if (cmbVacataires.SelectedItem is not IntervenantListItem vacataire) {
        SetMessage("Aucun vacataire sélectionné.",Brushes.Goldenrod);
        return;
      }
      SupprimerIntervenantDansTable(TableVacataires,ColonneVacataires,vacataire.Name);
      if (!SupprimerVacataireDansExcel(vacataire.Name))
        return;
      ChargerIntervenants();
      RegenererStatutsDepuisSelection(false);
      SetMessage($"Vacataire {vacataire.DisplayName} supprimé.",Brushes.LightGreen);
    }
  }

  private bool SupprimerTitulaireDansExcel(string nom) {
    return SupprimerIntervenantDansExcel(TableTitulaires,ColonneTitulaires,nom,"titulaire");
  }

  private bool SupprimerVacataireDansExcel(string nom) {
    return SupprimerIntervenantDansExcel(TableVacataires,ColonneVacataires,nom,"vacataire");
  }

  private bool SupprimerIntervenantDansExcel(string tableName,string nameColumn,string nom,string typeIntervenant) {
    string cheminFichier = FichierDeSelection.Selection.CheminFichier;  //GetSelectionWorkbookPath();
    if (string.IsNullOrWhiteSpace(cheminFichier)) {
      SetMessage("Fichier de sélection indisponible : suppression impossible.",Brushes.IndianRed);
      return false;
    }
    try {
      ClasseEpplus.SupprimeLigneExcel_Epplus(cheminFichier: cheminFichier,nomFeuille: tableName,condition: (feuille,ligne) => string.Equals(
                  ClasseEpplus.GetStringByColumnName_Epplus(
                      feuille,
                      ligne,
                      nameColumn).Trim(),
                  nom,
                  StringComparison.OrdinalIgnoreCase
              ));
      return true;
    }
    catch (Exception ex) {
      SetMessage($"Erreur suppression {typeIntervenant} : {ex.Message}",Brushes.IndianRed);
      return false;
    }
  }

  private void SupprimerIntervenantDansTable(string tableName,string nameColumn,string nom) {
    DataTable? table = GetTable(tableName);
    if (table == null)
      return;
    DataRow? row = FindRowByName(table,nameColumn,nom);
    if (row == null)
      return;
    table.Rows.Remove(row);
  }

  private void btnGenerateStatus_Click(object? sender,RoutedEventArgs e) {
    RegenererStatutsDepuisSelection(true);
  }

  private void RegenererStatutsDepuisSelection(bool afficherMessage) {
    try {
      string sourceFile = _selectionWorkbookPath;
      if (string.IsNullOrWhiteSpace(sourceFile))
        sourceFile = "Fichier de sélection non renseigné";
      TeacherStatusesFile statusFile = FicheServiceStatusHelper.BuildFromSelectionDataSet(_dataSetExcelSelection,sourceFile);
      string jsonPath = FicheServiceStatusHelper.GetDefaultJsonPath();
      FicheServiceStatusHelper.Save(jsonPath,statusFile);
      if (afficherMessage) {
        SetMessage($"{statusFile.Teachers.Count} statut(s) mis à jour.",Brushes.LightGreen);
      }
    }
    catch (Exception ex) {
      SetMessage($"Erreur lors de la régénération des statuts : {ex.Message}",Brushes.IndianRed);
    }
  }

  private string GetSelectionWorkbookPath() {
    return _selectionWorkbookPath;
  }

  private DataTable? GetTable(string tableName) {
    if (!_dataSetExcelSelection.Tables.Contains(tableName))
      return null;
    return _dataSetExcelSelection.Tables[tableName];
  }

  private static DataRow? FindRowByName(DataTable table,string nameColumn,string nom) {
    if (!table.Columns.Contains(nameColumn))
      return null;
    string searched = NormalizeNameKey(nom);
    foreach (DataRow row in table.Rows) {
      string current = NormalizeNameKey(row[nameColumn]?.ToString() ?? string.Empty);
      if (current == searched)
        return row;
    }
    return null;
  }

  private static string ReadValue(DataRow row,string columnName) {
    if (!row.Table.Columns.Contains(columnName))
      return string.Empty;
    return row[columnName]?.ToString()?.Trim() ?? string.Empty;
  }

  private static string NormalizeName(string value) {
    return (value ?? string.Empty)
        .Trim()
        .ToUpperInvariant();
  }

  private static string NormalizeNameKey(string value) {
    return (value ?? string.Empty)
        .Trim()
        .ToUpperInvariant();
  }

  private static string NormalizeStatus(string value) {
    string statut = (value ?? string.Empty).Trim();

    return statut switch {
      "PR" => "PR",
      "MCF" => "MCF",
      "PRAG" => "PRAG",
      "PRCE" => "PRCE",
      "VACATAIRE" => "Vacataire",
      "Vacataire" => "Vacataire",
      _ => statut
    };
  }

  private sealed class IntervenantTableConfig {
    internal IntervenantTableConfig(string tableName,string nameColumn,string typeLabel) {
      TableName = tableName;
      NameColumn = nameColumn;
      TypeLabel = typeLabel;
    }
    internal string TableName { get; }
    internal string NameColumn { get; }
    internal string TypeLabel { get; }
  }

  private static string BuildDisplayName(string firstName,string name) {
    firstName = (firstName ?? string.Empty).Trim();
    name = (name ?? string.Empty).Trim();
    if (string.IsNullOrWhiteSpace(firstName))
      return name;
    return $"{name} {firstName}";
  }

  private void SetMessage(string message,Brush color) {
    txtMessage.Content = message;
    txtMessage.Foreground = color;
  }

  private void btnOK_Click(object sender,RoutedEventArgs e) {
    Close();
  }

  private sealed class IntervenantListItem {
    internal string Name { get; }
    internal string FirstName { get; }
    internal string Status { get; }
    internal string DisplayName => BuildDisplayName(FirstName,Name);

    internal IntervenantListItem(string name,string firstName,string status) {
      Name = NormalizeName(name);
      FirstName = firstName.Trim();
      Status = NormalizeStatus(status);
    }

    public override string ToString() {
      return DisplayName;
    }
  }
}