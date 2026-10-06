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
║  Nom de fichier : Configuration des coursForm.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using GestionServiceGeii.Core.Profiles;
using GestionServiceGeii.Shared.Database;
using GestionServiceGeii.Shared.Librairie_Fichier;
using GestionServiceGeii.Shared.Librairie_GEII;
using PB.BZH.Theme.Theming;

namespace GestionServiceGeii.UI.Windows;

public partial class ConfigurationDesCours: Window {
  private ServiceManagerProfile? _profile;
  private DataSet _dataSetExcelSélection;
  internal bool FicheModifiee { get; private set; }
  internal string ValeurCMValidee { get; private set; } = string.Empty;
  internal string ValeurTDValidee { get; private set; } = string.Empty;
  internal string ValeurTPValidee { get; private set; } = string.Empty;

  public ConfigurationDesCours(
      string nomModule,
      Label matière,
      DataGrid visualisationDonnées,
      ServiceManagerProfile profile,
      DataSet dataSetExcelSélection) {

    InitializeComponent();
    _dataSetExcelSélection = dataSetExcelSélection;
    ThemeManager.ApplyTheme(this);
    ModifierVolumeHoraire(
        nomModule,
        matière,
        visualisationDonnées,
        profile);
  }

  internal void AfficheRépartitionEtudiants() {
    NombreEtudiantsGeii1.Text = FormationGEII.Geii_1_FI.NbEtudiants.ToString();
    NbGroupeTD_Geii1.Text = FormationGEII.Geii_1_FI.NbGroupeTD.ToString();
    NbGroupeTP_Geii1.Text = FormationGEII.Geii_1_FI.NbGroupeTP.ToString();
  }

  public void AfficheFicheModuleCourant(DataGrid table) {
    foreach (DataRowView ligne in table.Items) {
      string typeDeCours = ClasseBaseDeDonnées.GetBoundRowValue_vue(ligne,ExcelSchemaNames.Columns.Cours);
      string durée = ClasseBaseDeDonnées.GetBoundRowValue_vue(ligne,ExcelSchemaNames.Columns.Duree);
      switch (typeDeCours) {
        case "CM":
          Valeur_CM_Module.Text = durée;
          break;

        case "TD":
          Valeur_TD_Module.Text = durée;
          break;

        case "TP":
          Valeur_TP_Module.Text = durée;
          break;
      }
    }
  }

  private void RestaurerCheminsClasseursDepuisProfil() {
    if (_profile == null || _profile.Files == null)
      return;

    if (
      (string.IsNullOrWhiteSpace(FichierDeSelection.Selection.CheminFichier) || !File.Exists(FichierDeSelection.Selection.CheminFichier)) &&
      !string.IsNullOrWhiteSpace(_profile.Files.ListsWorkbookPath) && File.Exists(_profile.Files.ListsWorkbookPath)
    )
      FichierDeSelection.Selection.ChargerDepuisChemin(_profile.Files.ListsWorkbookPath);


    if (
      (string.IsNullOrWhiteSpace(FichierDeService.Service.CheminFichier) || !File.Exists(FichierDeService.Service.CheminFichier)) &&
      !string.IsNullOrWhiteSpace(_profile.Files.ServiceWorkbookPath) && File.Exists(_profile.Files.ServiceWorkbookPath)
    )
      FichierDeService.Service.ChargerDepuisChemin(_profile.Files.ServiceWorkbookPath);
  }

  internal void ModifierVolumeHoraire(string NomModule,Label Matière,DataGrid dataGrid,ServiceManagerProfile profile) {
    _profile = profile;
    ConfigurationFicheMatière(NomModule,dataGrid);
    ClasseBaseDeDonnées.ChangementBaseDeDonnéesFicheModule(_dataSetExcelSélection,Matière,NouvelleValeurCM,NouvelleValeurTD,NouvelleValeurTP);
  }

  private string _codeCourtModule = string.Empty;

  internal void ConfigurationFicheMatière(string NomModule,DataGrid dataGrid) {
    AfficheFicheModuleCourant(dataGrid);
    AfficheRépartitionEtudiants();

    DataRowView? ligneCourante = dataGrid.SelectedItem as DataRowView;

    if (ligneCourante == null)
      return;

    string module = ClasseBaseDeDonnées.GetBoundRowValue_vue(ligneCourante,ExcelSchemaNames.Columns.Module);
    string libelleCourt = ClasseBaseDeDonnées.GetBoundRowValue_vue(ligneCourante,ExcelSchemaNames.Columns.LibelleCourt);
    string codeOse = ClasseBaseDeDonnées.GetBoundRowValue_vue(ligneCourante,ExcelSchemaNames.Columns.OSE);

    _codeCourtModule = ClasseBaseDeDonnées.ExtraireCodeCourtModuleDepuisLibelleCourt(libelleCourt);

    if (string.IsNullOrWhiteSpace(codeOse))
      codeOse = ClasseBaseDeDonnées.ConvertirCodeCourtEnCodeOse(_codeCourtModule);

    texteLibelléModule.Text = module;
    texteMatière.Content = module;
    txtLibelleCourt.Content = _codeCourtModule;
    txtCodeOSE.Content = codeOse;

    ValeurCM.Text = ClasseBaseDeDonnées.RechercheValeur(_dataSetExcelSélection,NomModule,ExcelSchemaNames.Columns.CM,ExcelSchemaNames.Tables.Module);
    ValeurTD.Text = ClasseBaseDeDonnées.RechercheValeur(_dataSetExcelSélection,NomModule,ExcelSchemaNames.Columns.TD,ExcelSchemaNames.Tables.Module);
    ValeurTP.Text = ClasseBaseDeDonnées.RechercheValeur(_dataSetExcelSélection,NomModule,ExcelSchemaNames.Columns.TP,ExcelSchemaNames.Tables.Module);
    ShowDialog();
  }

  private void AfficherValeursCourantes() {
    NouvelleValeurCM.Text = ValeurCM.Text;
    NouvelleValeurTD.Text = ValeurTD.Text;
    NouvelleValeurTP.Text = ValeurTP.Text;
  }

  private void MiseAjourNouvellesValeurs() {
    ValeurCM.Text = NouvelleValeurCM.Text;
    ValeurTD.Text = NouvelleValeurTD.Text;
    ValeurTP.Text = NouvelleValeurTP.Text;
  }

  private void ModificationIUT_Click(object sender,RoutedEventArgs e) {
    AfficherValeursCourantes();
    FenêtreCorrectionIUT.Visibility = Visibility.Visible;
  }

  private void MiseAJourIUT_Click(object sender,RoutedEventArgs e) {
    MiseAjourNouvellesValeurs();
    FenêtreCorrectionIUT.Visibility = Visibility.Collapsed;
    ClasseBaseDeDonnées.ChangementBaseDeDonnéesFicheModule(_dataSetExcelSélection,texteMatière,NouvelleValeurCM,NouvelleValeurTD,NouvelleValeurTP);
  }

  internal void ValiderficheMatière_Click(object sender,RoutedEventArgs e) {
    string module = texteMatière.Content.ToString()!.Trim();
    if (string.IsNullOrWhiteSpace(module)) {
      MessageBox.Show(this,"Aucun module n'est sélectionné.","Validation fiche matière",MessageBoxButton.OK,MessageBoxImage.Warning);
      return;
    }

    Stopwatch chronoTotal = Stopwatch.StartNew();

    try {
      Mouse.OverrideCursor = Cursors.Wait;
      IsEnabled = false;

      Stopwatch chrono = Stopwatch.StartNew();

      RestaurerCheminsClasseursDepuisProfil();
      Debug.WriteLine($"[CHRONO] Restaurer chemins : {chrono.ElapsedMilliseconds} ms");

      chrono.Restart();

      DialogueServiceGEII.MiseAJourModule_fichierDeServiceExcel(
        FichierDeSelection.Selection,
        DialogueServiceGEII.DataSetExcelSélection,
        ExcelSchemaNames.Tables.Module,
        module,true);

      Debug.WriteLine($"[CHRONO] Selection / Module : {chrono.ElapsedMilliseconds} ms");

      chrono.Restart();

      //ClasseBaseDeDonnées.MiseAJourDureesAttenduesGlobal_Epplus(
      //  FichierDeService.Service,
      //  _codeCourtModule,
      //  NouvelleValeurCM.Text,
      //  NouvelleValeurTD.Text,
      //  NouvelleValeurTP.Text);

      ClasseBaseDeDonnées.MiseAJourDureesAttenduesGlobal_OleDb(
          FichierDeService.Service,
          _codeCourtModule,
          NouvelleValeurCM.Text,
          NouvelleValeurTD.Text,
          NouvelleValeurTP.Text);

      Debug.WriteLine($"[CHRONO] Service / DUREE_ATTENDUE : {chrono.ElapsedMilliseconds} ms");

      ValeurCMValidee = NouvelleValeurCM.Text;
      ValeurTDValidee = NouvelleValeurTD.Text;
      ValeurTPValidee = NouvelleValeurTP.Text;

      FicheModifiee = true;
      DialogResult = true;

      Debug.WriteLine($"[CHRONO] TOTAL validation : {chronoTotal.ElapsedMilliseconds} ms");
    }
    catch (Exception ex) {
      MessageBox.Show(
        this,
        "Erreur pendant la validation de la fiche matière." +
        Environment.NewLine +
        Environment.NewLine +
        ex.Message,
        "Validation fiche matière",
        MessageBoxButton.OK,
        MessageBoxImage.Error);
    }
    finally {
      IsEnabled = true;
      Mouse.OverrideCursor = null;
    }
  }
}


