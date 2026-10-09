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
║  Le 24/9/2026 - 00:11
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
║  Nom de fichier : DialogueServiceGEII.xaml.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using GestionServiceGeii.Core.Enums;
using GestionServiceGeii.Core.Groups;
using GestionServiceGeii.Core.Models;
using GestionServiceGeii.Core.Profiles;
using GestionServiceGeii.Core.Services;
using GestionServiceGeii.Shared.Database;
using GestionServiceGeii.Shared.Librairie_Fichier;
using GestionServiceGeii.Shared.Librairie_GEII;
using GestionServiceGeii.Shared.Librairie_Générique;
using GestionServiceGeii.UI.Printing;
using GestionServiceGeii.UI.Windows;
using Microsoft.Win32;
using PB.BZH.Theme.Theming;
using static GestionServiceGeii.Shared.Database.ClasseBaseDeDonnées;
using Brushes = System.Windows.Media.Brushes;
using DataTable = System.Data.DataTable;
using Path = System.IO.Path;

namespace GestionServiceGeii.UI;
/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class DialogueServiceGEII: Window {
  private string ancienneDonnée = string.Empty;
  private static FichierDeSelection selection = new();
  private ClasseGénérique générique = new();
  private static readonly FormationGEII formationGEII = new();
  private static ConfigurationEffectifs_GEII_1? _configurationEffectifs_GEII_1 { get; set; }
  private static ConfigurationEffectifs_GEII_2? _configurationEffectifs_GEII_2 { get; set; }
  private static ConfigurationDesCours? _configurationDesCours { get; set; }
  private static ChoixTypeDeCours _choixTypeDeCours = new();
  private static readonly GestionDesGroupes _gestionDesGroupes = new();
  private static DataSet _dataSetExcelSélection = new("FICHIER DE SELECTION");
  private static DataSet _dataSetExcelService = new("FICHIER DE SERVICE");
  private static DataTable _dataTableExcel = new("Table ExcelApp");
  private static ComboBox noms = new();
  private static string? choix;
  private static string? nomColonne;
  private static int indexLigne_vue;
  internal static string[] NomGroupeTD_Geii1 = new string[4];
  internal static string[] NomGroupeTP_Geii1 = new string[8];

  internal static string[] NomGroupeTD_Geii2_o_FI = new string[3];
  internal static int nbGroupeTD_officiel_Geii2_FI;
  internal static string[] NomGroupeTD_Geii2_officiel_FI = new string[3];
  internal static string[] NomGroupeTP_Geii2_FI = new string[6];
  internal static string[] NomGroupeTPSp_Geii2_FI = new string[6];

  internal static string[] NomGroupeTD_Geii2_FA = new string[3];
  internal static string[] NomGroupeTD_Geii2_FI = new string[3];
  internal static int nbGroupeTD_officiel_Geii2_FA;
  internal static string[] NomGroupeTD_Geii2_officiel_FA = new string[3];
  internal static string[] NomGroupeTP_Geii2_FA = new string[6];
  internal static string[] NomGroupeTPSp_Geii2_FA = new string[6];

  private ServiceManagerProfile _profile = new();
  private string _currentProfilePath = string.Empty;
  private bool _serviceWorkbookLoaded = false;
  private bool _isLoadingProfile = false;
  private bool _selectionWorkbookLoaded = false;
  private bool _isInitializingView = false;
  private bool _exitRequestedFromMenu;
  private ClasseExcel? SelectionWorkbook { get; set; }
  private ClasseExcel? ServiceWorkbook { get; set; }
  private readonly ApplicationSettings _applicationSettings = new();
  private readonly string _traceSessionId = DateTime.Now.ToString("yyyyMMdd_HHmmss");
  private DispatcherTimer? _refreshModuleViewTimer;
  private bool _suspendViewRefresh;
  private int _traceSequence;
  private bool _isRefreshingInfos;
  private bool _isRefreshingOse;
  private readonly List<ClasseBaseDeDonnées.ExcelSourceCellChange> _pendingSourceChanges = new();
  private readonly Dictionary<string,ClasseBaseDeDonnées.ExcelSourceRowRewrite> _pendingSourceRowRewrites = new(StringComparer.OrdinalIgnoreCase);
  private string _ancienCoursAvantModification = string.Empty;
  private string _ancienGroupeAvantModification = string.Empty;


  private string AncienneDonnée {
    get => ancienneDonnée;
    set => ancienneDonnée = value ?? ancienneDonnée;
  }

  private static FichierDeSelection Selection {
    get => selection;
    set => selection = value ?? selection;
  }

  private ClasseGénérique Générique {
    get => générique;
    set => générique = value ?? générique;
  }

  internal static DataSet DataSetExcelSélection {
    get => _dataSetExcelSélection;
    set => _dataSetExcelSélection = value ?? _dataSetExcelService;
  }

  internal static DataSet DataSetExcelService {
    get => _dataSetExcelService;
    set => _dataSetExcelService = value ?? _dataSetExcelService;
  }

  private static ComboBox Noms {
    get => noms;
    set => noms = value ?? noms;
  }

  private static string? Choix {
    get => choix;
    set => choix = value ?? choix;
  }

  private static string? NomColonne {
    get => nomColonne;
    set => nomColonne = value ?? nomColonne;
  }

  private static int IndexLigne_vue {
    get => indexLigne_vue;
    set => indexLigne_vue = value;
  }

  protected Grid BoutonsDeModification {
    get => boutonsDeModification;
    set => boutonsDeModification = value ?? boutonsDeModification;
  }

  private Label TexteMatière {
    get => texteMatière;
    set => texteMatière = value ?? texteMatière;
  }

  private Label TexteCodeOse {
    get => txtCodeOSE;
    set => txtCodeOSE = value ?? txtCodeOSE;
  }

  private TextBox TexteLibelléModule {
    get => texteLibelléModule;
    set => texteLibelléModule = value ?? texteLibelléModule;
  }

  private Label Valeur_CM {
    get => valeur_CM;
    set => valeur_CM = value ?? Valeur_CM;
  }

  private Label Valeur_TD {
    get => valeur_TD;
    set => valeur_TD = value ?? valeur_TD;
  }

  private Label Valeur_TP {
    get => valeur_TP;
    set => valeur_TP = value ?? valeur_TP;
  }

  public DialogueServiceGEII() {
    InitializeComponent();
    ThemeManager.SetTheme(AppTheme.Dark);
    InititialiseFormatEntetes();
    FormatageCouleurLignes(VisualisationDonnées);
    _refreshModuleViewTimer = new DispatcherTimer();
    _refreshModuleViewTimer.Interval = TimeSpan.FromMilliseconds(150);
    _refreshModuleViewTimer.Tick += RefreshModuleViewTimer_Tick;
    _applicationSettings = ApplicationSettingsService.Load();
    mnuAutoOpenLastProfileOnStartup.IsChecked = _applicationSettings.AutoOpenLastProfileOnStartup;
    ApplyProfileToUi();
    ContentRendered += DialogueServiceGEII_Shown;
  }

  private void ApplyUiToProfile() {
    if (_isLoadingProfile)
      return;

    if (_profile == null)
      return;

    _profile.AcademicYear = txtAcademicYear.Text.Trim();

    _profile.SaveState.CurrentSemester = Semestre.Text.Trim();
    _profile.SaveState.CurrentFormation = Formation.Text.Trim();
    _profile.SaveState.CurrentCourse = Cours.Text.Trim();
    _profile.SaveState.CurrentModule =
        texteMatière.Content?.ToString()?.Trim() ?? string.Empty;

    _profile.Window.Left = Left;
    _profile.Window.Top = Top;
    _profile.Window.Width = Width;
    _profile.Window.Height = Height;

    _profile.Window.State = WindowState switch {
      WindowState.Maximized => SaveWindowStateMode.Maximized,
      _ => SaveWindowStateMode.Normal
    };
  }
  private void ApplyProfileToUi() {
    string title = "Gestion Services GEII";

    if (_profile != null) {
      title += " - " + _profile.ProfileName + " - " + _profile.AcademicYear;
      txtAcademicYear.Text = _profile.AcademicYear;
    }

    if (!string.IsNullOrWhiteSpace(_currentProfilePath)) {
      title += " - " + Path.GetFileName(_currentProfilePath);
    }

    if (_selectionWorkbookLoaded && _serviceWorkbookLoaded) {
      title += " [Fichiers chargés]";
    }
    else if (_selectionWorkbookLoaded || _serviceWorkbookLoaded) {
      title += " [Chargement partiel]";
    }
    else {
      title += " [Aucun fichier chargé]";
    }

    Title = title;

    if (!string.IsNullOrWhiteSpace(_profile!.SaveState.CurrentSemester))
      SelectionnerValeurComboBox(Semestre,_profile.SaveState.CurrentSemester);

    if (!string.IsNullOrWhiteSpace(_profile.SaveState.CurrentFormation))
      SelectionnerValeurComboBox(Formation,_profile.SaveState.CurrentFormation);

    if (!string.IsNullOrWhiteSpace(_profile.SaveState.CurrentModule))
      SelectionnerValeurComboBox(Module,_profile.SaveState.CurrentModule);

    if (!string.IsNullOrWhiteSpace(_profile.SaveState.CurrentCourse))
      SelectionnerValeurComboBox(Cours,_profile.SaveState.CurrentCourse);

    Left = _profile.Window.Left;
    Top = _profile.Window.Top;
    Width = _profile.Window.Width;
    Height = _profile.Window.Height;

    WindowState = _profile.Window.State switch {
      SaveWindowStateMode.Maximized => WindowState.Maximized,
      _ => WindowState.Normal
    };
  }

  private void InititialiseFormatEntetes() {
    Style style = new(typeof(DataGridColumnHeader),VisualisationDonnées.ColumnHeaderStyle);

    style.Setters.Add(new Setter(Control.ForegroundProperty,Brushes.LightBlue));
    style.Setters.Add(new Setter(Control.FontFamilyProperty,new FontFamily("Verdana")));
    style.Setters.Add(new Setter(Control.FontSizeProperty,12.0));
    style.Setters.Add(new Setter(Control.FontWeightProperty,FontWeights.Bold));
    style.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty,HorizontalAlignment.Center));
    style.Setters.Add(new Setter(Control.VerticalContentAlignmentProperty,VerticalAlignment.Center));

    VisualisationDonnées.ColumnHeaderStyle = style;
  }

  private void InititialiseFormatTexteParDefaut() {
    VisualisationDonnées.FontFamily = new FontFamily("Verdana");
    VisualisationDonnées.FontSize = 12;
  }

  internal void FormatageCouleurLignes(DataGrid table) {
    InititialiseFormatTexteParDefaut();

    Style style = table.CellStyle == null
        ? new Style(typeof(DataGridCell))
        : new Style(typeof(DataGridCell),table.CellStyle);

    DataTrigger cm = new() { Binding = new Binding(ExcelSchemaNames.Columns.Cours),Value = "CM" };
    cm.Setters.Add(new Setter(Control.BackgroundProperty,Brushes.LightBlue));
    cm.Setters.Add(new Setter(Control.ForegroundProperty,Brushes.Black));

    DataTrigger td = new() { Binding = new Binding(ExcelSchemaNames.Columns.Cours),Value = "TD" };
    td.Setters.Add(new Setter(Control.BackgroundProperty,Brushes.LightGreen));
    td.Setters.Add(new Setter(Control.ForegroundProperty,Brushes.Black));

    DataTrigger tp = new() { Binding = new Binding(ExcelSchemaNames.Columns.Cours),Value = "TP" };
    tp.Setters.Add(new Setter(Control.BackgroundProperty,Brushes.LightGray));
    tp.Setters.Add(new Setter(Control.ForegroundProperty,Brushes.Black));

    DataTrigger attente = new() { Binding = new Binding(ExcelSchemaNames.Columns.Noms),Value = "En attente" };
    attente.Setters.Add(new Setter(Control.BackgroundProperty,Brushes.Red));
    attente.Setters.Add(new Setter(Control.ForegroundProperty,Brushes.Black));
    attente.Setters.Add(new Setter(Control.FontWeightProperty,FontWeights.Bold));

    style.Triggers.Add(cm);
    style.Triggers.Add(td);
    style.Triggers.Add(tp);
    style.Triggers.Add(attente);

    table.CellStyle = style;
  }

  private void AfficherVueGestionServices() {
    pnlGestionServices.Visibility = Visibility.Visible;
    pnlFichesService.Visibility = Visibility.Collapsed;

    mnuVueGestionServices.IsChecked = true;
    mnuVueFichesService.IsChecked = false;
  }

  private void AfficherVueFichesService() {
    pnlGestionServices.Visibility = Visibility.Collapsed;
    pnlFichesService.Visibility = Visibility.Visible;

    mnuVueGestionServices.IsChecked = false;
    mnuVueFichesService.IsChecked = true;

    ChargerIntervenantsFichesService();
  }

  private void ChargerIntervenantsFichesService() {
    if (!TryGetNomTableGlobal(out DataTable nomTableGlobal))
      return;

    List<string> intervenants = FicheServiceHelper.ExtraireIntervenantsDepuisTable(nomTableGlobal);

    cmbIntervenantFicheService.ItemsSource = intervenants;

    if (cmbIntervenantFicheService.Items.Count > 0)
      cmbIntervenantFicheService.SelectedIndex = 0;
  }

  private void RestaurerEffectifsGeii1_DepuisProfil() {
    if (_profile == null || _profile.Groups == null)
      return;

    if (!_profile.Groups.Geii1_IsConfigured)
      return;

    FormationGEII.Geii_1_FI.NbEtudiants = _profile.Groups.Geii1_StudentCount;
    FormationGEII.Geii_1_FI.NbGroupeTD = _profile.Groups.Geii1_TdGroupCount;
    FormationGEII.Geii_1_FI.NbGroupeTP = _profile.Groups.Geii1_TpGroupCount;
    NomGroupeTD_Geii1 = _profile.Groups.Geii1_TdGroupNames ?? [];
    NomGroupeTP_Geii1 = _profile.Groups.Geii1_TpGroupNames ?? [];
    étatEffectif_Geii1.Text = FormationGEII.Geii_1_FI.NbEtudiants.ToString() + " étudiants : ";
    étatNbGroupeTD_Geii1.Text = FormationGEII.Geii_1_FI.NbGroupeTD + " groupes de TD, ";
    étatNbGroupeTP_Geii1.Text = FormationGEII.Geii_1_FI.NbGroupeTP + " groupes de TP\t";

    étatTitreEffectif_Geii1.Visibility
      = étatEffectif_Geii1.Visibility
      = étatNbGroupeTD_Geii1.Visibility
      = étatNbGroupeTP_Geii1.Visibility
      = Visibility.Visible;
  }

  private void RestaurerEffectifsGeii2_FI_DepuisProfil() {
    if (_profile == null || _profile.Groups == null)
      return;

    if (!_profile.Groups.Geii1_IsConfigured)
      return;

    FormationGEII.Geii_2_FI.NbEtudiants = _profile.Groups.Geii2_FI_StudentCount;
    FormationGEII.Geii_2_FI.NbGroupeTD_officiel = _profile.Groups.Geii2_FI_Td_officiel_GroupCount;
    FormationGEII.Geii_2_FI.NbGroupeTP = _profile.Groups.Geii2_FI_TpGroupCount;
    FormationGEII.Geii_2_FI.NbGroupeTPSp = _profile.Groups.Geii2_FI_TpSpGroupCount;
    NomGroupeTD_Geii2_FI = _profile.Groups.Geii2_FI_TdGroupNames ?? [];
    NomGroupeTP_Geii2_FI = _profile.Groups.Geii2_FI_TpGroupNames ?? [];
    NomGroupeTPSp_Geii2_FI = _profile.Groups.Geii2_FI_TpSpGroupNames ?? [];
    étatEffectif_Geii2_FI.Text = FormationGEII.Geii_2_FI.NbEtudiants.ToString() + " étudiants : ";
    étatNbGroupeTD_Geii2_FI.Text = FormationGEII.Geii_2_FI.NbGroupeTD_officiel + " groupes de TD, ";
    étatNbGroupeTP_Geii2_FI.Text = FormationGEII.Geii_2_FI.NbGroupeTP + " groupes de TP, ";
    étatNbGroupeTPSp_Geii2_FI.Text = FormationGEII.Geii_2_FI.NbGroupeTPSp + " groupes de TP Sp";

    étatTitreEffectif_Geii2_FI.Visibility
      = étatEffectif_Geii2_FI.Visibility
      = étatNbGroupeTD_Geii2_FI.Visibility
      = étatNbGroupeTP_Geii2_FI.Visibility
      = Visibility.Visible;
  }

  private void RestaurerEffectifsGeii2_FA_DepuisProfil() {
    if (_profile == null || _profile.Groups == null)
      return;

    if (!_profile.Groups.Geii1_IsConfigured)
      return;

    FormationGEII.Geii_2_FA.NbEtudiants = _profile.Groups.Geii2_FA_StudentCount;
    FormationGEII.Geii_2_FA.NbGroupeTD_officiel = _profile.Groups.Geii2_FA_Td_officiel_GroupCount;
    FormationGEII.Geii_2_FA.NbGroupeTP = _profile.Groups.Geii2_FA_TpGroupCount;
    FormationGEII.Geii_2_FA.NbGroupeTPSp = _profile.Groups.Geii2_FA_TpSpGroupCount;
    NomGroupeTD_Geii2_FA = _profile.Groups.Geii2_FA_TdGroupNames ?? [];
    NomGroupeTP_Geii2_FA = _profile.Groups.Geii2_FA_TpGroupNames ?? [];
    NomGroupeTPSp_Geii2_FA = _profile.Groups.Geii2_FA_TpSpGroupNames ?? [];
    étatEffectif_Geii2_FA.Text = FormationGEII.Geii_2_FA.NbEtudiants.ToString() + " étudiants : ";
    étatNbGroupeTD_Geii2_FA.Text = FormationGEII.Geii_2_FA.NbGroupeTD_officiel + " groupes de TD, ";
    étatNbGroupeTP_Geii2_FA.Text = FormationGEII.Geii_2_FA.NbGroupeTP + " groupes de TP, ";
    étatNbGroupeTPSp_Geii2_FA.Text = FormationGEII.Geii_2_FA.NbGroupeTPSp + " groupes de TP Sp";

    étatTitreEffectif_Geii2_FA.Visibility
      = étatEffectif_Geii2_FA.Visibility
      = étatNbGroupeTD_Geii2_FA.Visibility
      = étatNbGroupeTP_Geii2_FA.Visibility
      = Visibility.Visible;
  }

  private void OpenProfileAndLoadView(string profilePath) {
    _profile = ServiceProfileSerializer.Load(profilePath);
    _currentProfilePath = profilePath;

    RestaurerEffectifsGeii1_DepuisProfil();
    RestaurerEffectifsGeii2_FI_DepuisProfil();
    RestaurerEffectifsGeii2_FA_DepuisProfil();

    _applicationSettings.LastProfilePath = profilePath;
    ApplicationSettingsService.Save(_applicationSettings);

    ApplyProfileToUi();
    LoadFilesFromProfile(showSuccessMessage: false);
  }

  private void ShowProfileStatus() {
    string message = ProfileFileValidator.BuildStatusText(_profile,_currentProfilePath);
  }

  private bool ProfileFilesAreAvailable() {
    return ProfileFileValidator.AreProfileFilesAvailable(_profile);
  }

  public void AfficherListesDonnées(ComboBox liste,ComboBox critèreDeSelection = null!) {
    switch (liste.Name) {
      case "Semestre":
        Semestre = ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name);
        break;
      case "Parcours":
        Parcours = ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name,Semestre,Formation);
        break;
      case "Formation":
        Formation = ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name);
        break;
      case "Noms":
        Noms = ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name);
        break;
      case "Intervenants":
        Intervenants = ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name);
        break;
      case "Titulaires":
        Titulaires = ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name);
        break;
      case "Vacataires":
        Vacataires = ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name);
        break;
      case "Cours":
        Cours = ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name);
        break;
      case "OSE":
        break;
      case "INFOS":
        break;
      case "Module":
        Module = ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name,Semestre,Formation,Parcours);
        break;
      default:
        MessageBox.Show("la variable " + liste.Name + " est introuvable");
        break;
    }
  }

  private static void SelectionnerValeurComboBox(ComboBox liste,string valeur) {
    if (liste == null || string.IsNullOrWhiteSpace(valeur) || liste.Items.Count == 0)
      return;

    string valeurRecherchée = valeur.Trim();

    foreach (object item in liste.Items) {
      string texte = item switch {
        DataRowView rowView when !string.IsNullOrWhiteSpace(liste.DisplayMemberPath) &&
                                 rowView.Row.Table.Columns.Contains(liste.DisplayMemberPath)
          => rowView[liste.DisplayMemberPath]?.ToString()?.Trim() ?? string.Empty,
        _ => item?.ToString()?.Trim() ?? string.Empty
      };

      if (!string.Equals(texte,valeurRecherchée,StringComparison.OrdinalIgnoreCase))
        continue;

      liste.SelectedItem = item;
      return;
    }
  }

  private void ActualiserComboOseDepuisModule(string moduleFiltre) {
    if (_isRefreshingOse)
      return;

    _isRefreshingOse = true;

    try {
      OSE.ItemsSource = null;
      OSE.Items.Clear();

      string moduleCourant = moduleFiltre?.Trim() ?? string.Empty;

      if (string.IsNullOrWhiteSpace(moduleCourant) ||
          DataSetExcelSélection == null ||
          !DataSetExcelSélection.Tables.Contains(ExcelSchemaNames.Tables.Module))
        return;

      DataTable tableModule = DataSetExcelSélection.Tables[ExcelSchemaNames.Tables.Module]!;

      string colonneModule = tableModule.Columns.Contains(ExcelSchemaNames.Columns.Module)
          ? ExcelSchemaNames.Columns.Module
          : "INTITULE";

      string colonneOse = tableModule.Columns.Contains(ExcelSchemaNames.Columns.OSE)
          ? ExcelSchemaNames.Columns.OSE
          : "OSE";

      if (!tableModule.Columns.Contains(colonneModule) || !tableModule.Columns.Contains(colonneOse))
        return;

      SortedSet<string> valeurs = new(StringComparer.OrdinalIgnoreCase);

      foreach (DataRow row in tableModule.Rows) {
        string module = row[colonneModule]?.ToString()?.Trim() ?? string.Empty;

        if (!string.Equals(module,moduleCourant,StringComparison.OrdinalIgnoreCase))
          continue;

        string codeOse = row[colonneOse]?.ToString()?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(codeOse) ||
            codeOse == "0" ||
            codeOse.Equals("#N/A",StringComparison.OrdinalIgnoreCase) ||
            codeOse.Contains("#REF!",StringComparison.OrdinalIgnoreCase))
          continue;

        valeurs.Add(codeOse);
      }

      foreach (string valeur in valeurs)
        OSE.Items.Add(valeur);

      OSE.IsEnabled = OSE.Items.Count > 0;

      if (OSE.Items.Count > 0)
        OSE.SelectedIndex = 0;
    }
    finally {
      _isRefreshingOse = false;
    }
  }

  private void ActualiserComboInfosDepuisGlobal(string moduleFiltre) {
    if (_isRefreshingInfos)
      return;

    _isRefreshingInfos = true;

    try {
      INFOS.ItemsSource = null;
      INFOS.Items.Clear();
      INFOS.Items.Add(string.Empty);

      string moduleCourant = moduleFiltre?.Trim() ?? string.Empty;

      if (DataSetExcelService != null &&
          DataSetExcelService.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal)) {

        DataTable tableGlobal = DataSetExcelService.Tables[ExcelSchemaNames.Tables.NomTableGlobal]!;

        if (tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Module) &&
            tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Infos)) {

          SortedSet<string> valeurs = new(StringComparer.OrdinalIgnoreCase);

          foreach (DataRow row in tableGlobal.Rows) {
            string module = row[ExcelSchemaNames.Columns.Module]?.ToString()?.Trim() ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(moduleCourant) &&
                !string.Equals(module,moduleCourant,StringComparison.OrdinalIgnoreCase))
              continue;

            string valeur = row[ExcelSchemaNames.Columns.Infos]?.ToString()?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(valeur) ||
                valeur == "0" ||
                valeur.Contains("#REF!",StringComparison.OrdinalIgnoreCase) ||
                valeur.Equals("#N/A",StringComparison.OrdinalIgnoreCase))
              continue;

            valeurs.Add(valeur);
          }

          foreach (string valeur in valeurs)
            INFOS.Items.Add(valeur);
        }
      }

      INFOS.Visibility = Visibility.Visible;
      INFOS.IsEnabled = INFOS.Items.Count > 1;

      if (INFOS.Items.Count > 0)
        INFOS.SelectedIndex = 0;
    }
    finally {
      _isRefreshingInfos = false;
    }
  }

  private IEnumerable<ComboBox> GetSelectionComboBoxes() {
    yield return Semestre;
    yield return Formation;
    yield return Parcours;
    yield return Cours;
    yield return Intervenants;
    yield return Titulaires;
    yield return Vacataires;
  }

  private void BindSelectionComboBoxes() {
    foreach (ComboBox comboBox in GetSelectionComboBoxes())
      AfficherListesDonnées(comboBox);

    AfficherListesDonnées(Module,Semestre);

    ActualiserComboInfosDepuisGlobal(Module.Text.Trim());
  }

  private void ActualiserPanneauCours() {
    string formation = Formation.Text?.Trim() ?? string.Empty;
    string semestre = Semestre.Text?.Trim() ?? string.Empty;

    splCoursGeii1.Visibility = Visibility.Collapsed;
    splCoursGeii2_FI.Visibility = Visibility.Collapsed;
    splCoursGeii2_FA.Visibility = Visibility.Collapsed;

    if (formation == "FI" && (semestre == "S1" || semestre == "S2"))
      splCoursGeii1.Visibility = Visibility.Visible;
    else if (formation == "FI" && (semestre == "S3" || semestre == "S4"))
      splCoursGeii2_FI.Visibility = Visibility.Visible;
    else if (formation == "FA" && (semestre == "S3" || semestre == "S4"))
      splCoursGeii2_FA.Visibility = Visibility.Visible;
  }

  private void Formation_SelectionChanged(object sender,SelectionChangedEventArgs e) {
    if (_suspendViewRefresh || _isInitializingView)
      return;
    _suspendViewRefresh = true;

    try {
      AfficherListesDonnées(Parcours);
      AfficherListesDonnées(Module);
    }
    finally {
      _suspendViewRefresh = false;
    }

    RequestRefreshModuleView();

    Dispatcher.BeginInvoke(() => {
      ActualiserPanneauCours();
    },DispatcherPriority.ContextIdle);
  }

  private void Parcours_SelectionChanged(object sender,SelectionChangedEventArgs e) {
    if (_suspendViewRefresh || _isInitializingView)
      return;

    _suspendViewRefresh = true;

    try {
      AfficherListesDonnées(Module);
    }
    finally {
      _suspendViewRefresh = false;
    }

    ActualiserPanneauCours();
    RequestRefreshModuleView();
  }

  private void Cours_SelectionChanged(object sender,SelectionChangedEventArgs e) {
    if (_suspendViewRefresh || _isInitializingView)
      return;

    RequestRefreshModuleView();
  }

  private void INFOS_SelectionChanged(object sender,SelectionChangedEventArgs e) {
    if (_suspendViewRefresh || _isInitializingView || _isRefreshingInfos)
      return;

    if (string.IsNullOrWhiteSpace(INFOS.Text)) {
      RequestRefreshModuleView();
      return;
    }

    AfficherVueAvecFiltreInfos();
  }

  private void VisualisationDonnées_SelectionChanged(object sender,SelectionChangedEventArgs e) {
    if (_isInitializingView || _suspendViewRefresh || VisualisationDonnées.SelectedItem == null)
      return;

    AfficheFicheModule();
  }

  private void AfficherVueAvecFiltreInfos() {
    if (!FichierDeService.Service.FichierOuvert)
      return;

    LireDonnées(
        VisualisationDonnées,
        Semestre,
        Module,
        Formation,
        Cours,
        INFOS,
        null!,
        ExcelSchemaNames.Tables.NomTableGlobal);

    TotalType(VisualisationDonnées);
    FormatageCouleurLignes(VisualisationDonnées);
    AfficheFicheModule();
  }

  private void SaveCurrentProfileIfPossible() {
    if (string.IsNullOrWhiteSpace(_currentProfilePath))
      return;

    ServiceProfileSerializer.Save(_currentProfilePath,_profile);
  }

  private void RedemarrerApplication() {
    string? chemin = Environment.ProcessPath;

    if (string.IsNullOrWhiteSpace(chemin))
      return;

    Process.Start(new ProcessStartInfo {
      FileName = chemin,
      UseShellExecute = true
    });

    Application.Current.Shutdown();
  }

  private void NettoyageDuFormulaire() {
    try {
      //FeuilleClasseur!.DataSource = null;
      //Semestre!.DataSource = null;
      //UE!.DataSource = null;
      //Module.DataSource = null;
      //Formation.DataSource = null;
      //Cours.DataSource = null;
      //Noms.DataSource = null;
      //Intervenants.DataSource = null;
      //Titulaires.DataSource = null;
      //Vacataires.DataSource = null;
      //Commentaires.DataSource = null;
      //boutonsDeModification.Visible = false;

      //DataSetExcelSélection = null!;
      //DataSetExcelService = null!;
      //DataTableExcel = null!;
      //VisualisationDonnées.DataSource = null;

      //Selection = null!;
      FichierDeService.Service = null!;
      ClasseFichier.SuprimeToutesLesTachesExcel();
    }
    catch (Exception ex) {
      MessageBox.Show(ex.Message);
    }
  }


  private void QuitterApplicationProprement() {
    if (_exitRequestedFromMenu)
      return;
    _exitRequestedFromMenu = true;
    ArrêtProcessus();
  }

  internal void ArrêtProcessus() {
    try {
      ApplyUiToProfile();
      SaveCurrentProfileIfPossible();

      NettoyageDuFormulaire();

      FichierDeService.Service?.FermerClasseurExcel();

      FichierDeSelection.Selection?.FermerClasseurExcel();
    }
    finally {
      Application.Current.Shutdown();
    }
  }

  private void ReinitialiserApplicationDepuisProfil() {
    if (_profile == null)
      return;

    ApplyUiToProfile();
    SaveCurrentProfileIfPossible();

    _applicationSettings.LastProfilePath = _currentProfilePath;
    _applicationSettings.AutoOpenLastProfileOnStartup = true;
    ApplicationSettingsService.Save(_applicationSettings);

    _exitRequestedFromMenu = true;

    RedemarrerApplication();
  }

  private DataSet ReadWorkbookTables(WorkbookRole role) {
    switch (role) {
      case WorkbookRole.Selection:
        FichierDeSelection.Selection ??= new ClasseExcel();

        bool selectionPathOk =
          !string.IsNullOrWhiteSpace(FichierDeSelection.Selection.CheminFichier) &&
          File.Exists(FichierDeSelection.Selection.CheminFichier);

        // Si l'objet Selection a perdu son chemin,
        // on tente d'abord de le restaurer depuis le profil JSON.
        if (!selectionPathOk) {
          string selectionPathFromProfile = _profile != null && _profile.Files != null
                  ? _profile.Files.ListsWorkbookPath
                  : string.Empty;

          if (!string.IsNullOrWhiteSpace(selectionPathFromProfile) && File.Exists(selectionPathFromProfile)) {
            FichierDeSelection.Selection.ChargerDepuisChemin(
                selectionPathFromProfile
            );

            selectionPathOk =
                !string.IsNullOrWhiteSpace(FichierDeSelection.Selection.CheminFichier) &&
                File.Exists(FichierDeSelection.Selection.CheminFichier);
          }
        }

        // On n'ouvre l'explorateur que si le chemin est toujours invalide.
        if (!selectionPathOk) {
          if (!FichierDeSelection.Selection.OuvrirFichier(WorkbookRole.Selection))
            return DataSetExcelSélection;
        }

        if (
            string.IsNullOrWhiteSpace(FichierDeSelection.Selection.CheminFichier) ||
            !File.Exists(FichierDeSelection.Selection.CheminFichier)
        )
          return DataSetExcelSélection;

        DataSet dataSetSelection =
            LectureFichierDeSélection(
                FichierDeSelection.Selection
            );

        if (
            dataSetSelection == null ||
            !dataSetSelection.Tables.Contains(ExcelSchemaNames.Tables.Formation)
        ) {
          ReinitialiserApplicationDepuisProfil();

          return DataSetExcelSélection;
        }

        return dataSetSelection;

      case WorkbookRole.Service:
        FichierDeService.Service ??= new ClasseExcel();

        bool servicePathOk =
            !string.IsNullOrWhiteSpace(FichierDeService.Service.CheminFichier) &&
            File.Exists(FichierDeService.Service.CheminFichier);

        // Si l'objet Service a perdu son chemin,
        // on tente d'abord de le restaurer depuis le profil JSON.
        if (!servicePathOk) {
          string servicePathFromProfile =
              _profile != null && _profile.Files != null
                  ? _profile.Files.ServiceWorkbookPath
                  : string.Empty;

          if (
              !string.IsNullOrWhiteSpace(servicePathFromProfile) &&
              File.Exists(servicePathFromProfile)
          ) {
            FichierDeService.Service.ChargerDepuisChemin(
                servicePathFromProfile
            );

            servicePathOk =
                !string.IsNullOrWhiteSpace(FichierDeService.Service.CheminFichier) &&
                File.Exists(FichierDeService.Service.CheminFichier);
          }
        }

        // On n'ouvre l'explorateur que si le chemin est toujours invalide.
        if (!servicePathOk) {
          if (!FichierDeService.Service.OuvrirFichier(WorkbookRole.Service))
            return DataSetExcelService;
        }

        if (
            string.IsNullOrWhiteSpace(FichierDeService.Service.CheminFichier) ||
            !File.Exists(FichierDeService.Service.CheminFichier)
        )
          return DataSetExcelService;

        DataSet dataSetService =
            LectureFichierDeService(
                FichierDeService.Service
            );

        if (
             dataSetService == null ||
             !dataSetService.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal)
         ) {
          ReinitialiserApplicationDepuisProfil();

          return DataSetExcelService;
        }

        return dataSetService;

      default:
        throw new InvalidOperationException(
            "Rôle de classeur inconnu : " + role
        );
    }
  }

  private void LoadWorkbookDataSets() {
    if (DataSetExcelSélection.Tables.Contains(ExcelSchemaNames.Tables.Formation) == false) {
      DataSetExcelSélection = ReadWorkbookTables(WorkbookRole.Selection);
    }

    _selectionWorkbookLoaded =
        DataSetExcelSélection != null &&
        DataSetExcelSélection.Tables.Contains(ExcelSchemaNames.Tables.Formation);

    if (DataSetExcelService.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal) == false) {
      DataSetExcelService = ReadWorkbookTables(WorkbookRole.Service);
      InitialiserReferencesSources(DataSetExcelService.Tables[ExcelSchemaNames.Tables.NomTableGlobal]!);

      DataTable global = DataSetExcelService.Tables[ExcelSchemaNames.Tables.NomTableGlobal]!;

      int nbReferences = global.AsEnumerable()
          .Count(r => r[ExcelSchemaNames.Columns.SourceReferences] != DBNull.Value);
    }

    _serviceWorkbookLoaded =
        DataSetExcelService != null &&
        DataSetExcelService.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal);
  }

  private static void TotalType(DataGrid table) {
    foreach (DataRowView ligne in table.Items.OfType<DataRowView>()) {
      int groupes = int.Parse(ligne[ExcelSchemaNames.Columns.Nombre]?.ToString() ?? string.Empty);
      int durée = int.Parse(ligne[ExcelSchemaNames.Columns.Duree]?.ToString() ?? string.Empty);

      int totalType = groupes * durée;

      ligne[ExcelSchemaNames.Columns.TotalType] = totalType.ToString();
    }
  }

  internal void AfFicheMatière(ComboBox CB1,ComboBox CB2,ComboBox CB3,string tableSelection) {
    if (!_serviceWorkbookLoaded)
      return;

    ClasseBaseDeDonnées.DataSetExcel = DataSetExcelService;

    LireDonnées(VisualisationDonnées,CB1,CB2,CB3,tableSelection);
    TotalType(VisualisationDonnées);
    FormatageCouleurLignes(VisualisationDonnées);
  }

  internal void AfFicheMatière(ComboBox CB1,ComboBox CB2,ComboBox CB3,ComboBox CB4,string tableSelection) {
    if (!_serviceWorkbookLoaded)
      return;

    ClasseBaseDeDonnées.DataSetExcel = DataSetExcelService;

    LireDonnées(VisualisationDonnées,CB1,CB2,CB3,CB4,tableSelection);
    TotalType(VisualisationDonnées);
    FormatageCouleurLignes(VisualisationDonnées);
  }

  internal void AfFicheMatière(ComboBox CB1,ComboBox CB2,ComboBox CB3,ComboBox CB4,ComboBox CB5,string tableSelection) {
    if (!_serviceWorkbookLoaded)
      return;

    ClasseBaseDeDonnées.DataSetExcel = DataSetExcelService;

    LireDonnées(VisualisationDonnées,CB1,CB2,CB3,CB4,CB5,tableSelection);
    TotalType(VisualisationDonnées);
    FormatageCouleurLignes(VisualisationDonnées);
  }

  internal void AfFicheMatière(ComboBox CB1,string tableSelection) {
    if (!_serviceWorkbookLoaded)
      return;

    ClasseBaseDeDonnées.DataSetExcel = DataSetExcelService;

    LireDonnées(VisualisationDonnées,tableSelection);
    TotalType(VisualisationDonnées);
    FormatageCouleurLignes(VisualisationDonnées);
  }

  internal void ControleCouleurBoutonsTDetTP() {
    ushort td = ushort.Parse(nbGroupeTd_Geii1.Content?.ToString() ?? "0");
    ushort tdTotal = ushort.Parse(nbGroupeTdTotal_Geii1.Content?.ToString() ?? "0");
    ushort tp = ushort.Parse(nbGroupeTp_Geii1.Content?.ToString() ?? "0");
    ushort tpTotal = ushort.Parse(nbGroupeTpTotal_Geii1.Content?.ToString() ?? "0");

    if (td != tdTotal) {
      if (td < tdTotal) {
        boutonAjouterTD.Background = Brushes.Red;
        boutonSupprimerTD.Background = Brushes.DarkSeaGreen;
      }
      else {
        boutonAjouterTD.Background = Brushes.DarkSeaGreen;
        boutonSupprimerTD.Background = Brushes.Red;
      }
    }
    else {
      boutonAjouterTD.Background = Brushes.DarkSeaGreen;
      boutonSupprimerTD.Background = Brushes.DarkSeaGreen;
    }

    if (tp != tpTotal) {
      if (tp < tpTotal) {
        boutonAjouterTP.Background = Brushes.Red;
        boutonSupprimerTP.Background = Brushes.DarkSeaGreen;
      }
      else {
        boutonAjouterTP.Background = Brushes.DarkSeaGreen;
        boutonSupprimerTP.Background = Brushes.Red;
      }
    }
    else {
      boutonAjouterTP.Background = Brushes.DarkSeaGreen;
      boutonSupprimerTP.Background = Brushes.DarkSeaGreen;
    }

    boutonAjouterCM.Background = Brushes.DarkSeaGreen;
    boutonSupprimerCM.Background = Brushes.DarkSeaGreen;
  }

  public String RechercheNomModule() {
    return Module.Text;
  }

  public void AfficheFicheModule() {
    DataRowView? ligneCourante = VisualisationDonnées.SelectedItem as DataRowView;

    if (ligneCourante == null)
      return;

    string module = ligneCourante[ExcelSchemaNames.Columns.Module]?.ToString()?.Trim() ?? string.Empty;
    string codeOse = ligneCourante[ExcelSchemaNames.Columns.OSE]?.ToString()?.Trim() ?? string.Empty;
    string libelleCourt = ligneCourante[ExcelSchemaNames.Columns.LibelleCourt]?.ToString()?.Trim() ?? string.Empty;

    string codeCourtModule = ExtraireCodeCourtModuleDepuisLibelleCourt(libelleCourt);

    if (string.IsNullOrWhiteSpace(codeOse))
      codeOse = ConvertirCodeCourtEnCodeOse(codeCourtModule);

    TexteMatière.Content = module;
    TexteCodeOse.Content = codeOse;
    TexteLibelléModule.Text = codeCourtModule;

    CompterLesTypesDeCours(VisualisationDonnées);

    nbGroupeTd_Geii1.Content = CompteurTD.ToString();
    nbGroupeTp_Geii1.Content = CompteurTP.ToString();
    nbGroupeTdTotal_Geii1.Content = FormationGEII.Geii_1_FI.NbGroupeTD.ToString();
    nbGroupeTpTotal_Geii1.Content = FormationGEII.Geii_1_FI.NbGroupeTP.ToString();
    nbEtudiants_Geii1.Content = FormationGEII.Geii_1_FI.NbEtudiants.ToString();

    int nbTdModuleFI = CompterGroupesDistinctsModule("FI","TD");
    int nbTpModuleFI = CompterGroupesDistinctsModule("FI","TP");
    int nbTdModuleFA = CompterGroupesDistinctsModule("FA","TD");
    int nbTpModuleFA = CompterGroupesDistinctsModule("FA","TP");

    nbGroupeTd_Geii2_FI.Content = CompteurTD.ToString();
    nbGroupeTp_Geii2_FI.Content = CompteurTP.ToString();
    nbGroupeTdTotal_Geii2_FI.Content = nbTdModuleFI.ToString();
    nbGroupeTpTotal_Geii2_FI.Content = nbTpModuleFI.ToString();
    nbEtudiants_Geii2_FI.Content = FormationGEII.Geii_2_FI.NbEtudiants.ToString();

    nbGroupeTd_Geii2_FA.Content = CompteurTD.ToString();
    nbGroupeTp_Geii2_FA.Content = CompteurTP.ToString();
    nbGroupeTdTotal_Geii2_FA.Content = nbTdModuleFA.ToString();
    nbGroupeTpTotal_Geii2_FA.Content = nbTpModuleFA.ToString();
    nbEtudiants_Geii2_FA.Content = FormationGEII.Geii_2_FA.NbEtudiants.ToString();

    ControleCouleurBoutonsTDetTP();

    BoutonsDeModification.Visibility = Visibility.Visible;
  }

  private int CompterGroupesDistinctsModule(string formation,string cours) {
    return VisualisationDonnées.Items
        .OfType<DataRowView>()
        .Where(row =>
            string.Equals(GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Formation),formation,StringComparison.OrdinalIgnoreCase) &&
            string.Equals(GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Cours),cours,StringComparison.OrdinalIgnoreCase))
        .Select(row => GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Groupe))
        .Where(groupe => !string.IsNullOrWhiteSpace(groupe) && groupe != "0")
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .Count();
  }

  private void HorairesModule() {
    string nomModule = RechercheNomModule();

    Valeur_CM.Content = RechercheValeur(DataSetExcelSélection,nomModule,"CM","Module") + " h";
    Valeur_TD.Content = RechercheValeur(DataSetExcelSélection,nomModule,"TD","Module") + " h";
    Valeur_TP.Content = RechercheValeur(DataSetExcelSélection,nomModule,"TP","Module") + " h";
  }

  private static string HorairesModule(string NomModule,string NomCours) {
    String ValeurHoraireCours = RechercheValeur(NomModule,NomCours,"Module") + " h";
    return ValeurHoraireCours;
  }

  private void RefreshModuleView() {
    if (!_serviceWorkbookLoaded)
      return;

    if (!DataSetExcelService.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal))
      return;

    if (string.IsNullOrWhiteSpace(Module.Text))
      return;

    Debug.WriteLine("===== FILTRE VUE MODULE =====");
    Debug.WriteLine("Semestre = [" + Semestre.Text + "]");
    Debug.WriteLine("Module   = [" + Module.Text + "]");
    Debug.WriteLine("Formation= [" + Formation.Text + "]");
    Debug.WriteLine("OSE      = [" + OSE.Text + "]");

    ClasseBaseDeDonnées.DataSetExcel = DataSetExcelService;

    AfFicheMatière(Semestre,Module,Formation,"");

    FormatageCouleurLignes(VisualisationDonnées);
    HorairesModule();
    AfficheFicheModule();
  }

  //private void RestaurerGroupesDepuisProfil() {
  //  if (_profile == null || _profile.Groups == null)
  //    return;

  //  if (!_profile.Groups.Geii1_IsConfigured)
  //    return;

  //  FormationGEII.Geii_1_FI.NbEtudiants = _profile.Groups.Geii1_StudentCount;
  //  FormationGEII.Geii_1_FI.NbGroupeTD = _profile.Groups.Geii1_TdGroupCount;
  //  FormationGEII.Geii_1_FI.NbGroupeTP = _profile.Groups.Geii1_TpGroupCount;
  //  NomGroupeTD_Geii1 = _profile.Groups.Geii1_TdGroupNames ?? [];
  //  NomGroupeTP_Geii1 = _profile.Groups.Geii1_TpGroupNames ?? [];
  //  étatEffectif_Geii1.Text = FormationGEII.Geii_1_FI.NbEtudiants.ToString() + " étudiants : ";
  //  étatNbGroupeTD_Geii1.Text = FormationGEII.Geii_1_FI.NbGroupeTD + " groupes de TD, ";
  //  étatNbGroupeTP_Geii1.Text = FormationGEII.Geii_1_FI.NbGroupeTP + " groupes de TP\t";
  //  étatTitreEffectif_Geii1.Visibility
  //    = étatEffectif_Geii1.Visibility
  //    = étatNbGroupeTD_Geii1.Visibility
  //    = étatNbGroupeTP_Geii1.Visibility
  //    = Visibility.Collapsed;
  //}

  private void RequestRefreshModuleView() {
    if (_refreshModuleViewTimer == null)
      return;

    _refreshModuleViewTimer.Stop();
    _refreshModuleViewTimer.Start();
  }

  internal void InitialiserVueDepuisClasseurs() {
    if (!_isInitializingView)
      _isInitializingView = true;

    try {
      LoadWorkbookDataSets();

      ClasseBaseDeDonnées.DataSetExcel = DataSetExcelService;

      BindSelectionComboBoxes();
      ActualiserComboInfosDepuisGlobal(Module.Text.Trim());
    }
    finally {
      _isInitializingView = false;
    }

    Debug.WriteLine(">>> APPEL DIRECT RefreshModuleView");
    RefreshModuleView();
  }

  private void LoadFilesFromProfile(bool showSuccessMessage) {
    if (!ProfileFilesAreAvailable()) {
      ShowProfileStatus();
      return;
    }

    try {
      _isLoadingProfile = true;
      _selectionWorkbookLoaded = false;
      _serviceWorkbookLoaded = false;
      ProfileWorkbookLoadResult loadResult = ProfileWorkbookLoader.Load(_profile);
      if (!loadResult.Success) {
        MessageBox.Show(
            this,
            loadResult.ErrorMessage,
            "Profil",
            MessageBoxButton.OK,
            MessageBoxImage.Warning
        );

        return;
      }

      FichierDeSelection.Selection = loadResult.SelectionWorkbook!;
      FichierDeService.Service = loadResult.ServiceWorkbook!;
      InitialiserVueDepuisClasseurs();
      ApplyProfileToUi();
    }
    catch (Exception ex) {
      MessageBox.Show(
          this,
          "Erreur pendant le chargement des fichiers du profil." +
          Environment.NewLine +
          Environment.NewLine +
          ex.Message,
          "Profil",
          MessageBoxButton.OK,
          MessageBoxImage.Error
      );
    }
    finally {
      _isLoadingProfile =
          false;
    }
  }

  private void AfficheModuleComplet_Click(object sender,RoutedEventArgs e) {

  }

  internal static DataGrid RechercheDoublon(DataGrid table) {
    string GetRowValue(DataRowView row,string columnName) {
      if (row == null || row.Row == null)
        return string.Empty;

      if (!row.Row.Table.Columns.Contains(columnName))
        return string.Empty;

      object value = row.Row[columnName];

      if (value == null || value == DBNull.Value)
        return string.Empty;

      return value.ToString()?.Trim() ?? string.Empty;
    }

    void SetRowValue(DataRowView row,string columnName,object value) {
      if (row == null || row.Row == null)
        return;

      if (!row.Row.Table.Columns.Contains(columnName))
        return;

      row.Row[columnName] = value;
    }

    string FusionnerGroupes(string groupeA,string groupeB) {
      return GroupAssignmentHelper.MergeGroups(
          groupeA,
          groupeB
      );
    }

    HashSet<DataRow> lignesASupprimer = [];

    try {
      for (int ligne = 0;ligne < table.Items.Count;ligne++) {
        if (table.Items[ligne] is not DataRowView row)
          continue;

        if (lignesASupprimer.Contains(row.Row))
          continue;

        for (int ligneEnTraitement = ligne + 1;
             ligneEnTraitement < table.Items.Count;
             ligneEnTraitement++) {

          if (table.Items[ligneEnTraitement] is not DataRowView rowEnTraitement)
            continue;
          if (lignesASupprimer.Contains(rowEnTraitement.Row))
            continue;

          string moduleLigne = GetRowValue(row,ExcelSchemaNames.Columns.Module);
          string moduleLigneEnTraitement = GetRowValue(rowEnTraitement,ExcelSchemaNames.Columns.Module);
          string coursLigne = GetRowValue(row,ExcelSchemaNames.Columns.Cours);
          string coursLigneEnTraitement = GetRowValue(rowEnTraitement,ExcelSchemaNames.Columns.Cours);
          string nomsLigne = GetRowValue(row,ExcelSchemaNames.Columns.Noms);
          string nomsLigneEnTraitement = GetRowValue(rowEnTraitement,ExcelSchemaNames.Columns.Noms);

          bool isMemeModule = moduleLigne == moduleLigneEnTraitement;
          bool isMemeCours = coursLigne == coursLigneEnTraitement;
          bool isMemeIntervenant = nomsLigne == nomsLigneEnTraitement;

          if (!isMemeModule || !isMemeCours || !isMemeIntervenant)
            continue;

          if (!int.TryParse(GetRowValue(row,ExcelSchemaNames.Columns.Nombre),out int nbGroupe))
            continue;

          if (!int.TryParse(GetRowValue(rowEnTraitement,ExcelSchemaNames.Columns.Nombre),out int nbGroupeSupplementaire))
            continue;

          string groupeLigne = GetRowValue(row,ExcelSchemaNames.Columns.Groupe);
          string groupeLigneEnTraitement = GetRowValue(rowEnTraitement,ExcelSchemaNames.Columns.Groupe);
          string groupesFusionnes = FusionnerGroupes(groupeLigne,groupeLigneEnTraitement);
          nbGroupe += nbGroupeSupplementaire;

          if (row[ExcelSchemaNames.Columns.SourceReferences] is Dictionary<string,ExcelSourceReference> references &&
              rowEnTraitement[ExcelSchemaNames.Columns.SourceReferences] is Dictionary<string,ExcelSourceReference> referencesSupplementaires) {
            foreach (KeyValuePair<string,ExcelSourceReference> reference in referencesSupplementaires)
              references[reference.Key] = reference.Value;
          }

          SetRowValue(row,ExcelSchemaNames.Columns.Nombre,nbGroupe);
          SetRowValue(row,ExcelSchemaNames.Columns.Groupe,groupesFusionnes);

          lignesASupprimer.Add(rowEnTraitement.Row);
        }
      }
      TotalType(table);

      foreach (DataRow ligneASupprimer in lignesASupprimer)
        ligneASupprimer.Delete();

      DataSetExcelService.Tables[ExcelSchemaNames.Tables.NomTableGlobal]!.AcceptChanges();
    }
    catch (Exception ex) {
      MessageBox.Show(ex.Message);
    }
    return table;
  }

  private void MiseAJourTableau_Click(object sender,RoutedEventArgs e) {
    if (!VerifierGroupesRenseignesAvantMiseAJour(VisualisationDonnées))
      return;
    RechercheDoublon(VisualisationDonnées);
    FormatageCouleurLignes(VisualisationDonnées);
  }

  public (string Cours,string Groupe) AfficheChoixTypeCours(DataGrid table,string groupeCourant = "") {
    string nouveauTypeCours = string.Empty;
    string nouveauGroupe = string.Empty;

    Hide();

    ChoixTypeDeCours choix = new();
    choix.InitialiserContexteGroupes(table,groupeCourant);

    if (choix.ShowDialog() == true) {
      nouveauTypeCours = choix.TypeDeCoursSelectionne;
      nouveauGroupe = choix.GroupeSelectionne;
    }

    Show();

    return (nouveauTypeCours,nouveauGroupe);
  }

  private string AfficheChoixDuGroupe(DataGrid table,string groupeCourant,string[] groupeTD,string[] groupeTP) {
    GestionDesGroupes gestionGroupes = new();

    gestionGroupes.InitialiserContexteGroupes(table,groupeCourant,groupeTD,groupeTP,false);

    if (gestionGroupes.ShowDialog() == true)
      return gestionGroupes.CbChoixGroupe.Text;

    return groupeCourant;
  }

  private static string GetCellValue(DataGrid table,Point cellule) {
    int colonne = (int)cellule.X;
    int ligne = (int)cellule.Y;

    if (ligne < 0 || ligne >= table.Items.Count)
      return string.Empty;

    if (colonne < 0 || colonne >= table.Columns.Count)
      return string.Empty;

    if (table.Items[ligne] is not DataRowView row)
      return string.Empty;

    if (table.Columns[colonne] is not DataGridTextColumn colonneTexte)
      return string.Empty;

    if (colonneTexte.Binding is not Binding binding)
      return string.Empty;

    string nomColonne = binding.Path.Path;

    if (!row.DataView.Table!.Columns.Contains(nomColonne))
      return string.Empty;

    return row[nomColonne]?.ToString() ?? string.Empty;
  }

  private static void SetCurrentCellValue(DataGrid table,object? value,Point celluleCible) {
    int colonne = (int)celluleCible.X;
    int ligne = (int)celluleCible.Y;

    if (ligne < 0 || ligne >= table.Items.Count)
      return;

    if (colonne < 0 || colonne >= table.Columns.Count)
      return;

    if (table.Items[ligne] is not DataRowView row)
      return;

    DataGridColumn colonneGrid = table.Columns[colonne];

    if (colonneGrid is not DataGridTextColumn colonneTexte)
      return;

    if (colonneTexte.Binding is not Binding binding)
      return;

    string nomColonne = binding.Path.Path;

    if (!row.DataView.Table!.Columns.Contains(nomColonne))
      return;

    row[nomColonne] = value ?? DBNull.Value;
  }

  private void ModifierModule_vue(DataGrid table,Point Cellule) {
    Point CelluleCible = Cellule;
    string groupeCourant = string.Empty;
    DataTable TableExcel = DataSetExcelService.Tables[ExcelSchemaNames.Tables.NomTableGlobal]!;
    do {
      AncienneDonnée = GetCellValue(table,CelluleCible);
      Choix = table.Columns[(int)CelluleCible.X].Header.ToString(); // cellule.X indique la colonne cible
      TableExcel.BeginInit();
      switch (Intervenants.Text) {
        case "Titulaires":
          Noms = Titulaires;
          break;
        case "Vacataires":
          Noms = Vacataires;
          break;
        case "En attente":
          Noms.Text = "En attente";
          break;
      }
      switch (Choix) {
        case "Noms":
          SetCurrentCellValue(table,Noms.Text,CelluleCible);
          break;

        case "Cours":
          DataRowView ligneCourante = (DataRowView)table.Items[(int)CelluleCible.Y];
          _ancienGroupeAvantModification = GetBoundRowValue_vue(ligneCourante,ExcelSchemaNames.Columns.Groupe);
          (string Cours,string Groupe) = AfficheChoixTypeCours(table,_ancienGroupeAvantModification);
          string duree = RechercheValeur(DataSetExcelSélection,Module.Text,Cours,"Module");
          if (string.IsNullOrWhiteSpace(Cours))
            break;
          SetCurrentCellValue(table,Cours,CelluleCible);
          ligneCourante[ExcelSchemaNames.Columns.Groupe] = Groupe;
          ligneCourante[ExcelSchemaNames.Columns.Duree] = duree;
          int.TryParse(duree,out int dureeCours);
          string nombre = GetBoundRowValue_vue(ligneCourante,ExcelSchemaNames.Columns.Nombre);
          int.TryParse(nombre,out int nbGroupes);
          ligneCourante[ExcelSchemaNames.Columns.TotalType] = (dureeCours * nbGroupes).ToString();
          break;
        case "Groupe":
          string[] groupesTD = [];
          string[] groupesTP = [];

          if (Semestre.Text == "S1" || Semestre.Text == "S2") {
            if (NomGroupeTD_Geii1 == null || NomGroupeTD_Geii1.Length == 0 || NomGroupeTD_Geii1[0] == null) {
              RestaurerEffectifsGeii1_DepuisProfil();
              if (NomGroupeTD_Geii1 == null || NomGroupeTD_Geii1.Length == 0 || NomGroupeTD_Geii1[0] == null)
                MenuEffectifGeii_1();
            }

            groupesTD = NomGroupeTD_Geii1!;
            groupesTP = NomGroupeTP_Geii1!;
          }
          else if ((Semestre.Text == "S3" || Semestre.Text == "S4") && Formation.Text == "FI") {
            if (NomGroupeTD_Geii2_officiel_FI == null || NomGroupeTD_Geii2_officiel_FI.Length == 0 || NomGroupeTD_Geii2_officiel_FI[0] == null) {
              RestaurerEffectifsGeii2_FI_DepuisProfil();
              if (NomGroupeTD_Geii2_officiel_FI == null || NomGroupeTD_Geii2_officiel_FI.Length == 0 || NomGroupeTD_Geii2_officiel_FI[0] == null)
                MenuEffectifGeii_2(_profile);
            }

            groupesTD = NomGroupeTD_Geii2_officiel_FI!;
            groupesTP = NomGroupeTP_Geii2_FI!;
          }
          else if ((Semestre.Text == "S3" || Semestre.Text == "S4") && Formation.Text == "FA") {
            if (NomGroupeTD_Geii2_officiel_FA == null || NomGroupeTD_Geii2_officiel_FA.Length == 0 || NomGroupeTD_Geii2_officiel_FA[0] == null) {
              RestaurerEffectifsGeii2_FA_DepuisProfil();
              if (NomGroupeTD_Geii2_officiel_FA == null || NomGroupeTD_Geii2_officiel_FA.Length == 0 || NomGroupeTD_Geii2_officiel_FA[0] == null)
                MenuEffectifGeii_2(_profile);
            }

            groupesTD = NomGroupeTD_Geii2_officiel_FA!;
            groupesTP = NomGroupeTP_Geii2_FA!;
          }

          groupeCourant = GetCellValue(table,CelluleCible) ?? string.Empty;

          Hide();
          SetCurrentCellValue(table,AfficheChoixDuGroupe(table,groupeCourant,groupesTD,groupesTP) ?? string.Empty,CelluleCible);
          Show();
          break;
        default:
          MessageBox.Show("Veuillez sélectionner la bonne CelluleCours à modifier");
          break;
      }
      TableExcel.AcceptChanges();
      TableExcel.EndInit();
    } while (string.IsNullOrEmpty(Noms.Text));
  }

  private bool AjouterModificationSourceEnAttenteDepuisVue(string colonne,Point celluleCible) {
    int ligne = (int)celluleCible.Y;
    object? item = VisualisationDonnées.Items[ligne];
    if (VisualisationDonnées.Items[ligne] is not DataRowView row)
      return false;
    string sourceSheet = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.SourceSheet).Trim();
    string sourceRowText = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.SourceRow).Trim();
    string sourceColumnText = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.SourceColumn).Trim();
    string nouvelleValeur = GetBoundRowValue_vue(row,colonne).Trim();
    if (string.IsNullOrWhiteSpace(sourceSheet)) {
      MessageBox.Show(
          this,
          "Impossible de mettre la modification en attente : SourceSheet est vide.",
          "Modification source",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
      return false;
    }

    if (!int.TryParse(sourceRowText,out int sourceRow) || sourceRow <= 0) {
      MessageBox.Show(
          this,
          "Impossible de mettre la modification en attente : SourceRow est invalide.",
          "Modification source",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);

      return false;
    }

    if (!int.TryParse(sourceColumnText,out int sourceColumn) || sourceColumn <= 0) {
      MessageBox.Show(
          this,
          "Impossible de mettre la modification en attente : SourceColumn est invalide.",
          "Modification source",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);

      return false;
    }

    ClasseBaseDeDonnées.ExcelSourceCellChange modification =
        new() {
          SourceSheet = sourceSheet,
          SourceRow = sourceRow,
          SourceColumn = sourceColumn,
          ColumnName = colonne,
          NewValue = nouvelleValeur
        };

    int indexExistant =
        _pendingSourceChanges.FindIndex(
            change =>
                string.Equals(change.SourceSheet,modification.SourceSheet,StringComparison.OrdinalIgnoreCase) &&
                change.SourceRow == modification.SourceRow &&
                change.SourceColumn == modification.SourceColumn);

    if (indexExistant >= 0) {
      _pendingSourceChanges[indexExistant] =
          modification;
    }
    else {
      _pendingSourceChanges.Add(
          modification);
    }

    Debug.WriteLine(
        "Modification source en attente : " +
        modification.SourceSheet +
        "!" +
        modification.SourceRow +
        "," +
        modification.SourceColumn +
        " = " +
        modification.NewValue);

    return true;
  }

  private bool AjouterReecritureLigneSourceEnAttenteDepuisVue(DataRowView row) {
    string sourceSheet = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.SourceSheet).Trim();
    string sourceRowText = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.SourceRow).Trim();
    string cours = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Cours).Trim();
    string sourceColumnText = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.SourceColumn).Trim();

    if (!int.TryParse(sourceColumnText,out int sourceColumn) || sourceColumn <= 0)
      return false;

    if (string.IsNullOrWhiteSpace(sourceSheet)) {
      MessageBox.Show(
          this,
          "Impossible de préparer le déplacement : SourceSheet est vide.",
          "Modification groupe",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);

      return false;
    }

    if (!int.TryParse(sourceRowText,out int sourceRow) || sourceRow <= 0) {
      MessageBox.Show(
          this,
          "Impossible de préparer le déplacement : SourceRow est invalide.",
          "Modification groupe",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);

      return false;
    }

    if (
        !string.Equals(cours,"TD",StringComparison.OrdinalIgnoreCase) &&
        !string.Equals(cours,"TP",StringComparison.OrdinalIgnoreCase)
    ) {
      MessageBox.Show(
          this,
          "Le déplacement de groupe est prévu uniquement pour TD et TP.",
          "Modification groupe",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);

      return false;
    }

    string cle =
        sourceSheet.Trim() +
        "!" +
        sourceRow +
        "!" +
        _ancienCoursAvantModification.ToUpperInvariant();

    string nouveauCours = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Cours).Trim();
    string nouveauGroupe = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Groupe).Trim();
    string nomIntervenant = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Noms).Trim();

    _pendingSourceRowRewrites[cle] =
        new ClasseBaseDeDonnées.ExcelSourceRowRewrite {
          SourceSheet = sourceSheet.Trim(),

          AncienneSourceRow = sourceRow,
          AncienneSourceColumn = sourceColumn,

          AncienCours = _ancienCoursAvantModification,
          AncienGroupe = _ancienGroupeAvantModification,

          NouveauCours = nouveauCours,
          NouveauGroupe = nouveauGroupe,

          NomIntervenant = nomIntervenant
        };

    return true;
  }

  private bool AjouterDeplacementCoursEnAttenteDepuisVue(Point celluleCible) {
    int ligne = (int)celluleCible.Y;

    if (ligne < 0 || ligne >= VisualisationDonnées.Items.Count)
      return false;

    if (VisualisationDonnées.Items[ligne] is not DataRowView row)
      return false;

    string sourceSheet = GetBoundRowValue_vue(
        row,ExcelSchemaNames.Columns.SourceSheet).Trim();

    string sourceRowText = GetBoundRowValue_vue(
        row,ExcelSchemaNames.Columns.SourceRow).Trim();

    string nouveauCours = GetBoundRowValue_vue(
        row,ExcelSchemaNames.Columns.Cours).Trim();

    string nouveauGroupe = GetBoundRowValue_vue(
        row,ExcelSchemaNames.Columns.Groupe).Trim();

    string nomIntervenant = GetBoundRowValue_vue(
        row,ExcelSchemaNames.Columns.Noms).Trim();

    if (!int.TryParse(sourceRowText,out int ancienneSourceRow))
      return false;

    string cle =
        sourceSheet + "!" +
        ancienneSourceRow + "!" +
        _ancienCoursAvantModification + "!" +
        _ancienGroupeAvantModification;

    _pendingSourceRowRewrites[cle] =
        new ClasseBaseDeDonnées.ExcelSourceRowRewrite {
          SourceSheet = sourceSheet,

          AncienneSourceRow = ancienneSourceRow,
          AncienCours = _ancienCoursAvantModification,
          AncienGroupe = _ancienGroupeAvantModification,

          NouvelleSourceRow = 0, // volontairement inconnue pour l'instant
          NouveauCours = nouveauCours,
          NouveauGroupe = nouveauGroupe,

          NomIntervenant = nomIntervenant
        };

    MessageBox.Show(
        $"Déplacement mémorisé :" +
        $"\n\nIntervenant : {nomIntervenant}" +
        $"\nSource : {sourceSheet}!{ancienneSourceRow}" +
        $"\nAvant : {_ancienCoursAvantModification} / {_ancienGroupeAvantModification}" +
        $"\nAprès : {nouveauCours} / {nouveauGroupe}");

    return true;
  }

  private void UpdateNameCourseHours(Point cellule) {
    int IndexColonne = (int)cellule.X;

    NomColonne = VisualisationDonnées.Columns[IndexColonne].Header?.ToString()
        ?? string.Empty;
    switch (NomColonne) {

      case "Noms":
        ModifierModule_vue(VisualisationDonnées,cellule);
        bool modificationAjoutée = AjouterModificationSourceEnAttenteDepuisVue(ExcelSchemaNames.Columns.Noms,cellule);
        if (!modificationAjoutée)
          break;

        FormatageCouleurLignes(VisualisationDonnées);
        AfficheFicheModule();
        break;

      case "Groupe":
        if (VisualisationDonnées.Items[(int)cellule.Y] is not DataRowView rowGroupe)
          break;

        _ancienCoursAvantModification = GetBoundRowValue_vue(rowGroupe,ExcelSchemaNames.Columns.Cours);
        _ancienGroupeAvantModification = GetBoundRowValue_vue(rowGroupe,ExcelSchemaNames.Columns.Groupe);

        ModifierModule_vue(VisualisationDonnées,cellule);

        if (!AjouterReecritureLigneSourceEnAttenteDepuisVue(rowGroupe))
          break;

        FormatageCouleurLignes(VisualisationDonnées);
        AfficheFicheModule();
        break;

      case "Cours":

        if (VisualisationDonnées.Items[(int)cellule.Y] is not DataRowView row)
          break;
        ModifierModule_vue(VisualisationDonnées,cellule);
        if (!AjouterReecritureLigneSourceEnAttenteDepuisVue(row))
          break;
        break;

      case "Durée":
        MessageBox.Show(
            this,
            "La modification des durées sera traitée dans une étape suivante."
            + Environment.NewLine + Environment.NewLine
            + "Pour l'instant, seules les colonnes Noms et Groupe sont modifiées directement dans les feuilles S1 / S2.",
            "Modification non encore activée",
            MessageBoxButton.OK,
            MessageBoxImage.Information);

        break;

      default:
        MessageBox.Show(
            this,
            "Cet élément n'est pas modifiable",
            "Modification",
            MessageBoxButton.OK,
            MessageBoxImage.Information);

        break;
    }
  }

  private Point _celluleSelectionnee = new Point(-1,-1);

  private void VisualisationDonnées_MouseLeftButtonUp(object sender,MouseButtonEventArgs e) {
    DependencyObject? element = e.OriginalSource as DependencyObject;
    DataGridCell? cellule = FindParent<DataGridCell>(element);
    if (cellule == null)
      return;
    int colonne = VisualisationDonnées.Columns.IndexOf(cellule.Column);
    if (cellule.DataContext == null)
      return;
    int ligne = VisualisationDonnées.Items.IndexOf(cellule.DataContext);
    _celluleSelectionnee = new Point(colonne,ligne);
    Console.WriteLine($"Cellule sélectionnée : X={colonne}, Y={ligne}");
  }

  private static T? FindParent<T>(DependencyObject? element)
    where T : DependencyObject {
    while (element != null) {
      if (element is T result)
        return result;

      element = VisualTreeHelper.GetParent(element);
    }

    return null;
  }

  private async void BoutonModifier_Click(object sender,RoutedEventArgs e) {
    bool flowControl = await ModifierCellule();

    if (!flowControl)
      return;
  }

  private async Task<bool> ModifierCellule() {
    Point cellule = _celluleSelectionnee;

    if (cellule.X < 0 || cellule.Y < 0) {
      MessageBox.Show(this,"Veuillez sélectionner une celluleCible.");
      return false;
    }

    UpdateNameCourseHours(cellule);
    return await MettreAJourSource();
  }

  private void BoutonAnnuler_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuFichierRéinitialiserEnvironnement_Click(object sender,RoutedEventArgs e) {
    ReinitialiserApplicationDepuisProfil();
  }

  private void AjouterLigne(DataGrid table,string coursAttendu) {
    int ligne = (int)_celluleSelectionnee.Y;

    if (ligne < 0 || ligne >= table.Items.Count ||
        table.Items[ligne] is not DataRowView row) {
      MessageBox.Show(
          this,
          "Veuillez sélectionner une cellule dans le tableau.",
          "Ajout ligne",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
      return;
    }

    string groupeCourant = string.Empty;
    string formation = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Formation).Trim();
    string cours = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Cours).Trim();
    string module = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Module).Trim();


    if (!string.Equals(cours,coursAttendu,StringComparison.OrdinalIgnoreCase)) {
      MessageBox.Show(
          this,
          "Veuillez sélectionner une ligne de " + coursAttendu + " dans le tableau.",
          "Ajout ligne",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
      return;
    }

    (cours,groupeCourant) = AfficheChoixTypeCours(table);

    if (string.IsNullOrWhiteSpace(cours))
      return;

    if (!string.Equals(cours,coursAttendu,StringComparison.OrdinalIgnoreCase)) {
      MessageBox.Show(
          this,
          "Le type de cours sélectionné doit être " + coursAttendu + ".",
          "Ajout ligne",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
      return;
    }

    switch (Intervenants.Text) {
      case "Titulaires":
        Noms = Titulaires;
        break;

      case "Vacataires":
        Noms = Vacataires;
        break;

      case "En attente":
        Noms.Text = "En attente";
        break;
    }

    string nomAjoute = Noms.Text.Trim();

    string sourceSheet = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.SourceSheet).Trim();
    string sourceRowTexte = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.SourceRow).Trim();

    if (!int.TryParse(sourceRowTexte,out int sourceRow) || string.IsNullOrWhiteSpace(sourceSheet)) {
      MessageBox.Show(this,"La ligne sélectionnée ne possède pas de référence source valide.","Ajout ligne",MessageBoxButton.OK,MessageBoxImage.Warning);
      return;
    }


    if (string.IsNullOrWhiteSpace(nomAjoute)) {
      MessageBox.Show(
          this,
          "Veuillez sélectionner un intervenant.",
          "Ajout ligneCalculée",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
      return;
    }

    bool ajoutEffectue = ClasseBaseDeDonnées.AjouterIntervenantSourceSemestre_Epplus(
      FichierDeService.Service,
      sourceSheet,
      sourceRow,
      coursAttendu,
      groupeCourant,
      nomAjoute,
      out int nouvelleSourceRow,
      out int nouvelleSourceColumn);

    if (!ajoutEffectue)
      return;

    DataRow nouvelleLigne = row.Row.Table.NewRow();

    foreach (DataColumn colonne in row.Row.Table.Columns)
      nouvelleLigne[colonne.ColumnName] = row.Row[colonne.ColumnName];

    nouvelleLigne[ExcelSchemaNames.Columns.Noms] = nomAjoute;
    nouvelleLigne[ExcelSchemaNames.Columns.Cours] = coursAttendu;
    nouvelleLigne[ExcelSchemaNames.Columns.Groupe] = groupeCourant;
    nouvelleLigne[ExcelSchemaNames.Columns.SourceSheet] = sourceSheet;
    nouvelleLigne[ExcelSchemaNames.Columns.SourceRow] = nouvelleSourceRow.ToString();
    nouvelleLigne[ExcelSchemaNames.Columns.SourceColumn] = nouvelleSourceColumn.ToString();

    string duree = ClasseBaseDeDonnées.RechercheValeur(DataSetExcelSélection,module,coursAttendu,"Module");

    nouvelleLigne[ExcelSchemaNames.Columns.Duree] = duree;
    nouvelleLigne[ExcelSchemaNames.Columns.Nombre] = "1";
    nouvelleLigne[ExcelSchemaNames.Columns.TotalType] = duree;

    row.Row.Table.Rows.Add(nouvelleLigne);
  }

  private async Task AjouterLigne(string type) {
    AjouterLigne(VisualisationDonnées,type);
    await MettreAJourSource();
  }

  private void SuppressionOuMiseAJourVueEtFichierExcel(DataGrid table,string typeCoursAttendu) {
    int ligne = (int)_celluleSelectionnee.Y;

    if (ligne < 0 || ligne >= table.Items.Count ||
        table.Items[ligne] is not DataRowView row) {
      MessageBox.Show(
          this,
          "Veuillez sélectionner une cellule dans le tableau.",
          "Suppression",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
      return;
    }

    if (row[ExcelSchemaNames.Columns.SourceReferences] is not Dictionary<string,ClasseBaseDeDonnées.ExcelSourceReference> references ||
        references.Count == 0) {
      MessageBox.Show(
          this,
          "Aucune référence vers la feuille source n'est disponible pour cette ligne.",
          "Suppression",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
      return;
    }

    string noms = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Noms).Trim();
    string cours = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Cours).Trim();
    string nombreText = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Nombre).Trim();
    string dureeText = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Duree).Trim();
    string groupe = GetBoundRowValue_vue(row,ExcelSchemaNames.Columns.Groupe).Trim();

    if (!int.TryParse(nombreText,out int nbGroupes)) {
      MessageBox.Show(
          this,
          "Le nombre de groupes est invalide : " + nombreText,
          "Suppression",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
      return;
    }

    if (!int.TryParse(dureeText,out int duree)) {
      MessageBox.Show(
          this,
          "La durée est invalide : " + dureeText,
          "Suppression",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
      return;
    }

    if (!string.Equals(cours,typeCoursAttendu,StringComparison.OrdinalIgnoreCase)) {
      MessageBox.Show(
          this,
          "Veuillez sélectionner une ligne de " + typeCoursAttendu + " dans le tableau.",
          "Suppression",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
      return;
    }

    string groupeASupprimer;

    if (references.Count == 1) {
      groupeASupprimer = references.Keys.First();
    }
    else {
      List<string> groupesAffectes = GroupAssignmentHelper.ExtractGroups(groupe);

      ChoixSuppressionGroupe choixSuppression = new();

      choixSuppression.Initialiser(
          noms,
          cours,
          groupesAffectes
      );

      if (choixSuppression.ShowDialog() != true)
        return;

      groupeASupprimer = choixSuppression.GroupeSelectionne;

      if (string.IsNullOrWhiteSpace(groupeASupprimer))
        return;
    }

    if (!references.TryGetValue(
            groupeASupprimer,
            out ClasseBaseDeDonnées.ExcelSourceReference? sourceReference)) {
      MessageBox.Show(
          this,
          "La référence source du groupe " + groupeASupprimer + " est introuvable.",
          "Suppression",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
      return;
    }

    if (!SupprimerAffectationSource(FichierDeService.Service,sourceReference.SourceSheet,sourceReference.SourceRow,sourceReference.SourceColumn))
      return;

    references.Remove(groupeASupprimer);

    if (references.Count == 0) {
      row.Row.Delete();
    }
    else {
      string nouveauGroupe = GroupAssignmentHelper.RemoveGroup(groupe,groupeASupprimer);
      int nouveauNombre = GroupAssignmentHelper.CountGroups(nouveauGroupe);

      row[ExcelSchemaNames.Columns.Groupe] = nouveauGroupe;
      row[ExcelSchemaNames.Columns.Nombre] = nouveauNombre;
      row[ExcelSchemaNames.Columns.TotalType] = nouveauNombre * duree;
    }
    DataSetExcelService.AcceptChanges();
  }

  private async Task SupprimerLigne(string type) {
    SuppressionOuMiseAJourVueEtFichierExcel(VisualisationDonnées,type);
    await MettreAJourSource();
  }

  private async void boutonAjouterCM_Click(object sender,RoutedEventArgs e) {
    await AjouterLigne("CM");
  }

  private async void boutonSupprimerCM_Click(object sender,RoutedEventArgs e) {
    await SupprimerLigne("CM");
  }

  private async void boutonAjouterTD_Click(object sender,RoutedEventArgs e) {
    await AjouterLigne("TD");
  }

  private async void boutonSupprimerTD_Click(object sender,RoutedEventArgs e) {
    await SupprimerLigne("TD");
  }

  private async void boutonAjouterTP_Click(object sender,RoutedEventArgs e) {
    await AjouterLigne("TP");
  }

  private async void boutonSupprimerTP_Click(object sender,RoutedEventArgs e) {
    await SupprimerLigne("TP");
  }

  internal static void MiseAJourModule_fichierDeServiceExcel(ClasseExcel FichierDialogue,DataSet DataSetExcel,string nomBase,string Matière,bool premierResultatSeulement = false) {
    MiseAJourLigneModule_fichierDeServiceExcel(FichierDialogue,DataSetExcel,nomBase,Matière,premierResultatSeulement);
  }

  private void ModifierFicheMatière(ServiceManagerProfile profile) {
    Stopwatch chronoTotal = Stopwatch.StartNew();
    Stopwatch chrono = Stopwatch.StartNew();

    Hide();
    string NomModule = RechercheNomModule();

    Debug.WriteLine($"[CHRONO RETOUR] Préparation : {chrono.ElapsedMilliseconds} ms");

    chrono.Restart();

    DataSetExcelSélection = ReadWorkbookTables(WorkbookRole.Selection);

    Debug.WriteLine($"[CHRONO RETOUR] Lecture Selection AVANT fenêtre : {chrono.ElapsedMilliseconds} ms");

    chrono.Restart();

    ConfigurationDesCours fenetre = new(NomModule,Matière,VisualisationDonnées,profile,DataSetExcelSélection);

    Debug.WriteLine($"[CHRONO RETOUR] ConfigurationDesCours complète : {chrono.ElapsedMilliseconds} ms");

    if (fenetre.FicheModifiee) {
      string valeurCM = fenetre.ValeurCMValidee;
      string valeurTD = fenetre.ValeurTDValidee;
      string valeurTP = fenetre.ValeurTPValidee;

      MiseAJourModuleSelection_DataTable(DataSetExcelSélection,NomModule,valeurCM,valeurTD,valeurTP);

      Valeur_CM.Content = valeurCM + " h";
      Valeur_TD.Content = valeurTD + " h";
      Valeur_TP.Content = valeurTP + " h";

      MiseAJourDureesAttenduesGlobal_DataTable(
          DataSetExcelService,
          NomModule,
          valeurCM,
          valeurTD,
          valeurTP);

      TotalType(VisualisationDonnées);
    }
    chrono.Restart();

    AfficheFicheModule();
    Show();

    Debug.WriteLine($"[CHRONO RETOUR] AfficheFicheModule + Show : {chrono.ElapsedMilliseconds} ms");
    Debug.WriteLine($"[CHRONO RETOUR] TOTAL ModifierFicheMatière : {chronoTotal.ElapsedMilliseconds} ms");
  }

  //private void ModifierFicheMatière(ServiceManagerProfile profile) {
  //  Hide();
  //  string NomModule = RechercheNomModule();
  //  DataSetExcelSélection = ReadWorkbookTables(WorkbookRole.Selection);

  //  ConfigurationDesCours fenetre = new(NomModule,Matière,VisualisationDonnées,profile,DataSetExcelSélection);

  //  if (fenetre.FicheModifiee) {
  //    DataSetExcelSélection = ReadWorkbookTables(WorkbookRole.Selection);

  //    string valeurCM = RechercheValeur(DataSetExcelSélection,NomModule,ExcelSchemaNames.Columns.CM,ExcelSchemaNames.Tables.Module);
  //    string valeurTD = RechercheValeur(DataSetExcelSélection,NomModule,ExcelSchemaNames.Columns.TD,ExcelSchemaNames.Tables.Module);
  //    string valeurTP = RechercheValeur(DataSetExcelSélection,NomModule,ExcelSchemaNames.Columns.TP,ExcelSchemaNames.Tables.Module);

  //    Valeur_CM.Content = valeurCM + " h";
  //    Valeur_TD.Content = valeurTD + " h";
  //    Valeur_TP.Content = valeurTP + " h";

  //    MiseAJourDureesAttenduesGlobal_DataTable(DataSetExcelService,NomModule,valeurCM,valeurTD,valeurTP);
  //    TotalType(VisualisationDonnées);
  //  }

  //  AfficheFicheModule();
  //  Show();
  //}

  private async void ModifierFicheMatière_Click(object sender,RoutedEventArgs e) {
    ModifierFicheMatière(_profile);
    await MettreAJourSource();
  }

  private void MenuOuvrirFichier_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuFichierEnregistrer_Click(object sender,RoutedEventArgs e) {
    FichierDeService.Service.EnregistrerFichier(false);
  }

  private void MenuFichierEnregistrerSous_Click(object sender,RoutedEventArgs e) {
    FichierDeService.Service.EnregistrerFichier(false);
  }

  private void MenuImprimerFicheService_Click(object sender,RoutedEventArgs e) {
    ImprimerFicheService(apercu: false);
  }

  private void MenuApercuFicheService_Click(object sender,RoutedEventArgs e) {
    ImprimerFicheService(apercu: true);
  }

  private void mnuQuitter_Click(object sender,RoutedEventArgs e) {
    QuitterApplicationProprement();
  }

  private void MenuAnnuler_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuRetablir_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuCouper_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuCopier_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuColler_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuSelectionnerTout_Click(object sender,RoutedEventArgs e) {

  }

  private bool TryLoadWorkbooksFromProfileIfAvailable() {
    if (!ProfileFilesAreAvailable())
      return false;

    ProfileWorkbookLoadResult loadResult =
        ProfileWorkbookLoader.Load(_profile);

    if (!loadResult.Success)
      return false;

    FichierDeSelection.Selection =
        loadResult.SelectionWorkbook!;

    FichierDeService.Service =
        loadResult.ServiceWorkbook!;

    return true;
  }


  internal void AfficherLesCritèresSélection() {
    ClasseFichier.SuprimeToutesLesTachesExcel();
    _isLoadingProfile = true;
    try {
      TryLoadWorkbooksFromProfileIfAvailable();
      InitialiserVueDepuisClasseurs();
      ApplyProfileToUi();
    }
    finally {
      _isLoadingProfile = false;
    }
  }

  private void EnregistrerEffectifsGeii1_DansProfil() {
    if (_profile == null)
      return;

    _profile.Groups ??= new ServiceGroupsOptions();
    _profile.Groups.Geii1_IsConfigured = true;
    _profile.Groups.Geii1_FormationName = "GEII 1";
    _profile.Groups.Geii1_StudentCount = FormationGEII.Geii_1_FI.NbEtudiants;
    _profile.Groups.Geii1_TdGroupCount = FormationGEII.Geii_1_FI.NbGroupeTD;
    _profile.Groups.Geii1_TdGroupCount = FormationGEII.Geii_1_FI.NbGroupeTD;
    _profile.Groups.Geii1_TpGroupCount = FormationGEII.Geii_1_FI.NbGroupeTP;
    _profile.Groups.Geii1_TdGroupNames = NomGroupeTD_Geii1;
    _profile.Groups.Geii1_TpGroupNames = NomGroupeTP_Geii1;
    if (!string.IsNullOrWhiteSpace(_currentProfilePath)) {
      ServiceProfileSerializer.Save(
          _currentProfilePath,
          _profile
      );
    }
  }

  private void EnregistrerEffectifsGeii2_FI_DansProfil() {
    if (_profile == null)
      return;

    _profile.Groups ??= new ServiceGroupsOptions();
    _profile.Groups.Geii2_FI_IsConfigured = true;
    _profile.Groups.Geii2_FI_FormationName = "GEII 2 FI";

    _profile.Groups.Geii2_FI_StudentCount = FormationGEII.Geii_2_FI.NbEtudiants;
    _profile.Groups.Geii2_FI_TdGroupCount = FormationGEII.Geii_2_FI.NbGroupeTD;
    _profile.Groups.Geii2_FI_Td_officiel_GroupCount = FormationGEII.Geii_2_FI.NbGroupeTD_officiel;
    _profile.Groups.Geii2_FI_TpGroupCount = FormationGEII.Geii_2_FI.NbGroupeTP;
    _profile.Groups.Geii2_FI_TpSpGroupCount = FormationGEII.Geii_2_FI.NbGroupeTPSp;
    _profile.Groups.Geii2_FI_TdGroupNames = NomGroupeTD_Geii2_FI;
    _profile.Groups.Geii2_FI_Td_officiel_GroupNames = NomGroupeTD_Geii2_officiel_FI;
    _profile.Groups.Geii2_FI_TpGroupNames = NomGroupeTP_Geii2_FI;
    _profile.Groups.Geii2_FI_TpSpGroupNames = NomGroupeTPSp_Geii2_FI;

    if (!string.IsNullOrWhiteSpace(_currentProfilePath))
      ServiceProfileSerializer.Save(_currentProfilePath,_profile);
  }

  private void EnregistrerEffectifsGeii2_FA_DansProfil() {
    if (_profile == null)
      return;

    _profile.Groups ??= new ServiceGroupsOptions();
    _profile.Groups.Geii2_FA_IsConfigured = true;
    _profile.Groups.Geii2_FA_FormationName = "GEII 2 FA";

    _profile.Groups.Geii2_FA_StudentCount = FormationGEII.Geii_2_FA.NbEtudiants;
    _profile.Groups.Geii2_FA_TdGroupCount = FormationGEII.Geii_2_FA.NbGroupeTD;
    _profile.Groups.Geii2_FA_Td_officiel_GroupCount = FormationGEII.Geii_2_FA.NbGroupeTD_officiel;
    _profile.Groups.Geii2_FA_TpGroupCount = FormationGEII.Geii_2_FA.NbGroupeTP;
    _profile.Groups.Geii2_FA_TpSpGroupCount = FormationGEII.Geii_2_FA.NbGroupeTPSp;
    _profile.Groups.Geii2_FA_TdGroupNames = NomGroupeTD_Geii2_FA;
    _profile.Groups.Geii2_FA_Td_officiel_GroupNames = NomGroupeTD_Geii2_officiel_FA;
    _profile.Groups.Geii2_FA_TpGroupNames = NomGroupeTP_Geii2_FA;
    _profile.Groups.Geii2_FA_TpSpGroupNames = NomGroupeTPSp_Geii2_FA;

    if (!string.IsNullOrWhiteSpace(_currentProfilePath))
      ServiceProfileSerializer.Save(_currentProfilePath,_profile);
  }

  private void MenuEffectifGeii_1() {
    string formation = "GEII 1";
    _configurationEffectifs_GEII_1 = new();
    Hide();
    _configurationEffectifs_GEII_1.nomDeLaFormation.Content = $"Répartition prévisionnelle des effectifs en {formation}";
    _configurationEffectifs_GEII_1.ShowDialog();

    NomGroupeTD_Geii1 = _configurationEffectifs_GEII_1.NomGroupeTD;
    NomGroupeTP_Geii1 = _configurationEffectifs_GEII_1.NomGroupeTP;
    EnregistrerEffectifsGeii1_DansProfil();

    if (!string.IsNullOrWhiteSpace(FichierDeService.Service.CheminFichier) && File.Exists(FichierDeService.Service.CheminFichier)) {
      AfficheFicheModule();
    }
    Show();
  }

  private void MenuEffectifGeii_2(ServiceManagerProfile profile) {
    ConfigurationEffectifs_GEII_2 window = new(profile);
    Hide();

    window.ShowDialog();

    nbGroupeTD_officiel_Geii2_FI = window.nbGroupeTD_officiel_FI;
    NomGroupeTD_Geii2_FI = window.NomGroupeTD_FI;
    NomGroupeTD_Geii2_officiel_FI = window.NomGroupeTD_officiel_FI;
    NomGroupeTP_Geii2_FI = window.NomGroupeTP_FI;
    NomGroupeTPSp_Geii2_FI = window.NomGroupeTPSp_FI;

    nbGroupeTD_officiel_Geii2_FA = window.nbGroupeTD_officiel_FA;
    NomGroupeTD_Geii2_FA = window.NomGroupeTD_FA;
    NomGroupeTD_Geii2_officiel_FA = window.NomGroupeTD_officiel_FA;
    NomGroupeTP_Geii2_FA = window.NomGroupeTP_FA;
    NomGroupeTPSp_Geii2_FA = window.NomGroupeTPSp_FA;

    EnregistrerEffectifsGeii2_FI_DansProfil();
    EnregistrerEffectifsGeii2_FA_DansProfil();

    if (!string.IsNullOrWhiteSpace(FichierDeService.Service.CheminFichier) && File.Exists(FichierDeService.Service.CheminFichier)) {
      AfficheFicheModule();
    }

    Show();
  }

  private void AfficherLesCritèresDeSélection_Click(object sender,RoutedEventArgs e) {
    AfficherLesCritèresSélection();
  }

  private void MenuEffectifGeii_1_Click(object sender,RoutedEventArgs e) {
    MenuEffectifGeii_1();
  }

  private void MenuEffectifGeii_2_Click(object sender,RoutedEventArgs e) {
    MenuEffectifGeii_2(_profile);
  }

  private void mnuGestionDesIntervenants_Click(object sender,RoutedEventArgs e) {
    AfficherGestionDesIntervenants();
  }

  private void mnuVueGestionServices_Click(object sender,RoutedEventArgs e) {
    AfficherVueGestionServices();
  }

  private void mnuGenererStatutsIntervenants_Click(object sender,RoutedEventArgs e) {
    GenererJsonStatutsIntervenants();
  }

  private void mnuCleanCache_Click(object sender,RoutedEventArgs e) {
    MessageBoxResult result =
        MessageBox.Show(
            this,
            "Voulez-vous vraiment vider le cache du fichier de service ?" +
            Environment.NewLine +
            Environment.NewLine +
            "Les données actuellement affichées ne seront pas supprimées." +
            Environment.NewLine +
            "Le cache sera reconstruit au prochain chargement du fichier Excel.",
            "Vider le cache service",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

    if (result != MessageBoxResult.Yes)
      return;

    int deletedFiles = ServiceWorkbookDiskCache.InvalidateServiceWorkbookCache();

    MessageBox.Show(
        this,
        deletedFiles + " fichier(s) de cache supprimé(s)." +
        Environment.NewLine +
        Environment.NewLine +
        "Le cache sera reconstruit au prochain chargement du fichier de service.",
        "Cache service",
        MessageBoxButton.OK,
        MessageBoxImage.Information);
  }

  private static int TrouverIndexColonneDataGrid(DataGrid table,string nomColonne) {
    for (int i = 0;i < table.Columns.Count;i++) {
      if (string.Equals(table.Columns[i].Header?.ToString(),nomColonne,StringComparison.OrdinalIgnoreCase))
        return i;
    }

    return -1;
  }

  private bool VerifierGroupesRenseignesAvantMiseAJour(DataGrid table) {
    int indexColonneGroupe = TrouverIndexColonneDataGrid(table,"Groupe");

    if (indexColonneGroupe < 0) {
      MessageBox.Show(
        this,
        "La colonne Groupe est introuvable dans le tableau.",
        "Mise à jour impossible",
        MessageBoxButton.OK,
        MessageBoxImage.Warning);
      return false;
    }

    List<int> lignesIncompletes = [];

    for (int i = 0;i < table.Items.Count;i++) {
      if (table.Items[i] is not DataRowView row)
        continue;

      string cours = row["Cours"]?.ToString()?.Trim() ?? string.Empty;

      if (string.Equals(cours,"CM",StringComparison.OrdinalIgnoreCase))
        continue;

      string groupe = row["Groupe"]?.ToString()?.Trim() ?? string.Empty;

      if (string.IsNullOrWhiteSpace(groupe))
        lignesIncompletes.Add(i + 1);
    }

    if (lignesIncompletes.Count == 0)
      return true;

    MessageBox.Show(this,"La mise à jour du tableau est impossible tant que toutes les lignes ne possèdent pas un groupe." +
      Environment.NewLine + Environment.NewLine +
      $"Nombre de champ(s) Groupe vide(s) : {lignesIncompletes.Count}" +
      Environment.NewLine +
      $"Première ligneCalculée concernée : {lignesIncompletes[0]}",
      "Groupes manquants",
      MessageBoxButton.OK,
      MessageBoxImage.Warning);

    return false;
  }

  private async void mnuMettreAJourSource_Click(object sender,RoutedEventArgs e) {
    bool flowControl = await MettreAJourSource();
    if (!flowControl) {
      return;
    }
  }

  private async Task<bool> MettreAJourSource() {
    if (_pendingSourceChanges.Count == 0 && _pendingSourceRowRewrites.Count == 0)
      return false;

    if (!VerifierGroupesRenseignesAvantMiseAJour(VisualisationDonnées))
      return false;

    int nbCellulesModifiees = AppliquerModificationsCellulesSourcesSemestre_Epplus(FichierDeService.Service,_pendingSourceChanges);

    int nbLignesReecrites = ReecrireLignesSourcesGroupesSemestre_Epplus(
        FichierDeService.Service,
        DataSetExcelService.Tables[ExcelSchemaNames.Tables.NomTableGlobal]!,
        _pendingSourceRowRewrites.Values);

    _pendingSourceChanges.Clear();
    _pendingSourceRowRewrites.Clear();

    MessageBox.Show(
        this,
        nbCellulesModifiees + " cellule(s) source modifiée(s)." +
        Environment.NewLine +
        nbLignesReecrites + " ligneCalculée(s) source réécrite(s).",
        "Mise à jour source",
        MessageBoxButton.OK,
        MessageBoxImage.Information);

    FormatageCouleurLignes(VisualisationDonnées);
    AfficheFicheModule();

    Stopwatch chronoCache = Stopwatch.StartNew();

    await ReconstruireCacheFichierServiceAsync(FichierDeService.Service);

    Debug.WriteLine($"[CHRONO CACHE ASYNC] Reconstruction : {chronoCache.ElapsedMilliseconds} ms");

    return true;
  }

  private void mnuTest_Click(object sender,RoutedEventArgs e) {
    Stopwatch chrono = Stopwatch.StartNew();

    ReconstruireCacheFichierService(FichierDeService.Service);

    Debug.WriteLine($"[TEST CACHE À CHAUD] Reconstruction complète : {chrono.ElapsedMilliseconds} ms");
  }

  private void mnuCheminsAcces_Click(object sender,RoutedEventArgs e) {
    if (_isLoadingProfile)
      return;

    ProfilePathsWindow dialog = new(_profile);

    if (dialog.ShowDialog() != true)
      return;

    ApplyProfileToUi();
    SaveCurrentProfileIfPossible();
    ShowProfileStatus();
  }
  private void mnuVueFichesService_Click(object sender,RoutedEventArgs e) {
    AfficherVueFichesService();
  }

  private void btnVueFichesService_Click(object sender,RoutedEventArgs e) {
    AfficherVueFichesService();
  }

  private void mnuOpenProfile_Click(object sender,RoutedEventArgs e) {
    OpenFileDialog dialog = new();
    dialog.Title = "Open service profile";
    dialog.Filter =
      "Service profile (*.service.json)|*.service.json|" +
      "All files (*.*)|*.*";

    if (dialog.ShowDialog(this) != true)
      return;
  }

  private void mnuSaveProfile_Click(object sender,RoutedEventArgs e) {
    if (string.IsNullOrWhiteSpace(_currentProfilePath)) {
      mnuSaveProfileAs_Click(sender,e);
      return;
    }

    ServiceProfileSerializer.Save(_currentProfilePath,_profile);
  }

  private void SaveProfileAs(string profilePath) {
    if (_profile == null)
      return;

    ApplyUiToProfile();

    ServiceProfileSerializer.Save(profilePath,_profile);

    _currentProfilePath = profilePath;

    _applicationSettings.LastProfilePath = profilePath;
    ApplicationSettingsService.Save(_applicationSettings);

    ShowProfileStatus();
  }

  private void mnuSaveProfileAs_Click(object sender,RoutedEventArgs e) {
    if (_profile == null)
      return;

    SaveFileDialog dialog = new() {
      Title = "Enregistrer le profil sous...",
      Filter = "Profil Gestion Service GEII (*.service.json)|*.service.json|Tous les fichiers (*.*)|*.*",
      FileName = Path.GetFileName(_currentProfilePath)
    };

    if (!string.IsNullOrWhiteSpace(_currentProfilePath))
      dialog.InitialDirectory = Path.GetDirectoryName(_currentProfilePath);

    if (dialog.ShowDialog(this) != true)
      return;

    SaveProfileAs(dialog.FileName);
  }

  private void mnuProfileStatus_Click(object sender,RoutedEventArgs e) {
    ShowProfileStatus();
  }

  private void mnuOpenLastProfile_Click(object sender,RoutedEventArgs e) {
    string lastProfilePath = _applicationSettings.LastProfilePath;

    if (string.IsNullOrWhiteSpace(lastProfilePath)) {
      MessageBox.Show(
          this,
          "Aucun dernier profil n'est enregistré.",
          "Profil",
          MessageBoxButton.OK,
          MessageBoxImage.Information
      );
      return;
    }

    if (!File.Exists(lastProfilePath)) {
      MessageBox.Show(
          this,
          "Le dernier profil enregistré est introuvable :" +
          Environment.NewLine +
          Environment.NewLine +
          lastProfilePath,
          "Profil",
          MessageBoxButton.OK,
          MessageBoxImage.Information
      );
      return;
    }
    OpenProfileAndLoadView(lastProfilePath);
  }

  private void mnuAutoOpenLastProfileOnStartup_Click(object sender,RoutedEventArgs e) {
    _applicationSettings.AutoOpenLastProfileOnStartup = mnuAutoOpenLastProfileOnStartup.IsChecked;

    ApplicationSettingsService.Save(_applicationSettings);
  }

  private void mnuloadProfil_Click(object sender,RoutedEventArgs e) {
    LoadFilesFromProfile(showSuccessMessage: true);
  }

  private void mnuAbout_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuCheckForUpdates_Click(object sender,RoutedEventArgs e) {

  }

  private void btnCreateServiceFiche_Click(object sender,RoutedEventArgs e) {
    if (_isLoadingProfile)
      return;
    GenerateServiceSheets window =
        new(
            DataSetExcelSélection,
            DataSetExcelService,
            SelectionWorkbook,
            ServiceWorkbook,
            _profile
          );
    Hide();
    window.ShowDialog();
    Show();

    ApplyProfileToUi();
    SaveCurrentProfileIfPossible();
  }

  private void btnGestionService_Click(object sender,RoutedEventArgs e) {
    AfficherVueGestionServices();
  }

  private void btnGestionDesIntervenants_Click(object sender,RoutedEventArgs e) {
    AfficherGestionDesIntervenants();
  }

  private void AfficherGestionDesIntervenants() {
    string selectionWorkbookPath = FichierDeSelection.Selection.CheminFichier;
    GestionIntervenants window = new(DataSetExcelSélection,selectionWorkbookPath);
    Hide();
    window.ShowDialog();
    Show();
  }

  private void btnExporterFicheService_Click(object sender,RoutedEventArgs e) {
    ExporterFicheServicePdf();
  }

  private void btnActualiserFicheService_Click(object sender,RoutedEventArgs e) {
    string intervenant = cmbIntervenantFicheService.Text.Trim();
    ActualiserApercuFicheService(intervenant);
  }

  private void ImprimerFicheService(bool apercu) {
    if (string.IsNullOrWhiteSpace(rtbFicheService.Document.Blocks.ToString())) {
      MessageBox.Show(
          this,
          "Aucune fiche de service n'est affichée.",
          "Impression",
          MessageBoxButton.OK,
          MessageBoxImage.Information);

      return;
    }

    string intervenant = cmbIntervenantFicheService.Text.Trim();

    string documentName =
        string.IsNullOrWhiteSpace(intervenant)
            ? "Service prévisionnel"
            : $"Service prévisionnel - {intervenant}";

    RichTextBox richTextBoxImpression = CreerRichTextBoxFicheServicePourImpression();

    RichTextBoxPrintHelper printHelper = new(richTextBoxImpression);

    if (apercu) {
      printHelper.ShowPreviewWithPrinterDialog(this,documentName);
    }
    else {
      printHelper.Print(this,documentName);
    }
  }

  private RichTextBox CreerRichTextBoxFicheServicePourImpression() {
    TextRange source = new(
        rtbFicheService.Document.ContentStart,
        rtbFicheService.Document.ContentEnd);

    FlowDocument document = new() {
      FontFamily = new FontFamily("Consolas"),
      FontSize = 14,
      Foreground = Brushes.Black,
      Background = Brushes.White,
      PagePadding = new Thickness(0)
    };

    Paragraph paragraphe = new(new Run(source.Text)) {
      Margin = new Thickness(0),
      FontFamily = new FontFamily("Consolas"),
      FontSize = 14,
      Foreground = Brushes.Black
    };

    document.Blocks.Add(paragraphe);

    RichTextBox richTextBoxImpression = new() {
      Document = document,
      FontFamily = new FontFamily("Consolas"),
      FontSize = 14,
      Background = Brushes.White,
      Foreground = Brushes.Black,
      BorderThickness = new Thickness(0),
      IsReadOnly = true
    };

    return richTextBoxImpression;
  }

  private void btnImprimerFicheService_Click(object sender,RoutedEventArgs e) {
    ImprimerFicheService(apercu: false);
  }

  private void DialogueServiceGEII_Shown(object sender,EventArgs e) {
    if (!_applicationSettings.AutoOpenLastProfileOnStartup)
      return;

    string lastProfilePath = _applicationSettings.LastProfilePath;

    if (string.IsNullOrWhiteSpace(lastProfilePath))
      return;

    if (!File.Exists(lastProfilePath)) {
      MessageBox.Show(
          this,
          "Le dernier profil enregistré est introuvable :" +
          Environment.NewLine +
          Environment.NewLine +
          lastProfilePath,
          "Profil",
          MessageBoxButton.OK,
          MessageBoxImage.Warning
      );

      return;
    }

    OpenProfileAndLoadView(lastProfilePath);
  }

  private void RefreshModuleViewTimer_Tick(object? sender,EventArgs e) {
    _refreshModuleViewTimer?.Stop();
    RefreshModuleView();
  }

  private void Semestre_SelectionChanged(object sender,SelectionChangedEventArgs e) {
    if (_suspendViewRefresh)
      return;

    _suspendViewRefresh = true;

    try {
      AfficherListesDonnées(Parcours);
      AfficherListesDonnées(Module);
    }
    finally {
      _suspendViewRefresh = false;
    }

    ActualiserPanneauCours();
    RequestRefreshModuleView();
  }

  private void Module_SelectionChanged(object sender,SelectionChangedEventArgs e) {
    if (_suspendViewRefresh)
      return;

    ActualiserComboOseDepuisModule(Module.Text.Trim());
    ActualiserComboInfosDepuisGlobal(Module.Text.Trim());

    if (!_isInitializingView)
      RequestRefreshModuleView();
  }

  private void btnFichierRéinitialiserEnvironnement_Click(object sender,RoutedEventArgs e) {
    ReinitialiserApplicationDepuisProfil();
  }

  private void DialogueServiceGEII_Closing(object? sender,CancelEventArgs e) {
    if (_exitRequestedFromMenu)
      return;

    e.Cancel = true;
    QuitterApplicationProprement();
  }

  private void élémentMenuGEII2_Click(object sender,RoutedEventArgs e) {

  }

  private async void _PreviewMouseLeftButtonDown(
    object sender,
    MouseButtonEventArgs e) {

  }
  private DataGridCell? _celluleActive;
  private void DesactiverCellule() {

    if (_celluleActive == null)
      return;

    _celluleActive.ClearValue(Control.BackgroundProperty);
    _celluleActive = null;
  }

  private void VisualisationDonnées_PreviewMouseLeftButtonDown(object sender,MouseButtonEventArgs e) {
    if (e.OriginalSource is not DependencyObject source) {
      DesactiverCellule();
      return;
    }

    DataGridCell? cellule =
        FindParent<DataGridCell>(source);

    if (cellule == null) {
      DesactiverCellule();
      return;
    }

    if (_celluleActive != null && _celluleActive != cellule)
      _celluleActive.ClearValue(Control.BackgroundProperty);

    _celluleActive = cellule;
    cellule.Background = Brushes.Yellow;
  }

  private async void VisualisationDonnées_MouseDoubleClick(object sender,MouseButtonEventArgs e) {
    if (_celluleActive == null)
      return;
    await ModifierCellule();
    DesactiverCellule();
  }

  internal string GenererJsonStatutsIntervenants() {
    string message;
    if (DataSetExcelService == null) {
      MessageBox.Show(
          this,
          "Aucun fichier de sélection n'est chargé.",
          "Génération des statuts",
          MessageBoxButton.OK,
          MessageBoxImage.Information);
      message = "Echec de la mise à jour du statut des intervenants";
      return message;
    }

    TeacherStatusesFile statusFile =
        FicheServiceStatusHelper.BuildFromSelectionDataSet(
            DataSetExcelSélection,
            FichierDeSelection.Selection.CheminFichier
            );

    string jsonPath = FicheServiceStatusHelper.GetDefaultJsonPath();

    FicheServiceStatusHelper.Save(
        jsonPath,
        statusFile);

    message = $"{statusFile.Teachers.Count} statut(s) intervenant(s) généré(s).{Environment.NewLine}{Environment.NewLine}{jsonPath}";

    MessageBox.Show(
        this,
        message,
        "Génération des statuts",
        MessageBoxButton.OK,
        MessageBoxImage.Information);
    return message;
  }


  private void SelectionnerStatutDepuisJson(string intervenant) {
    string jsonPath = FicheServiceStatusHelper.GetDefaultJsonPath();
    TeacherStatusesFile statusFile = FicheServiceStatusHelper.Load(jsonPath);
    if (!FicheServiceStatusHelper.TryGetStatus(statusFile,intervenant,out string statut)) {
      return;
    }
    int index = cmbStatutFicheService.Items.IndexOf(statut);
    if (index >= 0) {
      cmbStatutFicheService.SelectedIndex = index;
    }
  }

  private static string ConstruireNomFichierPdf(string intervenant) {
    string nom =
        $"Service_previsionnel_{intervenant}_{DateTime.Now:yyyyMMdd}.pdf";

    foreach (char caractere in Path.GetInvalidFileNameChars()) {
      nom =
          nom.Replace(caractere,'_');
    }

    return nom;
  }

  private void ExporterFicheServicePdf() {
    string intervenant = cmbIntervenantFicheService.Text.Trim();
    string statut = cmbStatutFicheService.Text.Trim();
    if (!ValiderSelectionFicheService(intervenant,statut))
      return;
    if (!TryGetNomTableGlobal(out DataTable nomTableGlobal))
      return;
    RegleStatutService regle = FicheServiceHelper.ObtenirRegleStatutService(statut);
    Dictionary<string,LigneServicePrevisionnel> lignesParModule = FicheServiceHelper.ConstruireLignesServicePrevisionnel(nomTableGlobal,intervenant,regle);
    if (lignesParModule.Count == 0) {
      MessageBox.Show(
          this,
          "Aucune ligne de service n'a été trouvée pour cet intervenant.",
          "Export PDF",
          MessageBoxButton.OK,
          MessageBoxImage.Information);
      return;
    }
    string nomFichier = ConstruireNomFichierPdf(intervenant);
    SaveFileDialog dialogue =
        new() {
          Title = "Exporter la fiche de service en PDF",
          Filter = "Fichier PDF (*.pdf)|*.pdf",
          FileName = nomFichier,
          AddExtension = true,
          DefaultExt = "pdf",
          OverwritePrompt = true
        };
    if (dialogue.ShowDialog(this) != true)
      return;
    string intervenantAffiche = GetIntervenantAfficheFicheService(intervenant);
    FicheServicePdfHelper.ExporterPdf(
        dialogue.FileName,
        ConstruireTitreServicePrevisionnel(),
        intervenantAffiche,
        statut,
        lignesParModule.Values,
        regle);
    Process.Start(
        new ProcessStartInfo {
          FileName = dialogue.FileName,
          UseShellExecute = true
        });
  }

  private void cmbIntervenantFicheService_SelectionChanged(object sender,SelectionChangedEventArgs e) {
    if (cmbIntervenantFicheService.SelectedItem is not string intervenant || string.IsNullOrWhiteSpace(intervenant))
      return;

    SelectionnerStatutDepuisJson(intervenant);
    ActualiserApercuFicheService(intervenant);
  }

  private void ActualiserApercuFicheService(string intervenant) {
    string statut = cmbStatutFicheService.Text.Trim();
    string intervenantAffiche = GetIntervenantAfficheFicheService(intervenant);
    if (!ValiderSelectionFicheService(intervenant,statut))
      return;
    if (!TryGetNomTableGlobal(out DataTable NomTableGlobal))
      return;
    RegleStatutService regle = FicheServiceHelper.ObtenirRegleStatutService(statut);
    Dictionary<string,LigneServicePrevisionnel> lignesParModule =
        FicheServiceHelper.ConstruireLignesServicePrevisionnel(
            NomTableGlobal,
            intervenant,
            regle);
    AfficherServicePrevisionnel(
        intervenant,
        statut,
        lignesParModule.Values,
        lignesParModule.Values.Sum(ligne => ligne.Cm),
        lignesParModule.Values.Sum(ligne => ligne.Td),
        lignesParModule.Values.Sum(ligne => ligne.Tp),
        lignesParModule.Values.Sum(ligne => ligne.TotalRetenu),
        regle.ServiceReference);
  }

  private static void AjouterLigneRichTextBox(RichTextBox richTextBox,string texte,Color couleur,FontWeight? fontWeight = null) {
    Run run = new(texte) {
      Foreground = new SolidColorBrush(couleur),
      FontWeight = fontWeight ?? FontWeights.Normal
    };

    Paragraph paragraphe = new(run) {
      Margin = new Thickness(0)
    };

    richTextBox.Document.Blocks.Add(paragraphe);
  }

  private string ConstruireTitreServicePrevisionnel() {
    string academicYear = _profile.AcademicYear.Trim();
    if (string.IsNullOrWhiteSpace(academicYear))
      return "SERVICE PRÉVISIONNEL";
    return $"SERVICE PRÉVISIONNEL {academicYear}";
  }


  private void AfficherServicePrevisionnel(string intervenant,string statut,IEnumerable<LigneServicePrevisionnel> lignes,decimal totalCm,decimal totalTd,decimal totalTp,decimal totalRetenu,decimal serviceReference) {
    string intervenantAffiche = GetIntervenantAfficheFicheService(intervenant);
    List<LigneServicePrevisionnel> lignesTriees =
      [.. lignes
        .OrderBy(ligne => FicheServiceHelper.ObtenirOrdreSemestre(ligne.Semestre))
        .ThenBy(ligne => ligne.Ose)
        .ThenBy(ligne => ligne.Module)];

    rtbFicheService.Document.Blocks.Clear();
    rtbFicheService.Background = new SolidColorBrush(Color.FromRgb(28,28,28));
    rtbFicheService.Foreground = Brushes.Gainsboro;
    rtbFicheService.FontFamily = new FontFamily("Consolas");
    rtbFicheService.FontSize = 12;

    Color couleurTitre = Color.FromRgb(255,180,80);
    Color couleurInfo = Color.FromRgb(140,200,255);
    Color couleurEntete = Color.FromRgb(255,220,120);
    Color couleurTexte = Colors.Gainsboro;
    Color couleurTotal = Color.FromRgb(150,230,170);
    Color couleurErreur = Color.FromRgb(255,130,130);

    AjouterLigneRichTextBox(rtbFicheService,ConstruireTitreServicePrevisionnel(),couleurTitre,FontWeights.Bold);
    AjouterLigneRichTextBox(rtbFicheService,string.Empty,couleurTexte);
    AjouterLigneRichTextBox(rtbFicheService,$"Intervenant : {intervenantAffiche}",couleurInfo,FontWeights.Bold);
    AjouterLigneRichTextBox(rtbFicheService,$"Statut      : {statut}",couleurInfo,FontWeights.Bold);
    AjouterLigneRichTextBox(rtbFicheService,string.Empty,couleurTexte);
    AjouterLigneRichTextBox(rtbFicheService,"OSE".PadRight(12) + "Module".PadRight(28) + "CM".PadLeft(8) + "TD".PadLeft(8) + "TP".PadLeft(8) + "Retenu".PadLeft(12),couleurEntete,FontWeights.Bold);
    AjouterLigneRichTextBox(rtbFicheService,new string('-',64),Colors.DimGray);

    foreach (IGrouping<string,LigneServicePrevisionnel> groupeSemestre in lignesTriees.GroupBy(ligne => ligne.Semestre)) {
      AjouterLigneRichTextBox(rtbFicheService,FicheServiceHelper.ConstruireLibelleSemestre(groupeSemestre.Key),couleurTitre,FontWeights.Bold);

      foreach (LigneServicePrevisionnel ligne in groupeSemestre) {
        AjouterLigneRichTextBox(
            rtbFicheService,
            ligne.Ose.PadRight(12) +
            ligne.Module.PadRight(28) +
            ligne.Cm.ToString("0.##").PadLeft(8) +
            ligne.Td.ToString("0.##").PadLeft(8) +
            ligne.Tp.ToString("0.##").PadLeft(8) +
            ligne.TotalRetenu.ToString("0.##").PadLeft(12),
            couleurTexte);
      }

      AjouterLigneRichTextBox(rtbFicheService,string.Empty,couleurTexte);
    }

    AjouterLigneRichTextBox(rtbFicheService,new string('-',76),Colors.DimGray);
    AjouterLigneRichTextBox(
        rtbFicheService,
        string.Empty.PadRight(12) +
        "TOTAL".PadRight(28) +
        totalCm.ToString("0.##").PadLeft(8) +
        totalTd.ToString("0.##").PadLeft(8) +
        totalTp.ToString("0.##").PadLeft(8) +
        totalRetenu.ToString("0.##").PadLeft(12),
        couleurTotal,
        FontWeights.Bold);

    if (serviceReference > 0m) {
      decimal ecartService = totalRetenu - serviceReference;
      Color couleurEcart;
      string messageHc;

      if (ecartService < 0m) {
        couleurEcart = Color.FromRgb(255,130,130);
        messageHc = $"Sous Service           : {ecartService:0.##} h";
      }
      else if (ecartService == 0m) {
        couleurEcart = Color.FromRgb(150,230,170);
        messageHc = $"Heures complémentaires : {ecartService:0.##} h";
      }
      else {
        couleurEcart = Color.FromRgb(255,180,80);
        messageHc = $"Heures complémentaires : {ecartService:0.##} h";
      }

      AjouterLigneRichTextBox(rtbFicheService,string.Empty,couleurTexte);
      AjouterLigneRichTextBox(rtbFicheService,$"Service de référence   : {serviceReference:0.##} h",couleurInfo);
      AjouterLigneRichTextBox(rtbFicheService,$"Service retenu         : {totalRetenu:0.##} h",couleurTotal,FontWeights.Bold);
      AjouterLigneRichTextBox(rtbFicheService,messageHc,couleurEcart,FontWeights.Bold);
    }

    if (!lignes.Any()) {
      AjouterLigneRichTextBox(rtbFicheService,string.Empty,couleurTexte);
      AjouterLigneRichTextBox(rtbFicheService,"Aucune ligne de service trouvée pour cet intervenant.",couleurErreur,FontWeights.Bold);
    }

    rtbFicheService.CaretPosition = rtbFicheService.Document.ContentStart;
    rtbFicheService.ScrollToHome();
  }
  private static string GetIntervenantAfficheFicheService(string intervenant) {
    if (string.IsNullOrWhiteSpace(intervenant))
      return string.Empty;
    string jsonPath = FicheServiceStatusHelper.GetDefaultJsonPath();
    if (!File.Exists(jsonPath))
      return intervenant;
    TeacherStatusesFile statusFile = FicheServiceStatusHelper.Load(jsonPath);
    return FicheServiceStatusHelper.GetDisplayName(statusFile,intervenant);
  }

  private static void SetRichTextBoxText(RichTextBox richTextBox,string texte) {
    richTextBox.Document.Blocks.Clear();
    richTextBox.Document.Blocks.Add(new Paragraph(new Run(texte)));
  }

  private bool ValiderSelectionFicheService(string intervenant,string statut) {
    if (string.IsNullOrWhiteSpace(intervenant)) {
      SetRichTextBoxText(rtbFicheService,"Veuillez sélectionner un intervenant.");
      return false;
    }

    if (string.IsNullOrWhiteSpace(statut)) {
      SetRichTextBoxText(rtbFicheService,"Veuillez sélectionner un statut.");
      return false;
    }

    return true;
  }

  private bool TryGetNomTableGlobal(out DataTable nomTableGlobal) {
    nomTableGlobal = null!;
    if (DataSetExcelService == null ||
    !DataSetExcelService.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal)) {
      SetRichTextBoxText(rtbFicheService,"La table complète n'est pas chargée.");
      return false;
    }

    DataTable? table =
        DataSetExcelService.Tables[ExcelSchemaNames.Tables.NomTableGlobal];

    if (table == null) {
      SetRichTextBoxText(rtbFicheService,"La table complète est introuvable.");
      return false;
    }

    nomTableGlobal = table;
    return true;
  }

}