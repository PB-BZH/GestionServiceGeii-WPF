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
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Threading;
using GestionServiceGeii.Core.Enums;
using GestionServiceGeii.Core.Models;
using GestionServiceGeii.Core.Profiles;
using GestionServiceGeii.Core.Services;
using GestionServiceGeii.Shared.Database;
using GestionServiceGeii.Shared.Librairie_Fichier;
using GestionServiceGeii.Shared.Librairie_GEII;
using GestionServiceGeii.Shared.Librairie_Générique;
using GestionServiceGeii.UI.Windows;
using PB.BZH.Theme.Theming;
using DataTable = System.Data.DataTable;

namespace GestionServiceGeii.UI;
/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class DialogueServiceGEII: Window {
  private string ancienneDonnée = string.Empty;
  private static FichierDeSelection selection = new();
  private ClasseGénérique générique = new();
  private static readonly FormationGEII formationGEII = new();
  private static ConfigurationEffectifs configurationEffectifs = new();
  private static ConfigurationDesCours configurationDesCours = new();
  private static ChoixTypeDeCours choixTypeDeCours = new();
  private static readonly GestionDesGroupes gestionDesGroupes = new();
  private static DataSet dataSetExcelSélection = new("FICHIER DE SELECTION");
  private static DataSet dataSetExcelService = new("FICHIER DE SERVICE");
  private static DataTable dataTableExcel = new("Table ExcelApp");
  private static ComboBox noms = new();
  private static string? choix;
  private static string? nomColonne;
  private static int indexLigne_vue;
  internal static string[] NomGroupeTD = new string[4];
  internal static string[] NomGroupeTP = new string[8];

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

  private static ConfigurationEffectifs ConfigurationEffectifs {
    get => configurationEffectifs;
    set => configurationEffectifs = value ?? configurationEffectifs;
  }

  private static ConfigurationDesCours ConfigurationDesCours {
    get => configurationDesCours;
    set => configurationDesCours = value ?? configurationDesCours;
  }

  private static ChoixTypeDeCours ChoixTypeDeCours {
    get => choixTypeDeCours;
    set => choixTypeDeCours = value ?? choixTypeDeCours;
  }

  internal static DataSet DataSetExcelSélection {
    get => dataSetExcelSélection;
    set => dataSetExcelSélection = value ?? dataSetExcelService;
  }

  internal static DataSet DataSetExcelService {
    get => dataSetExcelService;
    set => dataSetExcelService = value ?? dataSetExcelService;
  }

  private static DataTable DataTableExcel {
    get => dataTableExcel;
    set => dataTableExcel = value ?? dataTableExcel;
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
    FichierDeSelection.Selection = new ClasseExcel();
    FichierDeService.Service = new ClasseExcel();
    Closed += DialogueServiceGEII_Closed;
  }

  private void DialogueServiceGEII_Closed(object? sender,EventArgs e) {
    Application.Current.Shutdown();
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
    _profile.SaveState.CurrentModule = texteMatière.Content?.ToString()?.Trim() ?? string.Empty;
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
  }

  private void RestaurerEffectifsDepuisProfil() {
    if (_profile == null || _profile.Groups == null)
      return;

    if (!_profile.Groups.IsConfigured)
      return;

    FormationGEII.FI_Geii_1.NbEtudiants =
        _profile.Groups.StudentCount;

    FormationGEII.FI_Geii_1.NbGroupeTD =
        _profile.Groups.TdGroupCount;

    FormationGEII.FI_Geii_1.NbGroupeTP =
        _profile.Groups.TpGroupCount;

    NomGroupeTD =
        _profile.Groups.TdGroupNames ?? [];

    NomGroupeTP =
        _profile.Groups.TpGroupNames ?? [];

    étatEffectifGeii1.Text =
        FormationGEII.FI_Geii_1.NbEtudiants.ToString() + " étudiants : ";

    étatNbGroupeTD.Text =
        FormationGEII.FI_Geii_1.NbGroupeTD + " groupes de TD, ";

    étatNbGroupeTP.Text =
        FormationGEII.FI_Geii_1.NbGroupeTP + " groupes de TP\t";

    étatTitreEffectifGeii1.Visibility
      = étatEffectifGeii1.Visibility
      = étatNbGroupeTD.Visibility
      = étatNbGroupeTP.Visibility
      = Visibility.Collapsed;
  }

  private void OpenProfileAndLoadView(string profilePath) {
    _profile = ServiceProfileSerializer.Load(profilePath);
    RestaurerEffectifsDepuisProfil();
    _currentProfilePath = profilePath;
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
      case "UE":
        UE = critèreDeSelection == null
            ? ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name)
            : ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name,critèreDeSelection);
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
        Module = critèreDeSelection == null
            ? ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name)
            : ClasseGénérique.RempliComboBox(liste,DataSetExcelSélection,liste.Name,critèreDeSelection);
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

  private void Formation_SelectionChanged(object sender,SelectionChangedEventArgs e) {
    if (_suspendViewRefresh || _isInitializingView)
      return;

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

    ClasseBaseDeDonnées.LireDonnées(
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
          string selectionPathFromProfile =
              _profile != null && _profile.Files != null
                  ? _profile.Files.ListsWorkbookPath
                  : string.Empty;

          if (
              !string.IsNullOrWhiteSpace(selectionPathFromProfile) &&
              File.Exists(selectionPathFromProfile)
          ) {
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
            ClasseBaseDeDonnées.LectureFichierDeSélection(
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
            ClasseBaseDeDonnées.LectureFichierDeService(
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
      DataSetExcelSélection =
          ReadWorkbookTables(WorkbookRole.Selection);
    }

    _selectionWorkbookLoaded =
        DataSetExcelSélection != null &&
        DataSetExcelSélection.Tables.Contains(ExcelSchemaNames.Tables.Formation);

    Debug.WriteLine(
    "Avant chargement service : Global présent = " +
    DataSetExcelService.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal));

    if (DataSetExcelService.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal) == false) {
      DataSetExcelService =
          ReadWorkbookTables(WorkbookRole.Service);
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

    ClasseBaseDeDonnées.LireDonnées(VisualisationDonnées,CB1,CB2,CB3,tableSelection);
    TotalType(VisualisationDonnées);
    FormatageCouleurLignes(VisualisationDonnées);
  }

  internal void AfFicheMatière(ComboBox CB1,ComboBox CB2,ComboBox CB3,ComboBox CB4,string tableSelection) {
    if (!_serviceWorkbookLoaded)
      return;

    ClasseBaseDeDonnées.DataSetExcel = DataSetExcelService;

    ClasseBaseDeDonnées.LireDonnées(VisualisationDonnées,CB1,CB2,CB3,CB4,tableSelection);
    TotalType(VisualisationDonnées);
    FormatageCouleurLignes(VisualisationDonnées);
  }

  internal void AfFicheMatière(ComboBox CB1,ComboBox CB2,ComboBox CB3,ComboBox CB4,ComboBox CB5,string tableSelection) {
    if (!_serviceWorkbookLoaded)
      return;

    ClasseBaseDeDonnées.DataSetExcel = DataSetExcelService;

    ClasseBaseDeDonnées.LireDonnées(VisualisationDonnées,CB1,CB2,CB3,CB4,CB5,tableSelection);
    TotalType(VisualisationDonnées);
    FormatageCouleurLignes(VisualisationDonnées);
  }

  internal void AfFicheMatière(ComboBox CB1,string tableSelection) {
    if (!_serviceWorkbookLoaded)
      return;

    ClasseBaseDeDonnées.DataSetExcel = DataSetExcelService;

    ClasseBaseDeDonnées.LireDonnées(VisualisationDonnées,tableSelection);
    TotalType(VisualisationDonnées);
    FormatageCouleurLignes(VisualisationDonnées);
  }

  internal void ControleCouleurBoutonsTDetTP() {
    ushort td = ushort.Parse(nbGroupeTd.Content?.ToString() ?? "0");
    ushort tdTotal = ushort.Parse(nbGroupeTdTotal.Content?.ToString() ?? "0");
    ushort tp = ushort.Parse(nbGroupeTp.Content?.ToString() ?? "0");
    ushort tpTotal = ushort.Parse(nbGroupeTpTotal.Content?.ToString() ?? "0");

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

    string codeCourtModule = ClasseBaseDeDonnées.ExtraireCodeCourtModuleDepuisLibelleCourt(libelleCourt);

    if (string.IsNullOrWhiteSpace(codeOse))
      codeOse = ClasseBaseDeDonnées.ConvertirCodeCourtEnCodeOse(codeCourtModule);

    TexteMatière.Content = module;
    TexteCodeOse.Content = codeOse;
    TexteLibelléModule.Text = codeCourtModule;

    ClasseBaseDeDonnées.CompterLesTypesDeCours(VisualisationDonnées);

    nbGroupeTd.Content = ClasseBaseDeDonnées.CompteurTD.ToString();
    nbGroupeTp.Content = ClasseBaseDeDonnées.CompteurTP.ToString();
    nbGroupeTdTotal.Content = FormationGEII.FI_Geii_1.NbGroupeTD.ToString();
    nbGroupeTpTotal.Content = FormationGEII.FI_Geii_1.NbGroupeTP.ToString();
    nbEtudiants.Content = FormationGEII.FI_Geii_1.NbEtudiants.ToString();

    ControleCouleurBoutonsTDetTP();

    BoutonsDeModification.Visibility = Visibility.Visible;
  }

  private void HorairesModule() {
    string nomModule = RechercheNomModule();

    Valeur_CM.Content = ClasseBaseDeDonnées.RechercheValeur(DataSetExcelSélection,nomModule,"CM","Module") + " h";
    Valeur_TD.Content = ClasseBaseDeDonnées.RechercheValeur(DataSetExcelSélection,nomModule,"TD","Module") + " h";
    Valeur_TP.Content = ClasseBaseDeDonnées.RechercheValeur(DataSetExcelSélection,nomModule,"TP","Module") + " h";
  }

  private static string HorairesModule(string NomModule,string NomCours) {
    String ValeurHoraireCours = ClasseBaseDeDonnées.RechercheValeur(NomModule,NomCours,"Module") + " h";
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

  private void RestaurerGroupesDepuisProfil() {
    if (_profile == null || _profile.Groups == null)
      return;

    if (!_profile.Groups.IsConfigured)
      return;

    FormationGEII.FI_Geii_1.NbEtudiants = _profile.Groups.StudentCount;
    FormationGEII.FI_Geii_1.NbGroupeTD = _profile.Groups.TdGroupCount;
    FormationGEII.FI_Geii_1.NbGroupeTP = _profile.Groups.TpGroupCount;
    NomGroupeTD = _profile.Groups.TdGroupNames ?? [];
    NomGroupeTP = _profile.Groups.TpGroupNames ?? [];
    étatEffectifGeii1.Text = FormationGEII.FI_Geii_1.NbEtudiants.ToString() + " étudiants : ";
    étatNbGroupeTD.Text = FormationGEII.FI_Geii_1.NbGroupeTD + " groupes de TD, ";
    étatNbGroupeTP.Text = FormationGEII.FI_Geii_1.NbGroupeTP + " groupes de TP\t";
    étatTitreEffectifGeii1.Visibility
      = étatEffectifGeii1.Visibility
      = étatNbGroupeTD.Visibility
      = étatNbGroupeTP.Visibility
      = Visibility.Collapsed;
  }

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

  private void MiseAJourTableau_Click(object sender,RoutedEventArgs e) {

  }

  private void BoutonModifier_Click(object sender,RoutedEventArgs e) {

  }

  private void BoutonAnnuler_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuFichierRéinitialiserEnvironnement_Click(object sender,RoutedEventArgs e) {

  }

  private void boutonAjouterCM_Click(object sender,RoutedEventArgs e) {

  }

  private void boutonSupprimerCM_Click(object sender,RoutedEventArgs e) {

  }

  private void boutonSupprimerTD_Click(object sender,RoutedEventArgs e) {

  }

  private void boutonAjouterTP_Click(object sender,RoutedEventArgs e) {

  }

  private void boutonSupprimerTP_Click(object sender,RoutedEventArgs e) {

  }

  private void ModifierFicheMatière_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuOuvrirFichier_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuFichierEnregistrer_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuFichierEnregistrerSous_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuImprimerFicheService_Click(object sender,RoutedEventArgs e) {

  }

  private void MenuApercuFicheService_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuQuitter_Click(object sender,RoutedEventArgs e) {

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


  private void AfficherLesCritèresDeSélection_Click(object sender,RoutedEventArgs e) {
    AfficherLesCritèresSélection();
  }

  private void élémentMenuGEII1_Click(object sender,RoutedEventArgs e) {

  }

  private void élémentMenuGEII2_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuGestionDesIntervenants_Click(object sender,RoutedEventArgs e) {

  }

  private void optionsToolStripMenuItem_Click(object sender,RoutedEventArgs e) {

  }

  private void gestionDesCoursToolStripMenuItem_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuVueGestionServices_Click(object sender,RoutedEventArgs e) {
    AfficherVueGestionServices();
  }

  private void mnuGenererStatutsIntervenants_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuCleanCache_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuMettreAJourSource_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuTest_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuCheminsAcces_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuVueFichesService_Click(object sender,RoutedEventArgs e) {
    AfficherVueFichesService();
  }

  private void btnVueFichesService_Click(object sender,RoutedEventArgs e) {
    AfficherVueFichesService();
  }

  private void mnuOpenProfile_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuSaveProfile_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuSaveProfileAs_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuProfileStatus_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuOpenLastProfile_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuAutoOpenLastProfileOnStartup_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuloadProfil_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuAbout_Click(object sender,RoutedEventArgs e) {

  }

  private void mnuCheckForUpdates_Click(object sender,RoutedEventArgs e) {

  }

  private void btnCreateServiceFiche_Click(object sender,RoutedEventArgs e) {

  }

  private void btnGestionService_Click(object sender,RoutedEventArgs e) {

  }

  private void btnGestionDesIntervenants_Click(object sender,RoutedEventArgs e) {

  }

  private void btnExporterFicheService_Click(object sender,RoutedEventArgs e) {

  }

  private void btnActualiserFicheService_Click(object sender,RoutedEventArgs e) {

  }

  private void btnImprimerFicheService_Click(object sender,RoutedEventArgs e) {

  }

  private void DialogueServiceGEII_Shown(object sender,EventArgs e) {
    if (!_applicationSettings.AutoOpenLastProfileOnStartup)
      return;

    string lastProfilePath =
        _applicationSettings.LastProfilePath;

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
      AfficherListesDonnées(Module,Semestre);
    }
    finally {
      _suspendViewRefresh = false;
    }
  }

  private void UE_SelectionChanged(object sender,SelectionChangedEventArgs e) {
    // UE n'est plus utilisé comme filtre principal.
  }

  private void Module_SelectionChanged(object sender,SelectionChangedEventArgs e) {
    if (_suspendViewRefresh)
      return;

    ActualiserComboOseDepuisModule(Module.Text.Trim());
    ActualiserComboInfosDepuisGlobal(Module.Text.Trim());

    if (!_isInitializingView)
      RequestRefreshModuleView();
  }
}