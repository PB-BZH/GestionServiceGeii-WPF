using System.IO;
using System.Windows;
using System.Windows.Controls;
using GestionServiceGeii.Core.Profiles;
using Microsoft.Win32;
using PB.BZH.Theme.Theming;

namespace GestionServiceGeii.UI.Windows;

/// <summary>
/// Logique d'interaction pour ProfilePathsWindow.xaml
/// </summary>
public partial class ProfilePathsWindow: Window {
  private readonly ServiceManagerProfile _profile;

  public ProfilePathsWindow(ServiceManagerProfile profile) {
    InitializeComponent();
    ArgumentNullException.ThrowIfNull(profile);
    _profile = profile;
    ThemeManager.ApplyTheme(this);
    ApplyProfileToUi();
  }

  private void ApplyProfileToUi() {
    txtListsWorkbookPath.Text = _profile.Files.ListsWorkbookPath;
    txtServiceWorkbookPath.Text = _profile.Files.ServiceWorkbookPath;
    txtBackupDirectory.Text = _profile.Files.BackupDirectory;
    txtExportDirectory.Text = _profile.Files.ExportDirectory;
  }

  private void ApplyUiToProfile() {
    _profile.Files.ListsWorkbookPath = txtListsWorkbookPath.Text.Trim();
    _profile.Files.ServiceWorkbookPath = txtServiceWorkbookPath.Text.Trim();
    _profile.Files.BackupDirectory = txtBackupDirectory.Text.Trim();
    _profile.Files.ExportDirectory = txtExportDirectory.Text.Trim();
  }

  private void SelectExcelFile(TextBox targetTextBox,string title) {
    OpenFileDialog dialog = new();
    dialog.Title = title;

    dialog.Filter =
        "Fichiers ExcelApp (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|" +
        "Tous les fichiers (*.*)|*.*";

    if (!string.IsNullOrWhiteSpace(targetTextBox.Text)) {
      string? directory = Path.GetDirectoryName(targetTextBox.Text);

      if (Directory.Exists(directory))
        dialog.InitialDirectory = directory;
    }

    if (dialog.ShowDialog(this) != true)
      return;

    targetTextBox.Text = dialog.FileName;
  }

  private void SelectDirectory(TextBox targetTextBox,string description) {
    OpenFolderDialog dialog = new();
    dialog.Title = description;

    if (!string.IsNullOrWhiteSpace(targetTextBox.Text) && Directory.Exists(targetTextBox.Text)) {
      dialog.DefaultDirectory = targetTextBox.Text;
    }

    if (dialog.ShowDialog(this) != true)
      return;

    targetTextBox.Text = dialog.FolderName;
  }

  private bool ValidatePaths() {
    if (string.IsNullOrWhiteSpace(txtListsWorkbookPath.Text)) {
      MessageBox.Show(
          this,
          "Le fichier de sélection n'est pas défini.",
          "Chemins d'accès",
          MessageBoxButton.OK,
          MessageBoxImage.Warning
      );

      txtListsWorkbookPath.Focus();
      return false;
    }

    if (!File.Exists(txtListsWorkbookPath.Text)) {
      MessageBox.Show(
          this,
          "Le fichier de sélection est introuvable.",
          "Chemins d'accès",
          MessageBoxButton.OK,
          MessageBoxImage.Warning
      );

      txtListsWorkbookPath.Focus();
      return false;
    }

    if (string.IsNullOrWhiteSpace(txtServiceWorkbookPath.Text)) {
      MessageBox.Show(
          this,
          "Le fichier de service n'est pas défini.",
          "Chemins d'accès",
          MessageBoxButton.OK,
          MessageBoxImage.Warning
      );

      txtServiceWorkbookPath.Focus();
      return false;
    }

    if (!File.Exists(txtServiceWorkbookPath.Text)) {
      MessageBox.Show(
          this,
          "Le fichier de service est introuvable.",
          "Chemins d'accès",
          MessageBoxButton.OK,
          MessageBoxImage.Warning
      );

      txtServiceWorkbookPath.Focus();
      return false;
    }

    if (!string.IsNullOrWhiteSpace(txtBackupDirectory.Text) &&
        !Directory.Exists(txtBackupDirectory.Text)) {
      MessageBox.Show(
          this,
          "Le dossier de sauvegarde est introuvable.",
          "Chemins d'accès",
          MessageBoxButton.OK,
          MessageBoxImage.Warning
      );

      txtBackupDirectory.Focus();
      return false;
    }

    if (!string.IsNullOrWhiteSpace(txtExportDirectory.Text) &&
        !Directory.Exists(txtExportDirectory.Text)) {
      MessageBox.Show(
          this,
          "Le dossier d'export est introuvable.",
          "Chemins d'accès",
          MessageBoxButton.OK,
          MessageBoxImage.Warning
      );

      txtExportDirectory.Focus();
      return false;
    }

    return true;
  }

  private void btnBrowseListsWorkbook_Click(object sender,RoutedEventArgs e) {
    SelectExcelFile(txtListsWorkbookPath,"Sélectionner le fichier de sélection");
  }

  private void btnBrowseServiceWorkbook_Click(object sender,RoutedEventArgs e) {
    SelectExcelFile(txtServiceWorkbookPath,"Sélectionner le fichier de service");
  }

  private void btnBrowseBackupDirectory_Click(object sender,RoutedEventArgs e) {
    SelectDirectory(txtBackupDirectory,"Sélectionner le dossier de sauvegarde");
  }

  private void btnBrowseExportDirectory_Click(object sender,RoutedEventArgs e) {
    SelectDirectory(txtExportDirectory,"Sélectionner le dossier d'export");
  }

  private void btnOk_Click(object sender,RoutedEventArgs e) {
    if (!ValidatePaths())
      return;

    ApplyUiToProfile();

    DialogResult = true;
  }

  private void btnCancel_Click(object sender,RoutedEventArgs e) {
    DialogResult = false;
  }
}
