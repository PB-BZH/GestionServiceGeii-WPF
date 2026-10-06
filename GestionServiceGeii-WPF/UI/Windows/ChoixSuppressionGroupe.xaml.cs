using System.Windows;
using PB.BZH.Theme.Theming;

namespace GestionServiceGeii.UI.Windows {
  public partial class ChoixSuppressionGroupe: Window {
    internal string GroupeSelectionne =>
        CbChoixGroupe.SelectedItem?.ToString() ?? string.Empty;

    public ChoixSuppressionGroupe() {
      InitializeComponent();
      ThemeManager.ApplyTheme(this);
    }

    internal void Initialiser(string? intervenant,string? cours,IEnumerable<string>? groupes) {
      TxtIntervenant.Text = intervenant ?? string.Empty;
      TxtCours.Text = cours ?? string.Empty;

      CbChoixGroupe.Items.Clear();

      if (groupes != null) {
        foreach (string groupe in groupes) {
          if (!string.IsNullOrWhiteSpace(groupe))
            CbChoixGroupe.Items.Add(groupe.Trim());
        }
      }

      if (CbChoixGroupe.Items.Count > 0)
        CbChoixGroupe.SelectedIndex = 0;
    }

    private void BtnValider_Click(object sender,RoutedEventArgs e) {
      if (CbChoixGroupe.SelectedItem == null) {
        MessageBox.Show(
            this,
            "Veuillez sélectionner un groupe à supprimer.",
            "Suppression groupe",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return;
      }

      DialogResult = true;
    }
  }
}