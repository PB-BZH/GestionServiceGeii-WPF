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
║  Le 18/10/2026 - 12:04
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
║  Nom de fichier : ChoixSuppressionGroupe.xaml.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
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