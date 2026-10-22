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
║  Le 22/10/2026 - 11:46
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
║  Nom de fichier : ChoixTypeDeCours.xaml.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Data;
using System.Windows;
using System.Windows.Controls;
using GestionServiceGeii.Core.Groups;
using GestionServiceGeii.Shared.Database;
using PB.BZH.Theme.Theming;

namespace GestionServiceGeii.UI.Windows {
  /// <summary>
  /// Logique d'interaction pour ChoixTypeDeCours.xaml
  /// </summary>
  public partial class ChoixTypeDeCours: Window {
    public string TypeDeCoursSelectionne { get; private set; } = string.Empty;
    public string GroupeSelectionne { get; private set; } = string.Empty;
    private DataGrid? _vueCourante;
    private string _groupeCourant = string.Empty;
    private string _typeCoursCourant = string.Empty;
    private bool _selectionTypeCoursEffectuee;
    internal const string GroupeEnAttente = "";
    private string _formationCourante = string.Empty;
    private string _parcoursCourant = string.Empty;
    private bool _coursMutualisé = false;

    public ChoixTypeDeCours() {
      InitializeComponent();
      ThemeManager.ApplyTheme(this);
    }

    private string[] ObtenirGroupesDisponibles(string cours,bool coursMutualisé) {
      bool specialite = _parcoursCourant is "Parcours ESE" or "Parcours EME" or "Parcours AII";

      if (_coursMutualisé) {
        return cours switch {
          "APP" => ["Promo"],
          "TD" => DialogueServiceGEII.NomGroupeTD_Geii2_Mutuel,
          "TP" when specialite => DialogueServiceGEII.NomGroupeTPSp_Geii2_Mutuel,
          "TP" => DialogueServiceGEII.NomGroupeTP_Geii2_Mutuel,
          _ => []
        };
      }

      if (_formationCourante == "FI") {
        return cours switch {
          "APP" => ["Promo"],
          "TD" => DialogueServiceGEII.NomGroupeTD_Geii2_officiel_FI,
          "TP" when specialite => DialogueServiceGEII.NomGroupeTPSp_officiel_Geii2_FI,
          "TP" => DialogueServiceGEII.NomGroupeTP_officiel_Geii2_FI,
          _ => []
        };
      }

      if (_formationCourante == "FA") {
        return cours switch {
          "APP" => ["Promo"],
          "TD" => DialogueServiceGEII.NomGroupeTD_Geii2_officiel_FA,
          "TP" when specialite => DialogueServiceGEII.NomGroupeTPSp_officiel_Geii2_FA,
          "TP" => DialogueServiceGEII.NomGroupeTP_officiel_Geii2_FA,
          _ => []
        };
      }

      return cours switch {
        "APP" => ["Promo"],
        "TD" => DialogueServiceGEII.NomGroupeTD_Geii1,
        "TP" => DialogueServiceGEII.NomGroupeTP_Geii1,
        _ => []
      };
    }

    internal void InitialiserContexteGroupes(DataGrid vueCourante,string groupeCourant = "",
      string formation = "",string parcours = "",bool coursMutualise = false) {
      _vueCourante = vueCourante;
      _groupeCourant = groupeCourant ?? string.Empty;
      _formationCourante = formation;
      _parcoursCourant = parcours;
      _coursMutualisé = coursMutualise;
      _typeCoursCourant = DeterminerTypeCoursCourant();
    }

    private void Window_ContentRendered(object? sender,EventArgs e) {
      if (_selectionTypeCoursEffectuee)
        return;

      _selectionTypeCoursEffectuee = true;
      SelectionnerTypeCoursCourant();
    }

    private string DeterminerTypeCoursCourant() {
      if (_vueCourante == null)
        return string.Empty;

      if (!_vueCourante.CurrentCell.IsValid)
        return string.Empty;

      DataRowView? ligne = _vueCourante.CurrentCell.Item as DataRowView;

      if (ligne == null)
        return string.Empty;

      return ClasseBaseDeDonnées.GetBoundRowValue_vue(ligne,ExcelSchemaNames.Columns.Cours).Trim();
    }

    private void SelectionnerTypeCoursCourant() {
      switch (_typeCoursCourant.ToUpperInvariant()) {
        case "CM":
          Bouton_CM.IsChecked = true;
          break;
        case "APP":
          Bouton_APP.IsChecked = true;
          break;
        case "TD":
          Bouton_TD.IsChecked = true;
          break;
        case "TP":
          Bouton_TP.IsChecked = true;
          break;
        default:
          break;
      }
    }

    private void SelectionnerGroupeCourantOuPremier() {
      if (!string.IsNullOrWhiteSpace(_groupeCourant) && ChoixGroupe.Items.Contains(_groupeCourant)) {
        ChoixGroupe.SelectedItem = _groupeCourant;
        return;
      }
      if (ChoixGroupe.Items.Count > 0) ChoixGroupe.SelectedIndex = 0;
    }

    private HashSet<string> RechercherGroupesDejaAffectes(string typeCours) {
      HashSet<string> groupesDejaAffectes = [];

      if (_vueCourante == null)
        return groupesDejaAffectes;

      foreach (DataRowView ligne in _vueCourante.Items) {
        string cours = ClasseBaseDeDonnées.GetBoundRowValue_vue(ligne,ExcelSchemaNames.Columns.Cours).Trim();
        string groupe = ClasseBaseDeDonnées.GetBoundRowValue_vue(ligne,ExcelSchemaNames.Columns.Groupe).Trim();

        if (string.IsNullOrWhiteSpace(cours) || string.IsNullOrWhiteSpace(groupe))
          continue;

        if (!string.Equals(cours,typeCours,StringComparison.OrdinalIgnoreCase))
          continue;

        if (string.Equals(groupe,GroupeEnAttente,StringComparison.OrdinalIgnoreCase))
          continue;

        foreach (string groupeAffecte in GroupAssignmentHelper.ExtractGroups(groupe))
          groupesDejaAffectes.Add(groupeAffecte);
      }

      return groupesDejaAffectes;
    }

    private bool GroupeDisponible(string groupe,HashSet<string> groupesDejaAffectes) {
      if (string.IsNullOrWhiteSpace(groupe)) return false;

      groupe = groupe.Trim();
      bool groupeDejaAffecte = groupesDejaAffectes.Contains(groupe);
      bool groupeCourant = string.Equals(groupe,_groupeCourant,StringComparison.OrdinalIgnoreCase);

      if (groupeDejaAffecte && !groupeCourant) return false;
      return true;
    }

    private void ChargerGroupesDisponibles(string[] nomGroupe,string cours) {
      ChoixGroupe.Items.Clear();
      ChoixGroupe.Items.Add(GroupeEnAttente);
      HashSet<string> groupesDejaAffectes = RechercherGroupesDejaAffectes(cours);

      for (int i = 0;i < nomGroupe.Length;i++) {
        string groupe = nomGroupe[i];
        if (GroupeDisponible(groupe,groupesDejaAffectes)) ChoixGroupe.Items.Add(groupe);
      }
      if (ChoixGroupe.Items.Count > 0) ChoixGroupe.SelectedIndex = 0;
    }


    private void Bouton_CM_Checked(object sender,RoutedEventArgs e) {
      if ((Bouton_TD.IsChecked == true) || (Bouton_TP.IsChecked == true)) {
        Bouton_TD.IsChecked = false;
        Bouton_TP.IsChecked = false;
        Bouton_APP.IsChecked = false;
      }
      else
        Bouton_APP.IsChecked = true;

      TypeDeCoursSelectionne = "CM";

      ChargerGroupesDisponibles(ObtenirGroupesDisponibles("CM",false),"CM");
      SelectionnerGroupeCourantOuPremier();
    }

    private void Bouton_APP_Checked(object sender,RoutedEventArgs e) {
      if ((Bouton_TD.IsChecked == true) || (Bouton_TP.IsChecked == true)) {
        Bouton_TD.IsChecked = false;
        Bouton_TP.IsChecked = false;
        Bouton_CM.IsChecked = false;
      }
      else
        Bouton_APP.IsChecked = true;

      TypeDeCoursSelectionne = "APP";

      ChargerGroupesDisponibles(ObtenirGroupesDisponibles("APP",false),"APP");
      SelectionnerGroupeCourantOuPremier();
    }

    private void Bouton_TD_Checked(object sender,RoutedEventArgs e) {
      if ((Bouton_CM.IsChecked == true) || (Bouton_TP.IsChecked == true)) {
        Bouton_CM.IsChecked = false;
        Bouton_TP.IsChecked = false;
        Bouton_APP.IsChecked = false;
      }
      else
        Bouton_TD.IsChecked = true;

      TypeDeCoursSelectionne = "TD";

      ChargerGroupesDisponibles(ObtenirGroupesDisponibles("TD",false),"TD");
      SelectionnerGroupeCourantOuPremier();
    }

    private void Bouton_TP_Checked(object sender,RoutedEventArgs e) {
      if ((Bouton_CM.IsChecked == true) || (Bouton_TD.IsChecked == true)) {
        Bouton_CM.IsChecked = false;
        Bouton_TD.IsChecked = false;
        Bouton_APP.IsChecked = false;
      }
      else
        Bouton_TP.IsChecked = true;

      TypeDeCoursSelectionne = "TP";

      ChargerGroupesDisponibles(ObtenirGroupesDisponibles("TP",true),"TP");
      SelectionnerGroupeCourantOuPremier();
    }

    private void ValidationTypeDeCours_Click(object sender,RoutedEventArgs e) {
      GroupeSelectionne = TypeDeCoursSelectionne == "CM"
          ? "Promo"
          : ChoixGroupe.SelectedItem?.ToString()?.Trim() ?? string.Empty;

      GroupeSelectionne = TypeDeCoursSelectionne == "APP"
          ? "Promo"
          : ChoixGroupe.SelectedItem?.ToString()?.Trim() ?? string.Empty;

      DialogResult = true;
    }
  }
}
