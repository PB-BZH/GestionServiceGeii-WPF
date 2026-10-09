using System.Data;
using System.Windows;
using System.Windows.Controls;
using GestionServiceGeii.Core.Groups;
using GestionServiceGeii.Shared.Database;
using PB.BZH.Theme.Theming;

namespace GestionServiceGeii.UI.Windows {
  public partial class GestionDesGroupes: Window {
    private DataGrid? _vueCourante;
    private string _groupeCourant = string.Empty;
    private string _typeCoursCourant = string.Empty;
    private bool _selectionTypeCoursEffectuee;
    internal const string GroupeEnAttente = "";
    private bool _autoriserGroupesDejaAffectes;
    private string[] _nomGroupeTD = [];
    private string[] _nomGroupeTP = [];

    public GestionDesGroupes() {
      InitializeComponent();
      ckBoxCours.IsChecked = false;
      ckBoxTD.IsChecked = false;
      ckBoxTP.IsChecked = false;
      ThemeManager.ApplyTheme(this);
    }

    internal void InitialiserContexteGroupes(
      DataGrid vueCourante,
      string groupeCourant,
      string[] nomGroupeTD,
      string[] nomGroupeTP,
      bool autoriserGroupesDejaAffectes = false) {

      _vueCourante = vueCourante;
      _groupeCourant = groupeCourant ?? string.Empty;
      _nomGroupeTD = nomGroupeTD ?? [];
      _nomGroupeTP = nomGroupeTP ?? [];
      _typeCoursCourant = DeterminerTypeCoursCourant();
      _autoriserGroupesDejaAffectes = autoriserGroupesDejaAffectes;
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
          ckBoxCours.IsChecked = true;
          break;
        case "TD":
          ckBoxTD.IsChecked = true;
          break;
        case "TP":
          ckBoxTP.IsChecked = true;
          break;
        default:
          break;
      }
    }

    private void SelectionnerGroupeCourantOuPremier() {
      if (!string.IsNullOrWhiteSpace(_groupeCourant) && CbChoixGroupe.Items.Contains(_groupeCourant)) {
        CbChoixGroupe.SelectedItem = _groupeCourant;
        return;
      }
      if (CbChoixGroupe.Items.Count > 0) CbChoixGroupe.SelectedIndex = 0;
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

    private void CkBoxCours_CheckedChanged(object sender,RoutedEventArgs e) {
      if (ckBoxTD.IsChecked == true)
        ckBoxTD.IsChecked = false;
      else if (ckBoxTP.IsChecked == true) ckBoxTP.IsChecked = false;
      else ckBoxCours.IsChecked = true;

      CbChoixGroupe.Items.Clear();
      CbChoixGroupe.Items.Add(GroupeEnAttente);
      CbChoixGroupe.Items.Add("Promo");

      SelectionnerGroupeCourantOuPremier();
    }

    private void ChargerGroupesDisponibles(string[] nomGroupe,string cours) {
      CbChoixGroupe.Items.Clear();
      CbChoixGroupe.Items.Add(GroupeEnAttente);

      HashSet<string> groupesDejaAffectes = RechercherGroupesDejaAffectes(cours);

      for (int i = 0;i < nomGroupe.Length;i++) {
        string groupe = nomGroupe[i];

        if (_autoriserGroupesDejaAffectes || GroupeDisponible(groupe,groupesDejaAffectes))
          CbChoixGroupe.Items.Add(groupe);
      }

      if (CbChoixGroupe.Items.Count > 0)
        CbChoixGroupe.SelectedIndex = 0;
    }

    private void CkBoxTD_CheckedChanged(object sender,RoutedEventArgs e) {
      if (ckBoxCours.IsChecked == true) ckBoxCours.IsChecked = false;
      else if (ckBoxTP.IsChecked == true) ckBoxTP.IsChecked = false;
      else ckBoxTD.IsChecked = true;

      ChargerGroupesDisponibles(_nomGroupeTD,"TD");
      SelectionnerGroupeCourantOuPremier();
    }

    private void CkBoxTP_CheckedChanged(object sender,RoutedEventArgs e) {
      if (ckBoxCours.IsChecked == true) ckBoxCours.IsChecked = false;
      else if (ckBoxTD.IsChecked == true) ckBoxTD.IsChecked = false;
      else ckBoxTP.IsChecked = true;

      ChargerGroupesDisponibles(_nomGroupeTP,"TP");
      SelectionnerGroupeCourantOuPremier();
    }

    private void ValidationTypeDeCours_Click(object sender,RoutedEventArgs e) {
      DialogResult = true;
    }
  }
}

