using System.Windows;
using System.Windows.Controls;
using GestionServiceGeii.Core.Profiles;
using GestionServiceGeii.Shared.Librairie_GEII;
using PB.BZH.Theme.Theming;

namespace GestionServiceGeii.UI.Windows;

/// <summary>
/// Logique d'interaction pour _configurationEffectifs_GEII_2.xaml
/// </summary>
public partial class ConfigurationEffectifs_GEII_2: Window {
  private int nbEtudiantMaxParGroupeTP;
  private int nbEtudiant = 0;
  private int nbGroupeTD;
  private int nbGroupeTP;
  public string[] NomGroupeTD = new string[3];
  public string[] NomGroupeTP = new string[6];
  private readonly ServiceManagerProfile _profile;

  public ConfigurationEffectifs_GEII_2(ServiceManagerProfile profile) {
    InitializeComponent();
    ThemeManager.ApplyTheme(this);
    chkFi.IsChecked = true;
    _profile = profile;
  }

  private void ActualiserGroupesTp() {
    int.TryParse(txtNbEtudiantsEseFi.Text,out int nbEse);
    int.TryParse(txtNbEtudiantsEmeFi.Text,out int nbEme);
    int.TryParse(txtNbEtudiantsAiiFi.Text,out int nbAii);

    ActualiserGroupeTp(txtTpCommunEse21,tbkTpCommunEse21,nbEse > 0);
    ActualiserGroupeTp(txtTpCommunEse22,tbkTpCommunEse22,nbEse > 14);
    ActualiserGroupeTp(txtTpSpecialiséEse21,tbkTpSpecialiséEse21,nbEse > 0);
    ActualiserGroupeTp(txtTpSpecialiséEse22,tbkTpSpecialiséEse22,nbEse > 14);

    ActualiserGroupeTp(txtTpCommunEme21,tbkTpCommunEme21,nbEme > 0);
    ActualiserGroupeTp(txtTpCommunEme22,tbkTpCommunEme22,nbEme > 14);
    ActualiserGroupeTp(txtTpSpecialiséEme21,tbkTpSpecialiséEme21,nbEme > 0);
    ActualiserGroupeTp(txtTpSpecialiséEme22,tbkTpSpecialiséEme22,nbEme > 14);

    ActualiserGroupeTp(txtTpCommunAii21,tbkTpCommunAii21,nbAii > 0);
    ActualiserGroupeTp(txtTpCommunAii22,tbkTpCommunAii22,nbAii > 14);
    ActualiserGroupeTp(txtTpSpecialiséAii21,tbkTpSpecialiséAii21,nbAii > 0);
    ActualiserGroupeTp(txtTpSpecialiséAii22,tbkTpSpecialiséAii22,nbAii > 14);
  }

  private static void ActualiserGroupeTp(TextBox textBox,TextBlock textBlock,bool visible) {
    Visibility visibility = visible ? Visibility.Visible : Visibility.Collapsed;

    textBox.Visibility = visibility;
    textBox.IsEnabled = visible;
    textBlock.Visibility = visibility;

    if (!visible)
      textBox.Clear();
  }

  private void chkFa_Checked(object sender,RoutedEventArgs e) {
    if (chkFa.IsChecked == true) {
      chkFi.IsChecked = false;
    }
    lblFormationName.Text = "Formation par apprentissage en GEII-2";
    pnlFa.Visibility = Visibility.Visible;
    pnlFi.Visibility = Visibility.Collapsed;
  }

  private void chkFi_Checked(object sender,RoutedEventArgs e) {
    if (chkFi.IsChecked == true) {
      chkFa.IsChecked = false;
    }
    lblFormationName.Text = "Formation initiale en GEII-2";
    pnlFa.Visibility = Visibility.Collapsed;
    pnlFi.Visibility = Visibility.Visible;
  }

  private void EffectifParcoursFi_TextChanged(object sender,TextChangedEventArgs e) {
    if (!IsLoaded)
      return;

    ActualiserGroupesTp();
  }

  private void ActualiserGroupesParcours(CheckBox chkParcours,TextBox txtNbEtudiants,TextBlock tbkTpCommun1,TextBlock tbkTpCommun2,ref int i,ref int j) {
    if (chkParcours.IsChecked == true) {
      nbGroupeTD += 1;
      NomGroupeTD[i++] = chkParcours.Content?.ToString() ?? string.Empty;
      if (int.TryParse(txtNbEtudiants.Text,out int nbEse) && nbEse <= 14) {
        nbGroupeTP += 1;
        NomGroupeTP[j++] = tbkTpCommun1.Text;
      }
      else {
        nbGroupeTP += 2;
        NomGroupeTP[j++] = tbkTpCommun1.Text;
        NomGroupeTP[j++] = tbkTpCommun2.Text;
      }
    }
  }

  private void btnSave_Click(object sender,RoutedEventArgs e) {
    int i = 0, j = 0;
    nbGroupeTD = 0;
    nbGroupeTP = 0;
    NomGroupeTD = new string[3];
    NomGroupeTP = new string[6];

    ActualiserGroupesParcours(chkParcoursEseFi,txtNbEtudiantsEseFi,tbkTpCommunEse21,tbkTpCommunEse22,ref i,ref j);
    ActualiserGroupesParcours(chkParcoursEmeFi,txtNbEtudiantsEmeFi,tbkTpCommunEme21,tbkTpCommunEme22,ref i,ref j);
    ActualiserGroupesParcours(chkParcoursAiiFi,txtNbEtudiantsAiiFi,tbkTpCommunAii21,tbkTpCommunAii22,ref i,ref j);

    nbEtudiant = int.TryParse(txtNbEtudiantsFi.Text,out int nbEtudiantFi) ? nbEtudiantFi : 0;

    FormationGEII.Geii_2_FI.NbEtudiants = nbEtudiant;
    FormationGEII.Geii_2_FI.NbGroupeTD = nbGroupeTD;
    FormationGEII.Geii_2_FI.NbGroupeTP = nbGroupeTP;

    _profile.Groups.Geii2_StudentCount = nbEtudiant;
    _profile.Groups.Geii2_TdGroupCount = nbGroupeTD;
    _profile.Groups.Geii2_TpGroupCount = nbGroupeTP;
    _profile.Groups.Geii2_TdGroupNames = NomGroupeTD;
    _profile.Groups.Geii2_TpGroupNames = NomGroupeTP;

    DialogResult = true;
  }
}
