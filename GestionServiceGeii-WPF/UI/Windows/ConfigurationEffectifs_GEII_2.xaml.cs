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
  private int nbGroupeTPSp;
  public int nbGroupeTD_officiel_FI;
  public string[] NomGroupeTD_FI = new string[3];
  public string[] NomGroupeTD_officiel_FI = new string[3];
  public string[] NomGroupeTP_FI = new string[6];
  public string[] NomGroupeTPSp_FI = new string[6];
  public int nbGroupeTD_officiel_FA;
  public string[] NomGroupeTD_FA = new string[3];
  public string[] NomGroupeTD_officiel_FA = new string[3];
  public string[] NomGroupeTP_FA = new string[6];
  public string[] NomGroupeTPSp_FA = new string[6];
  private readonly ServiceManagerProfile _profile;
  private bool _configurationTDInitialisee;

  public ConfigurationEffectifs_GEII_2(ServiceManagerProfile profile) {
    InitializeComponent();
    ThemeManager.ApplyTheme(this);
    chkFi.IsChecked = true;
    _profile = profile;

    _configurationTDInitialisee = true;
    CalculEffectifGroupeTD_FI();
  }

  private void ActualiserGroupesTp() {
    int.TryParse(txtNbEtudiantsEseFi.Text,out int nbEseFi);
    int.TryParse(txtNbEtudiantsEmeFi.Text,out int nbEmeFi);
    int.TryParse(txtNbEtudiantsAiiFi.Text,out int nbAiiFi);

    int.TryParse(txtNbEtudiantsEseFa.Text,out int nbEseFa);
    int.TryParse(txtNbEtudiantsEmeFa.Text,out int nbEmeFa);
    int.TryParse(txtNbEtudiantsAiiFa.Text,out int nbAiiFa);

    ActualiserGroupeTp(txtTpCommunEse21,tbkTpCommunEse21,nbEseFi > 0);
    ActualiserGroupeTp(txtTpCommunEse22,tbkTpCommunEse22,nbEseFi > 14);
    ActualiserGroupeTp(txtTpSpecialiséEse21,tbkTpSpecialiséEse21,nbEseFi > 0);
    ActualiserGroupeTp(txtTpSpecialiséEse22,tbkTpSpecialiséEse22,nbEseFi > 14);

    ActualiserGroupeTp(txtTpCommunEme21,tbkTpCommunEme21,nbEmeFi > 0);
    ActualiserGroupeTp(txtTpCommunEme22,tbkTpCommunEme22,nbEmeFi > 14);
    ActualiserGroupeTp(txtTpSpecialiséEme21,tbkTpSpecialiséEme21,nbEmeFi > 0);
    ActualiserGroupeTp(txtTpSpecialiséEme22,tbkTpSpecialiséEme22,nbEmeFi > 14);

    ActualiserGroupeTp(txtTpCommunAii21,tbkTpCommunAii21,nbAiiFi > 0);
    ActualiserGroupeTp(txtTpCommunAii22,tbkTpCommunAii22,nbAiiFi > 14);
    ActualiserGroupeTp(txtTpSpecialiséAii21,tbkTpSpecialiséAii21,nbAiiFi > 0);
    ActualiserGroupeTp(txtTpSpecialiséAii22,tbkTpSpecialiséAii22,nbAiiFi > 14);

    ActualiserGroupeTp(txtTpCommunEse21_Fa,tbkTpCommunEse21_Fa,nbEseFa > 0);
    ActualiserGroupeTp(txtTpCommunEse22_Fa,tbkTpCommunEse22_Fa,nbEseFa > 14);
    ActualiserGroupeTp(txtTpSpecialiséEse21_Fa,tbkTpSpecialiséEse21_Fa,nbEseFa > 0);
    ActualiserGroupeTp(txtTpSpecialiséEse22_Fa,tbkTpSpecialiséEse22_Fa,nbEseFa > 14);

    ActualiserGroupeTp(txtTpCommunEme21_Fa,tbkTpCommunEme21_Fa,nbEmeFa > 0);
    ActualiserGroupeTp(txtTpCommunEme22_Fa,tbkTpCommunEme22_Fa,nbEmeFa > 14);
    ActualiserGroupeTp(txtTpSpecialiséEme21_Fa,tbkTpSpecialiséEme21_Fa,nbEmeFa > 0);
    ActualiserGroupeTp(txtTpSpecialiséEme22_Fa,tbkTpSpecialiséEme22_Fa,nbEmeFa > 14);

    ActualiserGroupeTp(txtTpCommunAii21_Fa,tbkTpCommunAii21_Fa,nbAiiFa > 0);
    ActualiserGroupeTp(txtTpCommunAii22_Fa,tbkTpCommunAii22_Fa,nbAiiFa > 14);
    ActualiserGroupeTp(txtTpSpecialiséAii21_Fa,tbkTpSpecialiséAii21_Fa,nbAiiFa > 0);
    ActualiserGroupeTp(txtTpSpecialiséAii22_Fa,tbkTpSpecialiséAii22_Fa,nbAiiFa > 14);
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

  private void EffectifParcours_TextChanged(object sender,TextChangedEventArgs e) {
    ActualiserGroupesTp();
  }

  private void ActualiserGroupesParcours(
    CheckBox chkParcours,
    TextBox txtNbEtudiants,
    TextBlock tbkTpCommun1,
    TextBlock tbkTpCommun2,
    TextBlock tbkTpSpecialisé1,
    TextBlock tbkTpSpecialisé2,
    string[] nomGroupeTD,
    string[] nomGroupeTP,
    string[] nomGroupeTPSp,
    ref int i,
    ref int j,
    ref int k) {

    if (chkParcours.IsChecked != true)
      return;

    nbGroupeTD++;
    nomGroupeTD[i++] = chkParcours.Content?.ToString() ?? string.Empty;

    if (int.TryParse(txtNbEtudiants.Text,out int nbEtudiants) && nbEtudiants > 0) {
      nbGroupeTP++;
      nbGroupeTPSp++;
      nomGroupeTP[j++] = tbkTpCommun1.Text;
      nomGroupeTPSp[k++] = tbkTpSpecialisé1.Text;

      if (nbEtudiants > 14) {
        nbGroupeTP++;
        nbGroupeTPSp++;
        nomGroupeTP[j++] = tbkTpCommun2.Text;
        nomGroupeTPSp[k++] = tbkTpSpecialisé2.Text;
      }
    }
  }

  private void EnregistrerConfigurationEffectifs_FI() {
    int i = 0, j = 0, k = 0;
    nbGroupeTD = 0;
    nbGroupeTP = 0;
    nbGroupeTPSp = 0;

    NomGroupeTD_FI = new string[3];
    NomGroupeTP_FI = new string[6];
    NomGroupeTPSp_FI = new string[6];

    ActualiserGroupesParcours(chkParcoursEseFi,txtNbEtudiantsEseFi,tbkTpCommunEse21,tbkTpCommunEse22,
        tbkTpSpecialiséEse21,tbkTpSpecialiséEse22,NomGroupeTD_FI,NomGroupeTP_FI,NomGroupeTPSp_FI,ref i,ref j,ref k);

    ActualiserGroupesParcours(chkParcoursEmeFi,txtNbEtudiantsEmeFi,tbkTpCommunEme21,tbkTpCommunEme22,
        tbkTpSpecialiséEme21,tbkTpSpecialiséEme22,NomGroupeTD_FI,NomGroupeTP_FI,NomGroupeTPSp_FI,ref i,ref j,ref k);

    ActualiserGroupesParcours(chkParcoursAiiFi,txtNbEtudiantsAiiFi,tbkTpCommunAii21,tbkTpCommunAii22,
        tbkTpSpecialiséAii21,tbkTpSpecialiséAii22,NomGroupeTD_FI,NomGroupeTP_FI,NomGroupeTPSp_FI,ref i,ref j,ref k);

    CalculEffectifGroupeTD_FI();

    nbEtudiant = int.TryParse(txtNbEtudiantsFi.Text,out int effectif) ? effectif : 0;

    FormationGEII.Geii_2_FI.NbEtudiants = nbEtudiant;
    FormationGEII.Geii_2_FI.NbGroupeTD = nbGroupeTD;
    FormationGEII.Geii_2_FI.NbGroupeTD_officiel = nbGroupeTD_officiel_FI;
    FormationGEII.Geii_2_FI.NbGroupeTP = nbGroupeTP;
    FormationGEII.Geii_2_FI.NbGroupeTPSp = nbGroupeTPSp;

    _profile.Groups.Geii2_FI_StudentCount = nbEtudiant;
    _profile.Groups.Geii2_FI_TdGroupCount = nbGroupeTD;
    _profile.Groups.Geii2_FI_Td_officiel_GroupCount = nbGroupeTD_officiel_FI;
    _profile.Groups.Geii2_FI_TpGroupCount = nbGroupeTP;
    _profile.Groups.Geii2_FI_TpSpGroupCount = nbGroupeTPSp;

    _profile.Groups.Geii2_FI_TdGroupNames = (string[])NomGroupeTD_FI.Clone();
    _profile.Groups.Geii2_FI_Td_officiel_GroupNames = (string[])NomGroupeTD_officiel_FI.Clone();
    _profile.Groups.Geii2_FI_TpGroupNames = (string[])NomGroupeTP_FI.Clone();
    _profile.Groups.Geii2_FI_TpSpGroupNames = (string[])NomGroupeTPSp_FI.Clone();
  }

  private void EnregistrerConfigurationEffectifs_FA() {
    int i = 0, j = 0, k = 0;
    nbGroupeTD = 0;
    nbGroupeTP = 0;
    nbGroupeTPSp = 0;

    NomGroupeTD_FA = new string[3];
    NomGroupeTP_FA = new string[6];
    NomGroupeTPSp_FA = new string[6];

    ActualiserGroupesParcours(chkParcoursEseFa,txtNbEtudiantsEseFa,tbkTpCommunEse21_Fa,tbkTpCommunEse22_Fa,
        tbkTpSpecialiséEse21_Fa,tbkTpSpecialiséEse22_Fa,NomGroupeTD_FA,NomGroupeTP_FA,NomGroupeTPSp_FA,ref i,ref j,ref k);

    ActualiserGroupesParcours(chkParcoursEmeFa,txtNbEtudiantsEmeFa,tbkTpCommunEme21_Fa,tbkTpCommunEme22_Fa,
        tbkTpSpecialiséEme21_Fa,tbkTpSpecialiséEme22_Fa,NomGroupeTD_FA,NomGroupeTP_FA,NomGroupeTPSp_FA,ref i,ref j,ref k);

    ActualiserGroupesParcours(chkParcoursAiiFa,txtNbEtudiantsAiiFa,tbkTpCommunAii21_Fa,tbkTpCommunAii22_Fa,
        tbkTpSpecialiséAii21_Fa,tbkTpSpecialiséAii22_Fa,NomGroupeTD_FA,NomGroupeTP_FA,NomGroupeTPSp_FA,ref i,ref j,ref k);

    CalculEffectifGroupeTD_FA();

    nbEtudiant = int.TryParse(txtNbEtudiantsFa.Text,out int effectif) ? effectif : 0;

    FormationGEII.Geii_2_FA.NbEtudiants = nbEtudiant;
    FormationGEII.Geii_2_FA.NbGroupeTD = nbGroupeTD;
    FormationGEII.Geii_2_FA.NbGroupeTD_officiel = nbGroupeTD_officiel_FA;
    FormationGEII.Geii_2_FA.NbGroupeTP = nbGroupeTP;
    FormationGEII.Geii_2_FA.NbGroupeTPSp = nbGroupeTPSp;

    _profile.Groups.Geii2_FA_StudentCount = nbEtudiant;
    _profile.Groups.Geii2_FA_TdGroupCount = nbGroupeTD;
    _profile.Groups.Geii2_FA_Td_officiel_GroupCount = nbGroupeTD_officiel_FA;
    _profile.Groups.Geii2_FA_TpGroupCount = nbGroupeTP;
    _profile.Groups.Geii2_FA_TpSpGroupCount = nbGroupeTPSp;

    _profile.Groups.Geii2_FA_TdGroupNames = (string[])NomGroupeTD_FA.Clone();
    _profile.Groups.Geii2_FA_Td_officiel_GroupNames = (string[])NomGroupeTD_officiel_FA.Clone();
    _profile.Groups.Geii2_FA_TpGroupNames = (string[])NomGroupeTP_FA.Clone();
    _profile.Groups.Geii2_FA_TpSpGroupNames = (string[])NomGroupeTPSp_FA.Clone();
  }

  private string[] CalculEffectifGroupeTD_FI() {
    int.TryParse(txtNbEtudiantsEseFi.Text,out int effectifESE);
    int.TryParse(txtNbEtudiantsEmeFi.Text,out int effectifEME);
    int.TryParse(txtNbEtudiantsAiiFi.Text,out int effectifAII);

    int totalTD1 = 0;
    int totalTD2 = 0;
    int totalTD3 = 0;

    if (ChkESE.IsChecked == true)
      totalTD1 += effectifESE;

    if (ChkEME.IsChecked == true)
      totalTD1 += effectifEME;

    if (ChkAII.IsChecked == true)
      totalTD1 += effectifAII;

    if (ChkESE_2.IsChecked == true)
      totalTD2 += effectifESE;

    if (ChkEME_2.IsChecked == true)
      totalTD2 += effectifEME;

    if (ChkAII_2.IsChecked == true)
      totalTD2 += effectifAII;

    if (ChkESE_3.IsChecked == true)
      totalTD3 += effectifESE;

    if (ChkEME_3.IsChecked == true)
      totalTD3 += effectifEME;

    if (ChkAII_3.IsChecked == true)
      totalTD3 += effectifAII;

    txtTotalTD1_FI.Text = totalTD1.ToString();
    txtTotalTD2_FI.Text = totalTD2.ToString();
    txtTotalTD3_FI.Text = totalTD3.ToString();

    NomGroupeTD_officiel_FI = new string[3];

    if (totalTD1 > 0)
      NomGroupeTD_officiel_FI[0] = txtTd1_FI.Text;

    if (totalTD2 > 0)
      NomGroupeTD_officiel_FI[1] = txtTd2_FI.Text;

    if (totalTD3 > 0)
      NomGroupeTD_officiel_FI[2] = txtTd3_FI.Text;

    nbGroupeTD_officiel_FI = NomGroupeTD_officiel_FI.Count(g => !string.IsNullOrWhiteSpace(g));
    return NomGroupeTD_officiel_FI;
  }
  private string[] CalculEffectifGroupeTD_FA() {
    int.TryParse(txtNbEtudiantsEseFa.Text,out int effectifESE);
    int.TryParse(txtNbEtudiantsEmeFa.Text,out int effectifEME);
    int.TryParse(txtNbEtudiantsAiiFa.Text,out int effectifAII);

    int totalTD1 = 0;
    int totalTD2 = 0;
    int totalTD3 = 0;

    if (ChkESE_FA.IsChecked == true)
      totalTD1 += effectifESE;

    if (ChkEME_FA.IsChecked == true)
      totalTD1 += effectifEME;

    if (ChkAII_FA.IsChecked == true)
      totalTD1 += effectifAII;

    if (ChkESE_2_FA.IsChecked == true)
      totalTD2 += effectifESE;

    if (ChkEME_2_FA.IsChecked == true)
      totalTD2 += effectifEME;

    if (ChkAII_2_FA.IsChecked == true)
      totalTD2 += effectifAII;

    if (ChkESE_3_FA.IsChecked == true)
      totalTD3 += effectifESE;

    if (ChkEME_3_FA.IsChecked == true)
      totalTD3 += effectifEME;

    if (ChkAII_3_FA.IsChecked == true)
      totalTD3 += effectifAII;

    txtTotalTD1_FA.Text = totalTD1.ToString();
    txtTotalTD2_FA.Text = totalTD2.ToString();
    txtTotalTD3_FA.Text = totalTD3.ToString();

    NomGroupeTD_officiel_FA = new string[3];

    if (totalTD1 > 0)
      NomGroupeTD_officiel_FA[0] = txtTd1_FI.Text;

    if (totalTD2 > 0)
      NomGroupeTD_officiel_FA[1] = txtTd2_FI.Text;

    if (totalTD3 > 0)
      NomGroupeTD_officiel_FA[2] = txtTd3_FI.Text;

    nbGroupeTD_officiel_FA = NomGroupeTD_officiel_FA.Count(g => !string.IsNullOrWhiteSpace(g));
    return NomGroupeTD_officiel_FA;
  }


  private void btnUpdate_Click(object sender,RoutedEventArgs e) {
    EnregistrerConfigurationEffectifs_FI();
    EnregistrerConfigurationEffectifs_FA();
  }

  private void btnOk_Click(object sender,RoutedEventArgs e) {
    DialogResult = true;
  }

  private void GroupeTD_CheckedChanged(object sender,RoutedEventArgs e) {
    if (!_configurationTDInitialisee)
      return;

    NomGroupeTD_officiel_FI = CalculEffectifGroupeTD_FI();
  }

  private void GroupeTD_FA_CheckedChanged(object sender,RoutedEventArgs e) {
    if (!_configurationTDInitialisee)
      return;

    NomGroupeTD_officiel_FA = CalculEffectifGroupeTD_FA();
  }
}
