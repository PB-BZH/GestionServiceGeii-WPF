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
║  Nom de fichier : ConfigurationEffectifs_GEII_2.xaml.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
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
  private int nbGroupeCM;
  private int nbGroupeTD;
  private int nbGroupeTP;
  private int nbGroupeTPSp;
  public string[] NomGroupeCM_FI = [];
  public string[] NomGroupeTD_FI = [];
  public string[] NomGroupeTD_officiel_FI = [];
  public string[] NomGroupeTP_FI = [];
  public string[] NomGroupeTP_officiel_FI = [];
  public string[] NomGroupeTPSp_FI = [];
  public string[] NomGroupeTPSp_officiel_FI = [];
  public int nbGroupeTD_officiel_FI;
  public int nbGroupeTD_officiel_FA;
  public int nbGroupeTP_officiel_FI;
  public int nbGroupeTP_officiel_FA;
  public int nbGroupeTPSp_officiel_FI;
  public int nbGroupeTPSp_officiel_FA;
  public string[] NomGroupeTD_FA = [];
  public string[] NomGroupeTD_officiel_FA = [];
  public string[] NomGroupeTP_FA = [];
  public string[] NomGroupeTP_officiel_FA = [];
  public string[] NomGroupeTPSp_FA = [];
  public string[] NomGroupeTPSp_officiel_FA = [];
  private readonly ServiceManagerProfile _profile;
  private bool _configurationTDInitialisee;

  public ConfigurationEffectifs_GEII_2(ServiceManagerProfile profile) {
    InitializeComponent();
    ThemeManager.ApplyTheme(this);

    _profile = profile;
    _configurationTDInitialisee = false;

    ApplyProfileToUi(profile);

    _configurationTDInitialisee = true;

    ActualiserGroupesTp_FI();
    ActualiserGroupesTp_FA();

    CalculEffectifGroupeTD_FI();
    CalculEffectifGroupeTD_FA();

    CalculEffectifGroupeTP_FI();
    CalculEffectifGroupeTP_FA();

    CalculEffectifGroupeTP_Sp_FI();
    CalculEffectifGroupeTP_Sp_FA();

    chkFi.IsChecked = true;
  }

  private static bool[] LireEtatsCheckBox(params CheckBox[] cases) {
    return cases.Select(c => c.IsChecked == true).ToArray();
  }

  private static void RestaurerEtatsCheckBox(bool[]? etats,params CheckBox[] cases) {
    for (int i = 0;i < cases.Length;i++)
      cases[i].IsChecked = etats != null && i < etats.Length && etats[i];
  }

  private static void RestaurerEffectifs(int[]? effectifs,params TextBox[] controles) {
    if (effectifs == null)
      return;

    for (int i = 0;i < controles.Length && i < effectifs.Length;i++)
      controles[i].Text = effectifs[i].ToString();
  }

  private void ApplyProfileToUi(ServiceManagerProfile profile) {
    RestaurerEtatsCheckBox(_profile.Groups.Geii2_FI_Checks,chkParcoursEseFi,chkParcoursEmeFi,chkParcoursAiiFi);
    RestaurerEtatsCheckBox(_profile.Groups.Geii2_FA_Checks,chkParcoursEseFa,chkParcoursEmeFa,chkParcoursAiiFa);

    RestaurerEtatsCheckBox(_profile.Groups.Geii2_FI_Td_Checks,ChkESE_Fi,ChkEME_Fi,ChkAII_Fi,ChkESE_2_Fi,ChkEME_2_Fi,ChkAII_2_Fi,ChkESE_3_Fi,ChkEME_3_Fi,ChkAII_3_Fi);
    RestaurerEtatsCheckBox(_profile.Groups.Geii2_FA_Td_Checks,ChkESE_FA,ChkEME_FA,ChkAII_FA,ChkESE_2_FA,ChkEME_2_FA,ChkAII_2_FA,ChkESE_3_FA,ChkEME_3_FA,ChkAII_3_FA);

    RestaurerEtatsCheckBox(_profile.Groups.Geii2_FI_Tp_Checks,
      chkTP11_Fi,chkTP12_Fi,chkTP13_Fi,chkTP14_Fi,chkTP15_Fi,chkTP16_Fi,
      chkTP21_Fi,chkTP22_Fi,chkTP23_Fi,chkTP24_Fi,chkTP25_Fi,chkTP26_Fi,
      chkTP31_Fi,chkTP32_Fi,chkTP33_Fi,chkTP34_Fi,chkTP35_Fi,chkTP36_Fi);

    RestaurerEtatsCheckBox(_profile.Groups.Geii2_FA_Tp_Checks,
      chkTP11_Fa,chkTP12_Fa,chkTP13_Fa,chkTP14_Fa,chkTP15_Fa,chkTP16_Fa,
      chkTP21_Fa,chkTP24_Fa,chkTP23_Fa,chkTP24_Fa,chkTP25_Fa,chkTP26_Fa,
      chkTP31_Fa,chkTP32_Fa,chkTP33_Fa,chkTP34_Fa,chkTP35_Fa,chkTP36_Fa);

    RestaurerEtatsCheckBox(_profile.Groups.Geii2_FI_TpSp_Checks,
      chkTP11Sp_Fi,chkTP12Sp_Fi,chkTP13Sp_Fi,chkTP14Sp_Fi,chkTP15Sp_Fi,chkTP16Sp_Fi,
      chkTP21Sp_Fi,chkTP22Sp_Fi,chkTP23Sp_Fi,chkTP24Sp_Fi,chkTP25Sp_Fi,chkTP26Sp_Fi,
      chkTP31Sp_Fi,chkTP32Sp_Fi,chkTP33Sp_Fi,chkTP34Sp_Fi,chkTP35Sp_Fi,chkTP36Sp_Fi);

    RestaurerEtatsCheckBox(_profile.Groups.Geii2_FA_TpSp_Checks,
      chkTP11Sp_Fa,chkTP12Sp_Fa,chkTP13Sp_Fa,chkTP14Sp_Fa,chkTP15Sp_Fa,chkTP16Sp_Fa,
      chkTP21Sp_Fa,chkTP22Sp_Fa,chkTP23Sp_Fa,chkTP24Sp_Fa,chkTP25Sp_Fa,chkTP26Sp_Fa,
      chkTP31Sp_Fa,chkTP32Sp_Fa,chkTP33Sp_Fa,chkTP34Sp_Fa,chkTP35Sp_Fa,chkTP36Sp_Fa);


    RestaurerEffectifs(_profile.Groups.TotalGoupe_Geii2_Td_FI,txtTotalTD1_Fi,txtTotalTD2_Fi,txtTotalTD3_Fi);
    RestaurerEffectifs(_profile.Groups.TotalGoupe_Geii2_Td_FA,txtTotalTD1_FA,txtTotalTD2_FA,txtTotalTD3_FA);

    RestaurerEffectifs(
      _profile.Groups.TotalGroupe_Geii2_Tp_FI_officiel,
      txtTP1_officiel_Fi,
      txtTP2_officiel_Fi,
      txtTP3_officiel_Fi);
    RestaurerEffectifs(
      _profile.Groups.TotalGroupe_Geii2_Tp_FA_officiel,
      txtTP1_officiel_Fa,
      txtTP2_officiel_Fa,
      txtTP3_officiel_Fa);
    RestaurerEffectifs(
      _profile.Groups.TotalGroupe_Geii2_TpSp_FI_officiel,
      txtTP1Sp_officiel_Fi,
      txtTP2Sp_officiel_Fi,
      txtTP3Sp_officiel_Fi);
    RestaurerEffectifs(
      _profile.Groups.TotalGroupe_Geii2_TpSp_FA_officiel,
      txtTP1Sp_officiel_Fa,
      txtTP2Sp_officiel_Fa,
      txtTP3Sp_officiel_Fa);

    RestaurerEffectifs(_profile.Groups.Geii2_FI_TD_Effectifs,txtNbEtudiantsEseFi,txtNbEtudiantsEmeFi,txtNbEtudiantsAiiFi);
    RestaurerEffectifs(_profile.Groups.Geii2_FA_TD_Effectifs,txtNbEtudiantsEseFa,txtNbEtudiantsEmeFa,txtNbEtudiantsAiiFa);

    RestaurerEffectifs(_profile.Groups.Geii2_FI_Tp_Effectifs,txtTpCommunEse21_Fi,txtTpCommunEse22_Fi,txtTpCommunEme21_Fi,txtTpCommunEme22_Fi,txtTpCommunAii21_Fi,txtTpCommunAii22_Fi);
    RestaurerEffectifs(_profile.Groups.Geii2_FA_Tp_Effectifs,txtTpCommunEse23_Fa,txtTpCommunEse24_Fa,txtTpCommunEme23_Fa,txtTpCommunEme24_Fa,txtTpCommunAii23_Fa,txtTpCommunAii24_Fa);

    RestaurerEffectifs(_profile.Groups.Geii2_FI_TpSp_Effectifs,txtTpSpEse21_Fi,txtTpSpEse22_Fi,txtTpSpEme21_Fi,txtTpSpEme22_Fi,txtTpSpAii21_Fi,txtTpSpAii22_Fi);
    RestaurerEffectifs(_profile.Groups.Geii2_FA_TpSp_Effectifs,txtTpSpEse23_Fa,txtTpSpEse24_Fa,txtTpSpEme23_Fa,txtTpSpEme24_Fa,txtTpSpAii23_Fa,txtTpSpAii24_Fa);

    txtNbEtudiantsFi.Text = _profile.Groups.Geii2_FI_StudentCount.ToString();
    txtNbEtudiantsFa.Text = _profile.Groups.Geii2_FA_StudentCount.ToString();

    NomGroupeTP_officiel_FI = _profile.Groups.Geii2_FI_Tp_officiel_GroupNames;
    NomGroupeTP_officiel_FA = _profile.Groups.Geii2_FA_Tp_officiel_GroupNames;
  }

  private void ApplyUiToProfile(ServiceManagerProfile profile) {
    _profile.Groups.Geii2_FI_Checks = LireEtatsCheckBox(
      chkParcoursEseFi,
      chkParcoursEmeFi,
      chkParcoursAiiFi);
    _profile.Groups.Geii2_FA_Checks = LireEtatsCheckBox(
      chkParcoursEseFa,
      chkParcoursEmeFa,
      chkParcoursAiiFa);

    _profile.Groups.Geii2_FI_Td_Checks = LireEtatsCheckBox(
      ChkESE_Fi,ChkEME_Fi,ChkAII_Fi,
      ChkESE_2_Fi,ChkEME_2_Fi,ChkAII_2_Fi,
      ChkESE_3_Fi,ChkEME_3_Fi,ChkAII_3_Fi);
    _profile.Groups.Geii2_FA_Td_Checks = LireEtatsCheckBox(
      ChkESE_FA,ChkEME_FA,ChkAII_FA,
      ChkESE_2_FA,ChkEME_2_FA,ChkAII_2_FA,
      ChkESE_3_FA,ChkEME_3_FA,ChkAII_3_FA);

    _profile.Groups.Geii2_FI_Tp_Checks = LireEtatsCheckBox(
      chkTP11_Fi,chkTP12_Fi,chkTP13_Fi,chkTP14_Fi,chkTP15_Fi,chkTP16_Fi,
      chkTP21_Fi,chkTP22_Fi,chkTP23_Fi,chkTP24_Fi,chkTP25_Fi,chkTP26_Fi,
      chkTP31_Fi,chkTP32_Fi,chkTP33_Fi,chkTP34_Fi,chkTP35_Fi,chkTP36_Fi);
    _profile.Groups.Geii2_FA_Tp_Checks = LireEtatsCheckBox(
      chkTP11_Fa,chkTP12_Fa,chkTP13_Fa,chkTP14_Fa,chkTP15_Fa,chkTP16_Fa,
      chkTP21_Fa,chkTP24_Fa,chkTP23_Fa,chkTP24_Fa,chkTP25_Fa,chkTP26_Fa,
      chkTP31_Fa,chkTP32_Fa,chkTP33_Fa,chkTP34_Fa,chkTP35_Fa,chkTP36_Fa);

    _profile.Groups.Geii2_FI_TpSp_Checks = LireEtatsCheckBox(
      chkTP11Sp_Fi,chkTP12Sp_Fi,chkTP13Sp_Fi,chkTP14Sp_Fi,chkTP15Sp_Fi,chkTP16Sp_Fi,
      chkTP21Sp_Fi,chkTP22Sp_Fi,chkTP23Sp_Fi,chkTP24Sp_Fi,chkTP25Sp_Fi,chkTP26Sp_Fi,
      chkTP31Sp_Fi,chkTP32Sp_Fi,chkTP33Sp_Fi,chkTP34Sp_Fi,chkTP35Sp_Fi,chkTP36Sp_Fi);
    _profile.Groups.Geii2_FA_TpSp_Checks = LireEtatsCheckBox(
      chkTP11Sp_Fa,chkTP12Sp_Fa,chkTP13Sp_Fa,chkTP14Sp_Fa,chkTP15Sp_Fa,chkTP16Sp_Fa,
      chkTP21Sp_Fa,chkTP22Sp_Fa,chkTP23Sp_Fa,chkTP24Sp_Fa,chkTP25Sp_Fa,chkTP26Sp_Fa,
      chkTP31Sp_Fa,chkTP32Sp_Fa,chkTP33Sp_Fa,chkTP34Sp_Fa,chkTP35Sp_Fa,chkTP36Sp_Fa);

    _profile.Groups.Geii2_FI_TD_Effectifs = [
      int.TryParse(txtNbEtudiantsEseFi.Text,out int eseTd_FI) ? eseTd_FI : 0,
      int.TryParse(txtNbEtudiantsEmeFi.Text,out int emeTd_FI) ? emeTd_FI : 0,
      int.TryParse(txtNbEtudiantsAiiFi.Text,out int AiiTd_FI) ? AiiTd_FI : 0,
    ];
    _profile.Groups.Geii2_FA_TD_Effectifs = [
      int.TryParse(txtNbEtudiantsEseFa.Text,out int eseTd_FA) ? eseTd_FA : 0,
      int.TryParse(txtNbEtudiantsEmeFa.Text,out int emeTd_FA) ? emeTd_FA : 0,
      int.TryParse(txtNbEtudiantsAiiFa.Text,out int AiiTd_FA) ? AiiTd_FA : 0,
    ];
    _profile.Groups.Geii2_FI_Tp_Effectifs = [
      int.TryParse(txtTpCommunEse21_Fi.Text,out int eseTp_FI_21) ? eseTp_FI_21 : 0,
      int.TryParse(txtTpCommunEse22_Fi.Text,out int eseTp_FI_22) ? eseTp_FI_22 : 0,
      int.TryParse(txtTpCommunEme21_Fi.Text,out int emeTp_FI_21) ? emeTp_FI_21 : 0,
      int.TryParse(txtTpCommunEme22_Fi.Text,out int emeTp_FI_22) ? emeTp_FI_22 : 0,
      int.TryParse(txtTpCommunAii21_Fi.Text,out int aiiTp_FI_21) ? aiiTp_FI_21 : 0,
      int.TryParse(txtTpCommunAii22_Fi.Text,out int aiiTp_FI_22) ? aiiTp_FI_22 : 0
    ];
    _profile.Groups.Geii2_FA_Tp_Effectifs = [
      int.TryParse(txtTpCommunEse23_Fa.Text,out int eseTp_FA_23) ? eseTp_FA_23 : 0,
      int.TryParse(txtTpCommunEse24_Fa.Text,out int eseTp_FA_24) ? eseTp_FA_24 : 0,
      int.TryParse(txtTpCommunEme23_Fa.Text,out int emeTp_FA_23) ? emeTp_FA_23 : 0,
      int.TryParse(txtTpCommunEme24_Fa.Text,out int emeTp_FA_24) ? emeTp_FA_24 : 0,
      int.TryParse(txtTpCommunAii23_Fa.Text,out int aiiTp_FA_23) ? aiiTp_FA_23 : 0,
      int.TryParse(txtTpCommunAii24_Fa.Text,out int aiiTp_FA_24) ? aiiTp_FA_24 : 0
    ];

    _profile.Groups.Geii2_FI_TpSp_Effectifs = [
      int.TryParse(txtTpSpEse21_Fi.Text,out int eseTpSp_FI_21) ? eseTpSp_FI_21 : 0,
      int.TryParse(txtTpSpEse22_Fi.Text,out int eseTpSp_FI_22) ? eseTpSp_FI_22 : 0,
      int.TryParse(txtTpSpEme21_Fi.Text,out int emeTpSp_FI_21) ? emeTpSp_FI_21 : 0,
      int.TryParse(txtTpSpEme22_Fi.Text,out int emeTpSp_FI_22) ? emeTpSp_FI_22 : 0,
      int.TryParse(txtTpSpAii21_Fi.Text,out int aiiTpSp_FI_21) ? aiiTpSp_FI_21 : 0,
      int.TryParse(txtTpSpAii22_Fi.Text,out int aiiTpSp_FI_22) ? aiiTpSp_FI_22 : 0
    ];
    _profile.Groups.Geii2_FA_TpSp_Effectifs = [
      int.TryParse(txtTpSpEse23_Fa.Text,out int eseTpSp_FA_23) ? eseTpSp_FA_23 : 0,
      int.TryParse(txtTpSpEse24_Fa.Text,out int eseTpSp_FA_24) ? eseTpSp_FA_24 : 0,
      int.TryParse(txtTpSpEme23_Fa.Text,out int emeTpSp_FA_23) ? emeTpSp_FA_23 : 0,
      int.TryParse(txtTpSpEme24_Fa.Text,out int emeTpSp_FA_24) ? emeTpSp_FA_24 : 0,
      int.TryParse(txtTpSpAii23_Fa.Text,out int aiiTpSp_FA_23) ? aiiTpSp_FA_23 : 0,
      int.TryParse(txtTpSpAii24_Fa.Text,out int aiiTpSp_FA_24) ? aiiTpSp_FA_24 : 0
    ];

    _profile.Groups.TotalGroupe_Geii2_Td_FI_officiel = [
      int.TryParse(txtTotalTD1_Fi.Text, out int td1_Fi) ? td1_Fi : 0,
      int.TryParse(txtTotalTD2_Fi.Text, out int td2_Fi) ? td2_Fi : 0,
      int.TryParse(txtTotalTD3_Fi.Text, out int td3_Fi) ? td3_Fi : 0,
      ];
    _profile.Groups.TotalGroupe_Geii2_Td_FA_officiel = [
      int.TryParse(txtTotalTD1_FA.Text, out int td1_Fa) ? td1_Fa : 0,
      int.TryParse(txtTotalTD2_FA.Text, out int td2_Fa) ? td2_Fa : 0,
      int.TryParse(txtTotalTD3_FA.Text, out int td3_Fa) ? td3_Fa : 0,
      ];

    _profile.Groups.TotalGroupe_Geii2_Tp_FI_officiel = [
      int.TryParse(txtTP1_officiel_Fi.Text, out int tp1_Fi) ? tp1_Fi : 0,
      int.TryParse(txtTP2_officiel_Fi.Text, out int tp2_Fi) ? tp2_Fi : 0,
      int.TryParse(txtTP3_officiel_Fi.Text, out int tp3_Fi) ? tp3_Fi : 0,
      ];
    _profile.Groups.TotalGroupe_Geii2_Tp_FA_officiel = [
      int.TryParse(txtTP1_officiel_Fa.Text, out int tp1_Fa) ? tp1_Fa : 0,
      int.TryParse(txtTP2_officiel_Fa.Text, out int tp2_Fa) ? tp2_Fa : 0,
      int.TryParse(txtTP3_officiel_Fa.Text, out int tp3_Fa) ? tp3_Fa : 0,
      ];

    _profile.Groups.TotalGroupe_Geii2_TpSp_FI_officiel = [
      int.TryParse(txtTP1Sp_officiel_Fi.Text, out int tp1Sp_Fi) ? tp1Sp_Fi : 0,
      int.TryParse(txtTP2Sp_officiel_Fi.Text, out int tp2Sp_Fi) ? tp2Sp_Fi : 0,
      int.TryParse(txtTP3Sp_officiel_Fi.Text, out int tp3Sp_Fi) ? tp3Sp_Fi: 0,
      ];
    _profile.Groups.TotalGroupe_Geii2_TpSp_FA_officiel = [
      int.TryParse(txtTP1Sp_officiel_Fa.Text, out int tp1Sp_Fa) ? tp1Sp_Fa : 0,
      int.TryParse(txtTP2Sp_officiel_Fa.Text, out int tp2Sp_Fa) ? tp2Sp_Fa : 0,
      int.TryParse(txtTP3Sp_officiel_Fa.Text, out int tp3Sp_Fa) ? tp3Sp_Fa: 0,
      ];


    _profile.Groups.Geii2_FI_StudentCount = int.TryParse(txtNbEtudiantsFi.Text,out int Effectif_FI) ? Effectif_FI : 0;
    _profile.Groups.Geii2_FA_StudentCount = int.TryParse(txtNbEtudiantsFa.Text,out int Effectif_FA) ? Effectif_FA : 0;

    _profile.Groups.Geii2_FI_Cm_GroupNames = (string[])NomGroupeCM_FI.Clone();
    _profile.Groups.Geii2_FI_Td_GroupNames = (string[])NomGroupeTD_FI.Clone();
    _profile.Groups.Geii2_FI_Td_officiel_GroupNames = (string[])NomGroupeTD_officiel_FI.Clone();
    _profile.Groups.Geii2_FI_Tp_GroupNames = (string[])NomGroupeTP_FI.Clone();
    _profile.Groups.Geii2_FI_Tp_officiel_GroupNames = (string[])NomGroupeTP_officiel_FI.Clone();
    _profile.Groups.Geii2_FI_TpSp_GroupNames = (string[])NomGroupeTPSp_FI.Clone();
    _profile.Groups.Geii2_FI_TpSp_officiel_GroupNames = (string[])NomGroupeTPSp_officiel_FI.Clone();

    _profile.Groups.Geii2_FA_Td_GroupNames = (string[])NomGroupeTD_FA.Clone();
    _profile.Groups.Geii2_FA_Td_officiel_GroupNames = (string[])NomGroupeTD_officiel_FA.Clone();
    _profile.Groups.Geii2_FA_Tp_GroupNames = (string[])NomGroupeTP_FA.Clone();
    _profile.Groups.Geii2_FA_Tp_officiel_GroupNames = (string[])NomGroupeTP_officiel_FA.Clone();
    _profile.Groups.Geii2_FA_TpSp_GroupNames = (string[])NomGroupeTPSp_FA.Clone();
    _profile.Groups.Geii2_FA_TpSp_officiel_GroupNames = (string[])NomGroupeTPSp_officiel_FA.Clone();
  }

  private static void ActualiserGroupeTp(StackPanel stackPanel,bool visible) {
    Visibility visibility = visible ? Visibility.Visible : Visibility.Collapsed;

    stackPanel.Visibility = visibility;
    stackPanel.IsEnabled = visible;
  }

  private static void ActualiserGroupesTp(
    int effectif,
    StackPanel tp21,StackPanel tp22,
    StackPanel tp21Specialise,StackPanel tp22Specialise) {

    bool premierGroupe = effectif > 0;
    bool secondGroupe = effectif > 14;

    ActualiserGroupeTp(tp21,premierGroupe);
    ActualiserGroupeTp(tp22,secondGroupe);

    ActualiserGroupeTp(tp21Specialise,premierGroupe);
    ActualiserGroupeTp(tp22Specialise,secondGroupe);
  }

  private void ActualiserGroupesTp_FI() {
    (TextBox effectif,StackPanel tp21,StackPanel tp22,StackPanel tp21Sp,StackPanel tp22Sp)[] parcours = [
      (txtNbEtudiantsEseFi,spEse21_Fi,spEse22_Fi,spEse21sp_Fi,spEse22Sp_Fi),
      (txtNbEtudiantsEmeFi,spEme21_Fi,spEme22_Fi,spEme21Sp_Fi,spEme22Sp_Fi),
      (txtNbEtudiantsAiiFi,spAii21_Fi,spAii22_Fi,spAii21Sp_Fi,spAii22Sp_Fi)
    ];

    foreach (var p in parcours) {
      int.TryParse(p.effectif.Text,out int effectif);
      ActualiserGroupesTp(effectif,p.tp21,p.tp22,p.tp21Sp,p.tp22Sp);
    }
  }

  private void ActualiserGroupesTp_FA() {
    (TextBox effectif,StackPanel tp23,StackPanel tp24,StackPanel tp23Sp,StackPanel tp24Sp)[] parcours = [
      (txtNbEtudiantsEseFa,spEse23_Fa,spEse24_Fa,spEse23Sp_Fa,spEse24Sp_Fa),
    (txtNbEtudiantsEmeFa,spEme23_Fa,spEme24_Fa,spEme23Sp_Fa,spEme24Sp_Fa),
    (txtNbEtudiantsAiiFa,spAii23_Fa,spAii24_Fa,spAii23Sp_Fa,spAii24Sp_Fa)
    ];

    foreach (var p in parcours) {
      int.TryParse(p.effectif.Text,out int effectif);
      ActualiserGroupesTp(effectif,p.tp23,p.tp24,p.tp23Sp,p.tp24Sp);
    }
  }

  private void chkFa_Checked(object sender,RoutedEventArgs e) {
    string formation = "GEII 2 FA";
    if (chkFa.IsChecked == true) {
      chkFi.IsChecked = false;
    }
    lblFormationName.Text = $"Répartition prévisionnelle des effectifs en {formation}";
    pnlFa.Visibility = Visibility.Visible;
    pnlFi.Visibility = Visibility.Collapsed;
  }

  private void chkFi_Checked(object sender,RoutedEventArgs e) {
    string formation = "GEII 2 FI";
    if (chkFi.IsChecked == true) {
      chkFa.IsChecked = false;
    }
    lblFormationName.Text = $"Répartition prévisionnelle des effectifs en {formation}";
    pnlFa.Visibility = Visibility.Collapsed;
    pnlFi.Visibility = Visibility.Visible;
  }

  private void EffectifParcours_TextChanged(object sender,TextChangedEventArgs e) {
    if (chkFi.IsChecked == true) ActualiserGroupesTp_FI();
    if (chkFa.IsChecked == true) ActualiserGroupesTp_FA();
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

    // Chaque parcours occupe deux positions fixes dans les tableaux TP.
    int index = Array.IndexOf(
      new[] { "ESE","EME","AII" },
      (chkParcours.Content?.ToString() ?? "").Split('_')[0]);

    if (index < 0) return;
    index *= 2;

    if (int.TryParse(txtNbEtudiants.Text,out int nbEtudiants) && nbEtudiants > 0) {
      nbGroupeTP++;
      nbGroupeTPSp++;

      nomGroupeTP[index] = tbkTpCommun1.Text;
      nomGroupeTPSp[index] = tbkTpSpecialisé1.Text;
      j++;
      k++;

      if (nbEtudiants > 14) {
        nbGroupeTP++;
        nbGroupeTPSp++;

        nomGroupeTP[index + 1] = tbkTpCommun2.Text;
        nomGroupeTPSp[index + 1] = tbkTpSpecialisé2.Text;
        j++;
        k++;
      }
    }
  }

  private void EnregistrerConfigurationEffectifs_FI() {
    int i = 0, j = 0, k = 0;
    nbGroupeCM = 1;
    nbGroupeTD = 0;
    nbGroupeTP = 0;
    nbGroupeTPSp = 0;

    NomGroupeCM_FI = new string[2];
    NomGroupeTD_FI = new string[3];
    NomGroupeTP_FI = new string[6];
    NomGroupeTPSp_FI = new string[6];

    ActualiserGroupesParcours(chkParcoursEseFi,txtNbEtudiantsEseFi,tbkTpCommunEse21_Fi,tbkTpCommunEse22_Fi,
        tbkTpSpEse21_Fi,tbkTpSpEse22_Fi,NomGroupeTD_FI,NomGroupeTP_FI,NomGroupeTPSp_FI,ref i,ref j,ref k);

    ActualiserGroupesParcours(chkParcoursEmeFi,txtNbEtudiantsEmeFi,tbkTpCommunEme21_Fi,tbkTpCommunEme22_Fi,
        tbkTpSpEme21_Fi,tbkTpSpEme22_Fi,NomGroupeTD_FI,NomGroupeTP_FI,NomGroupeTPSp_FI,ref i,ref j,ref k);

    ActualiserGroupesParcours(chkParcoursAiiFi,txtNbEtudiantsAiiFi,tbkTpCommunAii21_Fi,tbkTpCommunAii22_Fi,
        tbkTpSpAii21_Fi,tbkTpSpAii22_Fi,NomGroupeTD_FI,NomGroupeTP_FI,NomGroupeTPSp_FI,ref i,ref j,ref k);

    CalculEffectifGroupeTD_FI();

    nbEtudiant = int.TryParse(txtNbEtudiantsFi.Text,out int effectif) ? effectif : 0;

    FormationGEII.Geii_2_FI.NbEtudiants = nbEtudiant;
    FormationGEII.Geii_2_FI.NbGroupeCM_officiel = nbGroupeCM;
    FormationGEII.Geii_2_FI.NbGroupeTD = nbGroupeTD;
    FormationGEII.Geii_2_FI.NbGroupeTD_officiel = nbGroupeTD_officiel_FI;
    FormationGEII.Geii_2_FI.NbGroupeTP = nbGroupeTP;
    FormationGEII.Geii_2_FI.NbGroupeTP_officiel = nbGroupeTP_officiel_FI;
    FormationGEII.Geii_2_FI.NbGroupeTPSp = nbGroupeTPSp;
    FormationGEII.Geii_2_FI.NbGroupeTPSp_officiel = nbGroupeTPSp_officiel_FI;


    FormationGEII.Geii_2_FI.NomGroupeTD = (string[])NomGroupeTD_FI.Clone();
    FormationGEII.Geii_2_FI.NomGroupeTD_officiel = (string[])NomGroupeTP_FI.Clone();
    FormationGEII.Geii_2_FI.NomGroupeTP = (string[])NomGroupeTP_FI.Clone();
    FormationGEII.Geii_2_FI.NomGroupeTP_officiel = (string[])NomGroupeTP_officiel_FI.Clone();
    FormationGEII.Geii_2_FI.NomGroupeTP_Sp = (string[])NomGroupeTPSp_FI.Clone();
    FormationGEII.Geii_2_FI.NomGroupeTP_Sp_officiel = (string[])NomGroupeTP_officiel_FI.Clone();

    _profile.Groups.Geii2_FI_StudentCount = nbEtudiant;
    _profile.Groups.Geii2_FI_TdGroupCount = nbGroupeTD;
    _profile.Groups.Geii2_FI_Td_officiel_GroupCount = nbGroupeTD_officiel_FI;
    _profile.Groups.Geii2_FI_TpGroupCount = nbGroupeTP;
    _profile.Groups.Geii2_FI_Tp_officiel_GroupCount = nbGroupeTP_officiel_FI;
    _profile.Groups.Geii2_FI_TpSpGroupCount = nbGroupeTPSp;
    _profile.Groups.Geii2_FI_TpSp_officiel_GroupCount = nbGroupeTPSp_officiel_FI;

    ApplyUiToProfile(_profile);
  }

  private void EnregistrerConfigurationEffectifs_FA() {
    int i = 0, j = 0, k = 0;
    nbGroupeTD = 0;
    nbGroupeTP = 0;
    nbGroupeTPSp = 0;

    NomGroupeTD_FA = new string[3];
    NomGroupeTP_FA = new string[6];
    NomGroupeTPSp_FA = new string[6];

    ActualiserGroupesParcours(chkParcoursEseFa,txtNbEtudiantsEseFa,tbkTpCommunEse23_Fa,tbkTpCommunEse24_Fa,
        tbkTpSpEse23_Fa,tbkTpSpEse24_Fa,NomGroupeTD_FA,NomGroupeTP_FA,NomGroupeTPSp_FA,ref i,ref j,ref k);

    ActualiserGroupesParcours(chkParcoursEmeFa,txtNbEtudiantsEmeFa,tbkTpCommunEme23_Fa,tbkTpCommunEme24_Fa,
        tbkTpSpEme23_Fa,tbkTpSpEme24_Fa,NomGroupeTD_FA,NomGroupeTP_FA,NomGroupeTPSp_FA,ref i,ref j,ref k);

    ActualiserGroupesParcours(chkParcoursAiiFa,txtNbEtudiantsAiiFa,tbkTpCommunAii23_Fa,tbkTpCommunAii24_Fa,
        tbkTpSpAii23_Fa,tbkTpSpAii24_Fa,NomGroupeTD_FA,NomGroupeTP_FA,NomGroupeTPSp_FA,ref i,ref j,ref k);

    CalculEffectifGroupeTD_FA();

    nbEtudiant = int.TryParse(txtNbEtudiantsFa.Text,out int effectif) ? effectif : 0;

    FormationGEII.Geii_2_FA.NbEtudiants = nbEtudiant;
    FormationGEII.Geii_2_FA.NbGroupeTD = nbGroupeTD;
    FormationGEII.Geii_2_FA.NbGroupeTD_officiel = nbGroupeTD_officiel_FA;
    FormationGEII.Geii_2_FA.NbGroupeTP = nbGroupeTP;
    FormationGEII.Geii_2_FA.NbGroupeTP_officiel = nbGroupeTP_officiel_FA;
    FormationGEII.Geii_2_FA.NbGroupeTPSp = nbGroupeTPSp;
    FormationGEII.Geii_2_FA.NbGroupeTPSp_officiel = nbGroupeTPSp_officiel_FA;

    FormationGEII.Geii_2_FA.NomGroupeTD = (string[])NomGroupeTD_FA.Clone();
    FormationGEII.Geii_2_FA.NomGroupeTD_officiel = (string[])NomGroupeTD_officiel_FA.Clone();
    FormationGEII.Geii_2_FA.NomGroupeTP = (string[])NomGroupeTP_FA.Clone();
    FormationGEII.Geii_2_FA.NomGroupeTP_officiel = (string[])NomGroupeTP_officiel_FA.Clone();
    FormationGEII.Geii_2_FA.NomGroupeTP_Sp = (string[])NomGroupeTPSp_FA.Clone();
    FormationGEII.Geii_2_FA.NomGroupeTP_Sp_officiel = (string[])NomGroupeTP_officiel_FA.Clone();

    _profile.Groups.Geii2_FA_StudentCount = nbEtudiant;
    _profile.Groups.Geii2_FA_TdGroupCount = nbGroupeTD;
    _profile.Groups.Geii2_FA_Td_officiel_GroupCount = nbGroupeTD_officiel_FA;
    _profile.Groups.Geii2_FA_TpGroupCount = nbGroupeTP;
    _profile.Groups.Geii2_FA_Tp_officiel_GroupCount = nbGroupeTP_officiel_FA;
    _profile.Groups.Geii2_FA_TpSpGroupCount = nbGroupeTPSp;
    _profile.Groups.Geii2_FA_TpSp_officiel_GroupCount = nbGroupeTPSp_officiel_FA;

    ApplyUiToProfile(_profile);
  }

  private string[] CalculEffectifGroupeTD_FI() {
    (TextBox effectif,CheckBox td1,CheckBox td2,CheckBox td3)[] groupes = [
      (txtNbEtudiantsEseFi,ChkESE_Fi,ChkESE_2_Fi,ChkESE_3_Fi),
      (txtNbEtudiantsEmeFi,ChkEME_Fi,ChkEME_2_Fi,ChkEME_3_Fi),
      (txtNbEtudiantsAiiFi,ChkAII_Fi,ChkAII_2_Fi,ChkAII_3_Fi)
    ];

    int[] totaux = new int[3];

    foreach (var groupe in groupes) {
      int.TryParse(groupe.effectif.Text,out int effectif);

      if (groupe.td1.IsChecked == true) totaux[0] += effectif;
      if (groupe.td2.IsChecked == true) totaux[1] += effectif;
      if (groupe.td3.IsChecked == true) totaux[2] += effectif;
    }

    txtTotalTD1_Fi.Text = totaux[0].ToString();
    txtTotalTD2_Fi.Text = totaux[1].ToString();
    txtTotalTD3_Fi.Text = totaux[2].ToString();

    NomGroupeTD_officiel_FI = new string[3];

    if (totaux[0] > 0) NomGroupeTD_officiel_FI[0] = txtTd1_Fi.Text;
    if (totaux[1] > 0) NomGroupeTD_officiel_FI[1] = txtTd2_Fi.Text;
    if (totaux[2] > 0) NomGroupeTD_officiel_FI[2] = txtTd3_Fi.Text;

    nbGroupeTD_officiel_FI = NomGroupeTD_officiel_FI.Count(g => !string.IsNullOrWhiteSpace(g));

    return NomGroupeTD_officiel_FI;
  }

  private string[] CalculEffectifGroupeTD_FA() {
    (TextBox effectif,CheckBox td1,CheckBox td2,CheckBox td3)[] groupes = [
      (txtNbEtudiantsEseFa,ChkESE_FA,ChkESE_2_FA,ChkESE_3_FA),
      (txtNbEtudiantsEmeFa,ChkEME_FA,ChkEME_2_FA,ChkEME_3_FA),
      (txtNbEtudiantsAiiFa,ChkAII_FA,ChkAII_2_FA,ChkAII_3_FA)
    ];

    int[] totaux = new int[3];

    foreach (var groupe in groupes) {
      int.TryParse(groupe.effectif.Text,out int effectif);

      if (groupe.td1.IsChecked == true) totaux[0] += effectif;
      if (groupe.td2.IsChecked == true) totaux[1] += effectif;
      if (groupe.td3.IsChecked == true) totaux[2] += effectif;
    }

    txtTotalTD1_FA.Text = totaux[0].ToString();
    txtTotalTD2_FA.Text = totaux[1].ToString();
    txtTotalTD3_FA.Text = totaux[2].ToString();

    NomGroupeTD_officiel_FA = new string[3];

    if (totaux[0] > 0) NomGroupeTD_officiel_FA[0] = txtTd1_FA.Text;
    if (totaux[1] > 0) NomGroupeTD_officiel_FA[1] = txtTd2_FA.Text;
    if (totaux[2] > 0) NomGroupeTD_officiel_FA[2] = txtTd3_FA.Text;

    nbGroupeTD_officiel_FA = NomGroupeTD_officiel_FA.Count(g => !string.IsNullOrWhiteSpace(g));

    return NomGroupeTD_officiel_FA;
  }

  private string[] CalculEffectifGroupeTP_FA() {
    (TextBox effectif,CheckBox tp1,CheckBox tp2,CheckBox tp3)[] groupes = [
      (txtTpCommunEse23_Fa,chkTP11_Fa,chkTP21_Fa,chkTP31_Fa),
      (txtTpCommunEse24_Fa,chkTP12_Fa,chkTP24_Fa,chkTP32_Fa),
      (txtTpCommunEme23_Fa,chkTP13_Fa,chkTP23_Fa,chkTP33_Fa),
      (txtTpCommunEme24_Fa,chkTP14_Fa,chkTP24_Fa,chkTP34_Fa),
      (txtTpCommunAii23_Fa,chkTP15_Fa,chkTP25_Fa,chkTP35_Fa),
      (txtTpCommunAii24_Fa,chkTP16_Fa,chkTP26_Fa,chkTP36_Fa)
    ];

    int[] totaux = new int[3];

    foreach (var groupe in groupes) {
      int.TryParse(groupe.effectif.Text,out int effectif);

      if (groupe.tp1.IsChecked == true) totaux[0] += effectif;
      if (groupe.tp2.IsChecked == true) totaux[1] += effectif;
      if (groupe.tp3.IsChecked == true) totaux[2] += effectif;
    }

    txtTP1_officiel_Fa.Text = totaux[0].ToString();
    txtTP2_officiel_Fa.Text = totaux[1].ToString();
    txtTP3_officiel_Fa.Text = totaux[2].ToString();

    NomGroupeTP_officiel_FA = new string[3];

    for (int i = 0;i < totaux.Length;i++)
      if (totaux[i] > 0)
        NomGroupeTP_officiel_FA[i] = $"TP{i + 1}";

    nbGroupeTP_officiel_FA = NomGroupeTP_officiel_FA.Count(g => !string.IsNullOrWhiteSpace(g));

    return NomGroupeTP_officiel_FA;
  }

  private string[] CalculEffectifGroupeTP_FI() {
    (TextBox effectif,CheckBox tp1,CheckBox tp2,CheckBox tp3)[] groupes = [
      (txtTpCommunEse21_Fi,chkTP11_Fi,chkTP21_Fi,chkTP31_Fi),
      (txtTpCommunEse22_Fi,chkTP12_Fi,chkTP22_Fi,chkTP32_Fi),
      (txtTpCommunEme21_Fi,chkTP13_Fi,chkTP23_Fi,chkTP33_Fi),
      (txtTpCommunEme22_Fi,chkTP14_Fi,chkTP24_Fi,chkTP34_Fi),
      (txtTpCommunAii21_Fi,chkTP15_Fi,chkTP25_Fi,chkTP35_Fi),
      (txtTpCommunAii22_Fi,chkTP16_Fi,chkTP26_Fi,chkTP36_Fi)
    ];

    int[] totaux = new int[3];

    foreach (var groupe in groupes) {
      int.TryParse(groupe.effectif.Text,out int effectif);

      if (groupe.tp1.IsChecked == true) totaux[0] += effectif;
      if (groupe.tp2.IsChecked == true) totaux[1] += effectif;
      if (groupe.tp3.IsChecked == true) totaux[2] += effectif;
    }

    txtTP1_officiel_Fi.Text = totaux[0].ToString();
    txtTP2_officiel_Fi.Text = totaux[1].ToString();
    txtTP3_officiel_Fi.Text = totaux[2].ToString();

    NomGroupeTP_officiel_FI = new string[3];

    for (int i = 0;i < totaux.Length;i++)
      if (totaux[i] > 0)
        NomGroupeTP_officiel_FI[i] = $"TP{i + 1}";

    nbGroupeTP_officiel_FI = NomGroupeTP_officiel_FI.Count(g => !string.IsNullOrWhiteSpace(g));

    return NomGroupeTP_officiel_FI;
  }

  private string[] CalculEffectifGroupeTP_Sp_FA() {
    (TextBox effectif,CheckBox tp1,CheckBox tp2,CheckBox tp3)[] groupes = [
      (txtTpSpEse23_Fa,chkTP11Sp_Fa,chkTP21Sp_Fa,chkTP31Sp_Fa),
      (txtTpSpEse24_Fa,chkTP12Sp_Fa,chkTP22Sp_Fa,chkTP32Sp_Fa),
      (txtTpSpEme23_Fa,chkTP13Sp_Fa,chkTP23Sp_Fa,chkTP33Sp_Fa),
      (txtTpSpEme24_Fa,chkTP14Sp_Fa,chkTP24Sp_Fa,chkTP34Sp_Fa),
      (txtTpSpAii23_Fa,chkTP15Sp_Fa,chkTP25Sp_Fa,chkTP35Sp_Fa),
      (txtTpSpAii24_Fa,chkTP16Sp_Fa,chkTP26Sp_Fa,chkTP36Sp_Fa)
    ];

    int[] totaux = new int[3];

    foreach (var groupe in groupes) {
      int.TryParse(groupe.effectif.Text,out int effectif);

      if (groupe.tp1.IsChecked == true) totaux[0] += effectif;
      if (groupe.tp2.IsChecked == true) totaux[1] += effectif;
      if (groupe.tp3.IsChecked == true) totaux[2] += effectif;
    }

    txtTP1Sp_officiel_Fa.Text = totaux[0].ToString();
    txtTP2Sp_officiel_Fa.Text = totaux[1].ToString();
    txtTP3Sp_officiel_Fa.Text = totaux[2].ToString();

    NomGroupeTPSp_officiel_FA = new string[3];

    for (int i = 0;i < totaux.Length;i++)
      if (totaux[i] > 0)
        NomGroupeTPSp_officiel_FA[i] = $"TP{i + 1}";

    nbGroupeTPSp_officiel_FA = NomGroupeTPSp_officiel_FA.Count(g => !string.IsNullOrWhiteSpace(g));

    return NomGroupeTPSp_officiel_FA;
  }

  private string[] CalculEffectifGroupeTP_Sp_FI() {
    (TextBox effectif,CheckBox tp1,CheckBox tp2,CheckBox tp3)[] groupes = [
      (txtTpSpEse21_Fi,chkTP11Sp_Fi,chkTP21Sp_Fi,chkTP31Sp_Fi),
      (txtTpSpEse22_Fi,chkTP12Sp_Fi,chkTP22Sp_Fi,chkTP32Sp_Fi),
      (txtTpSpEme21_Fi,chkTP13Sp_Fi,chkTP23Sp_Fi,chkTP33Sp_Fi),
      (txtTpSpEme22_Fi,chkTP14Sp_Fi,chkTP24Sp_Fi,chkTP34Sp_Fi),
      (txtTpSpAii21_Fi,chkTP15Sp_Fi,chkTP25Sp_Fi,chkTP35Sp_Fi),
      (txtTpSpAii22_Fi,chkTP16Sp_Fi,chkTP26Sp_Fi,chkTP36Sp_Fi)
    ];

    int[] totaux = new int[3];

    foreach (var groupe in groupes) {
      int.TryParse(groupe.effectif.Text,out int effectif);

      if (groupe.tp1.IsChecked == true) totaux[0] += effectif;
      if (groupe.tp2.IsChecked == true) totaux[1] += effectif;
      if (groupe.tp3.IsChecked == true) totaux[2] += effectif;
    }

    txtTP1Sp_officiel_Fi.Text = totaux[0].ToString();
    txtTP2Sp_officiel_Fi.Text = totaux[1].ToString();
    txtTP3Sp_officiel_Fi.Text = totaux[2].ToString();

    NomGroupeTPSp_officiel_FI = new string[3];

    for (int i = 0;i < totaux.Length;i++)
      if (totaux[i] > 0)
        NomGroupeTPSp_officiel_FI[i] = $"TP{i + 1}";

    nbGroupeTPSp_officiel_FI = NomGroupeTPSp_officiel_FI.Count(g => !string.IsNullOrWhiteSpace(g));

    return NomGroupeTPSp_officiel_FI;
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

  private void GroupeTP_FA_CheckedChanged(object sender,RoutedEventArgs e) {
    NomGroupeTP_officiel_FA = CalculEffectifGroupeTP_FA();
  }

  private void GroupeTP_Sp_FA_CheckedChanged(object sender,RoutedEventArgs e) {
    NomGroupeTPSp_officiel_FA = CalculEffectifGroupeTP_Sp_FA();
  }

  private void GroupeTP_FI_CheckedChanged(object sender,RoutedEventArgs e) {
    NomGroupeTP_officiel_FI = CalculEffectifGroupeTP_FI();
  }

  private void GroupeTP_Sp_FI_CheckedChanged(object sender,RoutedEventArgs e) {
    NomGroupeTPSp_officiel_FI = CalculEffectifGroupeTP_Sp_FI();
  }
}
