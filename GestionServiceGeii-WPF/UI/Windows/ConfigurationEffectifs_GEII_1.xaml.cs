using System.Windows;
using System.Windows.Documents;
using GestionServiceGeii.Shared.Librairie_GEII;
using PB.BZH.Theme.Theming;
using Brushes = System.Windows.Media.Brushes;

namespace GestionServiceGeii.UI.Windows {
  /// <summary>
  /// Logique d'interaction pour _configurationEffectifs_GEII_1.xaml
  /// </summary>
  public partial class ConfigurationEffectifs_GEII_1: Window {
    private int nbEtudiantMaxParGroupeTP;
    private int nbEtudiant = 0;
    private int nbGroupeTD;
    private int nbGroupeTP;
    public string[] NomGroupeTD = new string[4];
    public string[] NomGroupeTP = new string[8];

    private bool calculGroupeFait;

    public ConfigurationEffectifs_GEII_1() {
      InitializeComponent();
      Title = "Répartition des étudiants";
      StartContentFlowDocument();
      ThemeManager.ApplyTheme(this);
      CouleurParDefaut();
    }

    private struct GroupeTD {
      private int numéro;
      private int a;
      private int b;
      private int nbTotal;
      private string nomGroupeTD;
      private string nomGroupeTP;

      public string NomGroupeTD {
        get => nomGroupeTD;
        internal set => nomGroupeTD = value;
      }
      public string NomGroupeTP {
        get => nomGroupeTP;
        internal set => nomGroupeTP = value;
      }
      public int Numéro {
        get => numéro;
        internal set => numéro = value;
      }
      public int A {
        get => a;
        internal set => a = value;
      }
      public int B {
        get => b;
        internal set => b = value;
      }
      public int NbTotal {
        get => nbTotal;
        set => nbTotal = value;
      }
    }

    private void CouleurParDefaut() {
      nomDeLaFormation.Foreground = Brushes.LightSkyBlue;
    }
    private void BoutonGroupeConfig_Click(object sender,RoutedEventArgs e) {
      try {
        int année = 1;
        nbEtudiant = int.Parse(saisieEffectif.Text);
        nbGroupeTD = 0;
        nbGroupeTP = 0;
        GroupeTD[] groupeTD = new GroupeTD[4];
        int[] groupeTp = new int[8];

        if (choixMax12.IsChecked == true) {
          nbEtudiantMaxParGroupeTP = 12;
        }
        else {
          nbEtudiantMaxParGroupeTP = 14;
        }
        if (nbEtudiant != 0) {
          // calcul des nombres de nombre de TP
          if ((nbEtudiant % nbEtudiantMaxParGroupeTP) < 7)
            nbGroupeTP = nbEtudiant / nbEtudiantMaxParGroupeTP;
          else
            nbGroupeTP = (nbEtudiant / nbEtudiantMaxParGroupeTP) + 1;
          // calcul des nombres de nombre de TD
          if ((nbGroupeTP % 2) != 0)
            nbGroupeTD = (nbGroupeTP / 2) + (nbGroupeTP % 2);
          else
            nbGroupeTD = nbGroupeTP / 2;

          // calcul des nombres d'étudiants par nombre de TP
          int baseCalculNbEtudiants = ((nbEtudiant / nbGroupeTP) / 2) * 2;
          int test = nbEtudiant % nbGroupeTP;
          int nbEtudiantsRestants = nbEtudiant - (baseCalculNbEtudiants * nbGroupeTP);
          int nbRestant = nbEtudiantsRestants;
          for (int i = 0;i < nbGroupeTP;i++) {
            if ((nbRestant % 2) == 0) {
              if (nbRestant != 0) {
                groupeTp[i] = baseCalculNbEtudiants + 2;
                nbRestant -= 2;
              }
              else {
                groupeTp[i] = baseCalculNbEtudiants;
              }
            }
            else {
              groupeTp[i] = baseCalculNbEtudiants + 1;
              nbRestant--;
            }
          }
          for (int i = 0;i < nbGroupeTP;i++) {
            for (int j = i + 1;j < nbGroupeTP;j++) {
              MaxValeur(ref groupeTp[i],ref groupeTp[j]);
            }
          }
          for (int i = 0;i < nbGroupeTD;i++) {
            groupeTD[i].Numéro = 11 + i;
            groupeTD[i].A = groupeTp[i];
            if (i != nbGroupeTP - 1 - i) {
              groupeTD[i].B = groupeTp[nbGroupeTP - 1 - i];
            }
            else {
              groupeTD[i].B = 0;
            }
            groupeTD[i].NbTotal = groupeTD[i].A + groupeTD[i].B;
            // nommage nombre TD
            NomGroupeTD[i] = année.ToString() + (i + 1).ToString();
          }
          // nommage nombre TP
          char suffixe = 'A';
          int numéro = 0;
          for (int i = 0;i < nbGroupeTP;i++) {
            if (i % 2 == 0) {
              suffixe = 'A';
              numéro++;
            }
            NomGroupeTP[i] = année.ToString() + numéro.ToString() + suffixe++;
          }
          Console.WriteLine("Groupe de TD : ");
          for (int i = 0;i < NomGroupeTD.Length;i++)
            if (NomGroupeTD[i] != null)
              Console.WriteLine(NomGroupeTD[i].ToString());
          Console.WriteLine("Groupe de TP : ");
          for (int i = 0;i < NomGroupeTP.Length;i++)
            if (NomGroupeTP[i] != null)
              Console.WriteLine(NomGroupeTP[i].ToString());
          AfficheResultats(groupeTp,groupeTD);
          calculGroupeFait = true;
        }
      }
      catch (Exception ex) {
        MessageBox.Show
          (
            ex.Message +
            "\tEntrer le nombre d'étudiants prévus pour la promotion",
            "Alerte",
            MessageBoxButton.RetryCancel,
            MessageBoxImage.Hand,
            MessageBoxResult.Retry
          );
      }
    }

    private void StartContentFlowDocument() {
      FlowDocument StartDocument = new();

      StartDocument.Blocks.Add(new Paragraph(new Run(
        "EFFECTIF DES GROUPES\n" +
        "1 - Remplir le champs Effectif prévu\n" +
        "2 - Choisir entre 12 ou 14 par groupe de TP\n" +
        "3 - Le bouton[Groupes] affiche le résultat\n" +
        "4 - La case à cocher permet d'éffacer la fenêtre\n" +
        "5 - Modifier éventuellement à partir de l'étape 2\n" +
        "6 - Valider le choix par le bouton[Valider]\n" +
        "\tles valeurs seront affichées en bas de la\n" +
        "\tfenêtre principale de l'application\n" +
        "7 - Fermer la fenêtre par le bouton X")));

      fenetreRésumé.Document = StartDocument;
    }

    private void AfficheResultats(int[] groupeTp,GroupeTD[] groupeTD) {
      fenetreRésumé.Document.Blocks.Clear();
      fenetreRésumé.AppendText(
        "ÉFFECTIFS DES GROUPES EN GEII 1\n" +
        "Étudiants prévus en première année\r" +
        "soit :\r");

      //constitution des groupe de TD
      //-----------------------------
      fenetreRésumé.AppendText(
        "- " + nbGroupeTD + " Groupes de TD\r" +
        "- " + nbGroupeTP + " Groupes de TP\r");

      for (uint i = 0;i < nbGroupeTD;i++) {
        fenetreRésumé.AppendText("\r-> TD"
          + (groupeTD[i].Numéro).ToString()
          + " : " + groupeTD[i].NbTotal + " Étudiants\r");

        char j = 'A';

        if (i != nbGroupeTP - 1 - i) {
          if (groupeTp[i] != 0)
            fenetreRésumé.AppendText("\t\t-> TP"
              + (groupeTD[i].Numéro).ToString() + j++ + " : "
              + groupeTp[i] + " Étudiants\r");
          if (groupeTp[nbGroupeTP - 1 - i] != 0)
            fenetreRésumé.AppendText("\t\t-> TP"
              + (groupeTD[i].Numéro).ToString() + j + " : "
              + (groupeTp[nbGroupeTP - 1 - i]) + " Étudiants\r");
        }
        else {
          if (groupeTp[i] != 0)
            fenetreRésumé.AppendText("\t\t-> TP"
              + (groupeTD[i].Numéro + i).ToString() + j++ + " : "
              + groupeTp[i] + " Étudiants\r");
        }
      }
    }

    private void MaxValeur(ref int value1,ref int value2) {
      int max;
      if (value1 < value2) {
        max = value1;
        value1 = value2;
        value2 = max;
      }
    }

    private void ChoixMax12_CheckedChanged(object sender,RoutedEventArgs e) {
      if (choixMax14.IsChecked == true)
        choixMax14.IsChecked = false;
      else
        choixMax12.IsChecked = true;
    }

    private void ChoixMax14_CheckedChanged(object sender,RoutedEventArgs e) {
      if (choixMax12.IsChecked == true)
        choixMax12.IsChecked = false;
      else
        choixMax14.IsChecked = true;
    }

    private void EffacerFenêtre_CheckedChanged(object sender,RoutedEventArgs e) {
      if (EffacerFenêtre.IsChecked != true)
        return;
      fenetreRésumé.Document.Blocks.Clear();
      EffacerFenêtre.IsChecked = false;
    }


    private void BoutonValiderConfig_Click(object sender,RoutedEventArgs e) {
      if (nbEtudiant != 0) {
        if (calculGroupeFait) {
          FormationGEII.Geii_1_FI.NbEtudiants = nbEtudiant;
          FormationGEII.Geii_1_FI.NbGroupeTD = nbGroupeTD;
          FormationGEII.Geii_1_FI.NbGroupeTP = nbGroupeTP;

          DialogResult = true;
        }
        else
          MessageBox.Show
            (
              "Vous devez effectuer le calcul des nombre",
              "Alerte",
              MessageBoxButton.RetryCancel,
              MessageBoxImage.Hand,
              MessageBoxResult.Retry
            );
      }
      else
        MessageBox.Show
          (
            "Entrer le nombre d'étudiants prévus pour la promotion",
            "Alerte",
            MessageBoxButton.RetryCancel,
            MessageBoxImage.Hand,
            MessageBoxResult.Retry
         );
    }
  }
}
