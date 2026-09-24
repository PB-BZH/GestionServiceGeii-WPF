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
║  Le 24/9/2026 - 00:10
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
║  Nom de fichier : ClasseExcel.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.IO;
using System.Windows;
using System.Windows.Controls;
using GestionServiceGeii.Core.Enums;
using GestionServiceGeii.Core.Services;
using GestionServiceGeii.Shared.Librairie_Générique;
using Microsoft.Win32;


namespace GestionServiceGeii.Shared.Librairie_Fichier {
  class ClasseExcel: ClasseFichier {
    #region Champs de la classe
    //-------------------------
    private readonly ClasseGénérique classeGénérique = new();
    string excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='ExcelApp 8.0;HDR={1}'";
    string excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties ='ExcelApp 8.0;HDR={1}'";
    //----------------------------
    #endregion champs de la classe

    #region Propriétés de la classe
    //-----------------------------

    internal string ConStr {
      get {
        if (FichierExcelOuvert) {
          conStr = ChaineDeConnexion(ExtensionFichier,CheminFichier);
        }
        return conStr;
      }
      set => conStr = value;
    }

    private Boolean FichierExcelOuvert {
      get {
        //Boolean testFichierExcel;
        fichierOuvert = ExtensionFichier switch {
          //Excel1 97-03
          ".xls" or ".xlsx" or ".xlms" => true,
          _ => false,
        };
        return fichierOuvert;
      }
    }

    private string Excel03ConString {
      get => excel03ConString;
      set => excel03ConString = value;
    }

    private string Excel07ConString {
      get => excel07ConString;
      set => excel07ConString = value;
    }

    //--------------------------------
    #endregion Propriétés de la classe

    #region Méthodes de la classe
    //---------------------------
    public ClasseExcel() { }

    protected string ChaineDeConnexion(string ExtensionFichier,string CheminFichier) {
      conStr = ExtensionFichier switch {
        //ExcelApp 97-03
        ".xls" => string.Format(Excel03ConString,CheminFichier,"YES"),
        //ExcelApp 07
        ".xlsx" => string.Format(Excel07ConString,CheminFichier,"YES"),
        //ExcelApp 07
        ".xlms" => string.Format(Excel07ConString,CheminFichier,"YES"),
        _ => "non défini",
      };
      return conStr;
    }

    #region Test de l'existence d'un fichier ouvert
    //---------------------------------------------
    /// <summary>
    /// Test de l'existence d'un fichier ouvert
    /// </summary>
    private bool TestFichierOuvert() {
      return fichierOuvert;
    }

    //-----------------------------------------
    #endregion Test de l'existance d'un fichier

    protected internal static void ChangeFeuilleVisible(ComboBox feuillesClasseur) {
      try {
        if (feuillesClasseur.SelectedItem == null)
          return;

        string feuilleSélectionnée = feuillesClasseur.SelectedItem.ToString() ?? string.Empty;

        // Ici devra se trouver l'action réelle
        // permettant de changer la feuille visible.
      }
      catch (Exception ex) {
        MessageBox.Show(
            ex.Message,
            "Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
      }
    }

    #region Ouvrir le fichier Excel
    //-----------------------------
    /// <summary>
    /// Ouverture du fichier ExcelApp
    /// </summary>

    internal bool ChargerDepuisChemin(string cheminFichier) {
      FichierOuvert = false;

      try {
        if (string.IsNullOrWhiteSpace(cheminFichier))
          return false;

        if (!File.Exists(cheminFichier))
          return false;

        CheminFichier = cheminFichier;
        RépertoireFichier = Path.GetDirectoryName(CheminFichier)!;
        NomFichier = Path.GetFileNameWithoutExtension(CheminFichier);
        ExtensionFichier = Path.GetExtension(CheminFichier);

        ConStr = ChaineDeConnexion(ExtensionFichier,CheminFichier);

        FichierOuvert = true;
        return true;
      }
      catch (Exception ex) {
        CheminFichier = string.Empty;
        RépertoireFichier = string.Empty;
        NomFichier = string.Empty;
        ExtensionFichier = string.Empty;
        ConStr = string.Empty;
        FichierOuvert = false;

        MessageBox.Show(ex.Message,"Erreur",MessageBoxButton.OK,MessageBoxImage.Error);

        return false;
      }
    }

    internal void RefreshConnectionString() {
      ConStr = ChaineDeConnexion(
          ExtensionFichier,
          CheminFichier
      );
    }

    public void OuvrirFichierExcel(WorkbookRole role,bool visibilitéFichier = false) {
      OuvrirFichierExcel(
          WorkbookDialogHelper.GetWorkbookDialogTitle(role),
          visibilitéFichier
      );
    }

    public void OuvrirFichierExcel(
        string titreFenetre,
        bool visibilitéFichier = false
    ) {
      FichierOuvert =
          OuvrirFichier(
              titreFenetre
          );

      if (!FichierOuvert)
        return;

      RefreshConnectionString();
    }

    //-----------------------------------
    #endregion Ouverture du Fichier Excel

    #region Fermer et enregistrer Fichier et classeur
    //------------------------------------------------
    /// <summary>
    /// Ouverture du Fichier ExcelApp
    /// </summary>

    protected void FermerFichierExcel() {
      FichierOuvert =
          false;
    }

    public void FermerClasseurExcel() {
      if (!FichierOuvert)
        return;

      try {
        EnregistrerFichier();
        FermerFichierExcel();
      }
      catch (Exception ex) {
        MessageBox.Show(
            ex.Message,
            "Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error
        );
      }
    }

    #region Obtenir la permission de fermeture du fichier
    //------------------------------------------------------
    public string permissionFermeture = string.Empty;

    //------------------------------------------------------
    #endregion Obtenir la permission de fermeture du fichier

    internal bool EstPrêtPourLectureOleDb() {
      if (!FichierOuvert)
        return false;

      if (string.IsNullOrWhiteSpace(CheminFichier))
        return false;

      if (!File.Exists(CheminFichier))
        return false;

      if (string.IsNullOrWhiteSpace(ExtensionFichier))
        return false;

      if (string.IsNullOrWhiteSpace(ConStr))
        return false;

      if (ConStr == "non défini")
        return false;

      return true;
    }

    public string EnregistrerFichier(bool demande,bool afficherDialogue) {
      if (!TestFichierOuvert())
        return NomFichier;

      if (!demande) {
        if (
            !afficherDialogue &&
            !string.IsNullOrEmpty(NomFichier) &&
            !string.IsNullOrEmpty(RépertoireFichier)
        ) {
          MessageBoxResult réponse =
              MessageBox.Show(
                  "Voulez-vous sauvegarder les modifications ?",
                  "Sauver",
                  MessageBoxButton.YesNoCancel,
                  MessageBoxImage.Information,
                  MessageBoxResult.Yes
              );

          switch (réponse) {
            case MessageBoxResult.Yes:
              permissionFermeture = "Oui";
              break;

            case MessageBoxResult.No:
              permissionFermeture = "Non";
              break;

            case MessageBoxResult.Cancel:
              permissionFermeture = "Annuler";
              break;
          }
        }
        else if (afficherDialogue) {
          SauvegardeFichier =
            DialogueExplorateur<SaveFileDialog>(
              "Enregistrer le fichier Excel"
            );

          if (SauvegardeFichier.ShowDialog() == true) {
            NomFichier =
                Path.GetFileName(
                    SauvegardeFichier.FileName
                );

            RépertoireFichier =
                Path.GetDirectoryName(
                    SauvegardeFichier.FileName
                )!;
          }
        }
      }
      else {
        permissionFermeture = "Oui";
      }

      return NomFichier;
    }

    public new string EnregistrerFichier() {
      NomFichier = EnregistrerFichier(true,false);
      return NomFichier;
    }

    public string EnregistrerFichier(bool demande) {
      NomFichier = EnregistrerFichier(demande,true);
      return NomFichier;
    }

    internal bool EstPretPourLectureEpplus() {
      if (string.IsNullOrWhiteSpace(CheminFichier))
        return false;

      if (!File.Exists(CheminFichier))
        return false;

      return true;
    }
    //------------------------------
    #endregion Méthodes de la classe
  }
    #endregion Méthodes de la classe
}
