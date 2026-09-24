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
║  Nom de fichier : PasserelleBdD_vers_Excel.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Data;
using System.IO;
using System.Windows;
using GestionServiceGeii.Shared.Database;

namespace GestionServiceGeii.Shared.Librairie_Fichier {
  public class PasserelleBdD_vers_Excel {
    #region Champs de la classe
    //-------------------------
    /// <summary>
    /// Champs de la classe
    /// </summary>
    private static ClasseExcel classeFichierExcel = new();
    private static DataSet dataSetExcel = new();
    private static ClasseExcel fichierDeSelection = new();
    //----------------------------
    #endregion Champs de la classe

    #region Propriétés de la classe
    //-----------------------------
    /// <summary>
    /// Propriétés de la classe
    /// </summary>
    internal static ClasseExcel ClasseFichierExcel {
      get => classeFichierExcel;
      set => classeFichierExcel = value;
    }

    public static DataSet DataSetExcel {
      get => dataSetExcel;
      set => dataSetExcel = value;
    }

    internal static ClasseExcel FichierDeSelection {
      get => fichierDeSelection;
      set => fichierDeSelection = value;
    }
    //--------------------------------
    #endregion Propriétés de la classe

    #region Méthodes de la classe
    //---------------------------
    /// <summary>
    /// Sauvegarde de la base de données stockée en mémoire
    /// </summary>
    /// <param name="DataSetExcel"></param>
    public void SauvegardeBaseDeDonnées_Vers_Excel(DataSet DataSetExcel) {
      ClasseExcel FichierClasseExcel = new();

      FichierClasseExcel.FichierOuvert =
          FichierClasseExcel.EnregistrerFichier(
              "Sauvegarde du fichier de service"
          );

      string nouveauChemin =
          FichierClasseExcel.RépertoireFichier +
          "\\" +
          FichierClasseExcel.NomFichier +
          ".xlsx";

      try {
        if (File.Exists(nouveauChemin)) {
          File.Delete(nouveauChemin);
        }

        ClasseEpplus.CreerClasseurDepuisDataSet_Epplus(
            nouveauChemin,
            DataSetExcel
        );
      }
      catch (Exception exc) {
        MessageBox.Show(
            exc.Message,
            "Sauvegarde Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Error
        );
      }
    }

    internal void MiseAJourTable_Vers_Excel(
        ClasseExcel fichierDialogue,
        DataSet dataSetExcel,
        string nomBase
    ) {
      if (fichierDialogue == null)
        return;

      if (dataSetExcel == null)
        return;

      if (!dataSetExcel.Tables.Contains(nomBase))
        return;

      try {
        ClasseEpplus.MiseAJourTableDepuisDataTable_Epplus(
            fichierDialogue.CheminFichier,
            nomBase,
            dataSetExcel.Tables[nomBase]!
        );
      }
      catch (Exception ex) {
        MessageBox.Show(
            ex.Message,
            "Mise à jour table vers Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Error
        );
      }
    }

    //------------------------------
    #endregion Méthodes de la classe
  }
}
