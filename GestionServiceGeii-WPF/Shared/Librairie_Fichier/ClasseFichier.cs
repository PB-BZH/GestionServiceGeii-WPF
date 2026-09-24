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
║  Nom de fichier : ClasseFichier.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.IO;
using System.Windows;
using GestionServiceGeii.Core.Enums;
using GestionServiceGeii.Core.Services;
using GestionServiceGeii.Shared.Librairie_Générique;
using Microsoft.Win32;

namespace GestionServiceGeii.Shared.Librairie_Fichier {
  class ClasseFichier: ClasseGénérique {
    #region Champ de la classe
    //------------------------
    /// <summary>
    /// Déclaration des champs de la classe Fichier
    /// </summary>
    protected string cheminFichier = string.Empty;
    protected string répertoireFichier = string.Empty;
    protected string nomFichier = string.Empty;
    protected string extensionFichier = string.Empty;
    protected bool fichierOuvert = false;
    protected string conStr = string.Empty;
    private OpenFileDialog explorateurDeFichier = new();
    private SaveFileDialog sauvegardeFichier = new();
    private ClasseGénérique? ClasseGénérique;
    //---------------------------
    #endregion Champ de la classe

    #region propriétés de la classe
    //-----------------------------
    public string CheminFichier {
      get => cheminFichier;
      protected set => cheminFichier = value ?? cheminFichier;
    }
    public bool FichierOuvert {
      get => fichierOuvert;
      internal set => fichierOuvert = value;
    }
    public string ExtensionFichier {
      get => MethodeOuvrir();
      protected set => extensionFichier = value ?? extensionFichier;
    }

    private string MethodeOuvrir() {
      if (extensionFichier == string.Empty)
        OuvrirFichier(WorkbookRole.Selection);
      return extensionFichier;
    }

    public string RépertoireFichier {
      get {
        if (répertoireFichier == string.Empty)
          OuvrirFichier(WorkbookRole.Selection);
        return répertoireFichier;
      }
      protected set => répertoireFichier = value ?? répertoireFichier;
    }

    public string NomFichier {
      get => nomFichier;
      set => nomFichier = value ?? nomFichier;
    }
    public OpenFileDialog ExplorateurDeFichier {
      get => explorateurDeFichier;
      protected set => explorateurDeFichier = value ?? explorateurDeFichier;
    }
    public SaveFileDialog SauvegardeFichier {
      get => sauvegardeFichier;
      protected set => sauvegardeFichier = value ?? sauvegardeFichier;
    }
    //--------------------------------
    #endregion propriétés de la classe

    #region méthodes de la classe
    //---------------------------
    #region Constructeur, destructeur, Dispose
    //----------------------------------------
    ~ClasseFichier() {
      NettoyageDesVariablesFichier();
      explorateurDeFichier = null!;
      sauvegardeFichier = null!;
    }

    //-------------------------------------------
    #endregion Constructeur, destructeur, Dispose

    #region Ouverture de fichier
    //--------------------------
    /// <summary>
    /// Appel du dialogue ouverture de fichier
    /// </summary>
    internal bool OuvrirFichier(string titreFenetre) {
      ClasseGénérique = new ClasseGénérique();
      FichierOuvert = false;

      try {
        OpenFileDialog openFileDialog =
            DialogueExplorateur<OpenFileDialog>(
                "Ouvrir un fichier Excel"
            );

        if (ExplorateurDeFichier.ShowDialog() != true)
          return false;

        CheminFichier =
            ExplorateurDeFichier.FileName;

        RépertoireFichier =
            Path.GetDirectoryName(CheminFichier)!;

        NomFichier =
            Path.GetFileNameWithoutExtension(CheminFichier);

        ExtensionFichier =
            Path.GetExtension(CheminFichier);

        FichierOuvert = true;
      }
      catch (Exception erreur) {
        MessageBox.Show(
            "Erreur : impossible de lire le fichier sur le disque." +
            Environment.NewLine +
            "Erreur d'origine : " + erreur.Message
        );

        FichierOuvert = false;
      }

      return FichierOuvert;
    }

    internal bool OuvrirFichier(WorkbookRole role) {
      return OuvrirFichier(
          WorkbookDialogHelper.GetWorkbookDialogTitle(role)
      );
    }    //-----------------------------
    #endregion Ouverture de fichier

    #region Enregistrement de fichier
    //-------------------------------
    /// <summary>
    /// Appel du dialogue sauegarde de fichier
    /// </summary>
    /// <param name="titreFenetre"></param>
    /// <returns></returns>
    internal bool EnregistrerFichier(string titreFenetre) {
      ClasseGénérique = new ClasseGénérique();
      FichierOuvert = true;
      try {
        SauvegardeFichier =
          DialogueExplorateur<SaveFileDialog>(
            "Enregistrer le fichier Excel"
          );

        if (SauvegardeFichier.ShowDialog() != false) {
          CheminFichier = SauvegardeFichier.FileName;
          RépertoireFichier = Path.GetDirectoryName(CheminFichier)!;
          NomFichier = Path.GetFileNameWithoutExtension(CheminFichier);
          ExtensionFichier = Path.GetExtension(CheminFichier);
          fichierOuvert = true;
        }
      }
      catch (Exception erreur) {
        MessageBox.Show("Erreur:Impossible d'enregistrer le fichier sur le disque.\nErreur d'origine : " + erreur.Message);
        fichierOuvert = false;
      }
      finally {
        // rien à faire
      }
      return FichierOuvert;
    }
    internal bool EnregistrerFichier() {
      EnregistrerFichier("Sauvegarde du fichier");
      return FichierOuvert;
    }
    //----------------------------------
    #endregion Enregistrement de fichier

    #region Fermeture et nettoyage des fichiers ouverts
    //-------------------------------------------------
    protected void NettoyageDesVariablesFichier() {
      CheminFichier = string.Empty;
      RépertoireFichier = string.Empty;
      NomFichier = string.Empty;
      ExtensionFichier = string.Empty;
      FichierOuvert = false;
      ExplorateurDeFichier = null!;
      SauvegardeFichier = null!;
    }
    internal static void SuprimeToutesLesTachesExcel() {
      System.Diagnostics.Process[] processExcel = System.Diagnostics.Process.GetProcessesByName("ExcelApp");
      foreach (System.Diagnostics.Process p in processExcel) {
        if (!string.IsNullOrEmpty(p.ProcessName)) {
          try {
            p.Kill();
          }
          catch { }
        }
      }
      System.Diagnostics.Process[] processConhost = System.Diagnostics.Process.GetProcessesByName("conhost");
      foreach (System.Diagnostics.Process p in processConhost) {
        if (!string.IsNullOrEmpty(p.ProcessName)) {
          try {
            p.Kill();
          }
          catch { }
        }
      }
    }
    //----------------------------------------------------
    #endregion Fermeture et nettoyage des fichiers ouverts
    //--------------------------------
    #endregion propriétés de la classe

  }
}
