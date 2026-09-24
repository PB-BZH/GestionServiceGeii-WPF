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
║  Le 24/9/2026 - 00:11
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
║  Nom de fichier : ServiceGeii.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using DataColumn = System.Data.DataColumn;

namespace GestionServiceGeii.Shared.Librairie_GEII {
  public class ServiceGeii {
    #region classe ServiceGeii
    //------------------------

    #region Champs de la classe ServiceGEII
    //-------------------------------------
    private static FormationGEII geii_1 = new();
    private static FormationGEII geii_2 = new();
    private static FormationGEII lP_Sari = new();
    private static FormationGEII lP_MEE = new();

    private static DataColumn nomTitulaires = new();
    private static DataColumn nomVacataires = new();
    private static DataColumn nomNb_Groupe = new();
    private static DataColumn nom_Durée = new();
    //----------------------------------------
    #endregion Champs de la classe ServiceGEII

    #region propriété de ServiceGeii
    //------------------------------
    public static FormationGEII Geii_1 {
      get => geii_1;
      protected set => geii_1 = value ?? geii_1;
    }

    public static FormationGEII Geii_2 {
      get => geii_2;
      protected set => geii_2 = value ?? geii_2;
    }

    public static FormationGEII LP_Sari {
      get => lP_Sari;
      protected set => lP_Sari = value ?? lP_Sari;
    }

    public static FormationGEII LP_MEE {
      get => lP_MEE;
      protected set => lP_MEE = value ?? lP_MEE;
    }

    public static DataColumn NomTitulaires {
      get => nomTitulaires;
      protected set => nomTitulaires = value ?? nomTitulaires;
    }

    public static DataColumn NomVacataires {
      get => nomVacataires;
      protected set => nomVacataires = value ?? nomVacataires;
    }
    //---------------------------------
    #endregion propriété de ServiceGeii

    #region Methode de ServiceGeii
    //----------------------------
    public ServiceGeii() { }
    //-------------------------------
    #endregion Methode de ServiceGeii

    #region Class FicheMatière
    //------------------------
    public class FicheMatière {
      #region champs de FicheMatière
      //----------------------------
      private string id = string.Empty;
      internal string noms = string.Empty;
      internal string semestre = string.Empty;
      internal string formation = string.Empty;
      internal string uE = string.Empty;
      internal string parcours = string.Empty;
      internal string libellé_PPN = string.Empty;
      internal string module = string.Empty;
      internal string libellé = string.Empty;
      internal string commentaires = string.Empty;
      internal string cours = string.Empty;
      internal string nombre = string.Empty;
      internal string groupe = string.Empty;
      internal string durée = string.Empty;
      internal string salle = string.Empty;
      internal string pPN = string.Empty;
      internal string apogée = string.Empty;
      internal string total_type = string.Empty;
      //-------------------------------
      #endregion champs de FicheMatière

      #region Propriété de la classe FicheMatière
      //-----------------------------------------
      public string Noms {
        get => noms;
        protected set => noms = value ?? string.Empty;
      }

      public string Semestre {
        get => semestre;
        protected set => semestre = value ?? string.Empty;
      }

      public string Formation {
        get => formation;
        protected set => formation = value ?? string.Empty;
      }

      public string UE {
        get => uE;
        protected set => uE = value ?? string.Empty;
      }

      public string Parcours {
        get => parcours;
        protected set => parcours = value ?? string.Empty;
      }

      public string Libellé_PPN {
        get => libellé_PPN;
        protected set => libellé_PPN = value ?? string.Empty;
      }

      public string Module {
        get => module;
        protected set => module = value ?? string.Empty;
      }

      public string Libellé {
        get => libellé;
        protected set => libellé = value ?? string.Empty;
      }

      public string Commentaires {
        get => commentaires;
        protected set => commentaires = value ?? string.Empty;
      }

      public string Cours {
        get => cours;
        protected set => cours = value ?? string.Empty;
      }

      public string Nombre {
        get => nombre;
        protected set => nombre = value ?? string.Empty;
      }

      public string Groupe {
        get => groupe;
        protected set => groupe = value ?? string.Empty;
      }

      public string Durée {
        get => durée;
        protected set => durée = value ?? string.Empty;
      }

      public string Total_type {
        get => total_type;
        protected set => total_type = value ?? string.Empty;
      }

      public string Salle {
        get => salle;
        protected set => salle = value ?? string.Empty;
      }

      public string PPN {
        get => pPN;
        protected set => pPN = value ?? string.Empty;
      }

      public string Apogée {
        get => apogée;
        protected set => apogée = value ?? string.Empty;
      }
      internal string Id {
        get => id;
        set => id = value;
      }

      //--------------------------------------------
      #endregion Propriété de la classe FicheMatière

      #region méthode de la classe FicheMatière
      //---------------------------------------
      public FicheMatière() { }

      public FicheMatière(string id,string module,string formation,string cours,string noms) {
        Id = id;
        Module = module;
        Formation = formation;
        Cours = cours;
        Noms = noms;
      }

      public static List<FicheMatière>.Enumerator GetEnumerator() {
        return new List<FicheMatière>.Enumerator();
      }

      public void NettoyageDesDonnées() {
        Id = string.Empty;
        Noms = string.Empty;
        Semestre = string.Empty;
        Formation = string.Empty;
        UE = string.Empty;
        Parcours = string.Empty;
        Libellé_PPN = string.Empty;
        Module = string.Empty;
        Libellé = string.Empty;
        Commentaires = string.Empty;
        Cours = string.Empty;
        Nombre = string.Empty;
        Groupe = string.Empty;
        Durée = string.Empty;
        Total_type = string.Empty;
        Salle = string.Empty;
        PPN = string.Empty;
        Apogée = string.Empty;
      }
      //------------------------------------------
      #endregion méthode de la classe FicheMatière
    }
    //----------------------------
    #endregion classe FicheMatière

    #region classe NomFicheMatière
    //----------------------------
    public class NomFicheMatière {
      #region Champs de la classe NomFicheMatière
      //-----------------------------------------
      private DataColumn nomId = new();
      internal DataColumn nomNoms = new();
      internal DataColumn nomFormation = new();
      internal DataColumn nomUE = new();
      internal DataColumn nomParcours = new();
      internal DataColumn nomLibellé_PPN = new();
      internal DataColumn nomModule = new();
      internal DataColumn nomLibelé = new();
      internal DataColumn nomCommentaires = new();
      internal DataColumn nomCours = new();
      internal DataColumn nomGroupes = new();
      internal DataColumn nomTotal_type = new();
      internal DataColumn nomSalle = new();
      internal DataColumn nomPPN = new();
      internal DataColumn nomApogée = new();

      private static DataColumn titulaires = new();
      private static DataColumn vacataires = new();
      //--------------------------------------------
      #endregion Champs de la classe NomFicheMatière

      #region Propriété de la classe NomFicheMatière
      //--------------------------------------------
      public DataColumn NomNoms {
        get => nomNoms;
        protected set => nomNoms = value ?? nomNoms;
      }

      public DataColumn NomFormation {
        get => nomFormation;
        protected set => nomFormation = value ?? nomFormation;
      }

      public DataColumn NomUE {
        get => nomUE;
        protected set => nomUE = value ?? nomUE;
      }

      public DataColumn NomParcours {
        get => nomParcours;
        protected set => nomParcours = value ?? nomParcours;
      }

      public DataColumn NomLibellé_PPN {
        get => nomLibellé_PPN;
        protected set => nomLibellé_PPN = value ?? nomLibellé_PPN;
      }

      public DataColumn NomModule {
        get => nomModule;
        protected set => nomModule = value ?? nomModule;
      }

      public DataColumn NomLibellé {
        get => nomLibelé;
        protected set => nomLibelé = value ?? nomLibelé;
      }

      public DataColumn NomCommentaires {
        get => nomCommentaires;
        protected set => nomCommentaires = value ?? nomCommentaires;
      }

      public DataColumn NomCours {
        get => nomCours;
        protected set => nomCours = value ?? nomCours;
      }

      public DataColumn NomGroupes {
        get => nomGroupes;
        protected set => nomGroupes = value ?? nomGroupes;
      }

      public DataColumn NomTotal_type {
        get => nomTotal_type;
        protected set => nomTotal_type = value ?? nomTotal_type;
      }

      public DataColumn NomSalle {
        get => nomSalle;
        protected set => nomSalle = value ?? nomSalle;
      }

      public DataColumn NomPPN {
        get => nomPPN;
        protected set => nomPPN = value ?? nomPPN;
      }

      public DataColumn NomApogée {
        get => nomApogée;
        protected set => nomApogée = value ?? nomApogée;
      }

      public static DataColumn NomNb_Groupe {
        get => nomNb_Groupe;
        protected set => nomNb_Groupe = value ?? nomNb_Groupe;
      }

      public static DataColumn NomDurée {
        get => nom_Durée;
        protected set => nom_Durée = value ?? nom_Durée;
      }

      public static DataColumn Titulaires {
        get => titulaires;
        protected set => titulaires = value ?? titulaires;
      }

      public static DataColumn Vacataires {
        get => vacataires;
        protected set => vacataires = value ?? vacataires;
      }
      public DataColumn NomId {
        get => nomId;
        set => nomId = value;
      }

      //-----------------------------------------------
      #endregion Propriété de la classe NomFicheMatière

      #region méthodes de NomFicheMatière
      //---------------------------------
      public NomFicheMatière() { }

      public void NettoyageDesDonnées() {
        NomId = new DataColumn();
        NomNoms = new DataColumn();
        NomFormation = new DataColumn();
        NomUE = new DataColumn();
        NomParcours = new DataColumn();
        NomLibellé_PPN = new DataColumn();
        NomModule = new DataColumn();
        NomLibellé = new DataColumn();
        NomCommentaires = new DataColumn();
        NomCours = new DataColumn();
        NomGroupes = new DataColumn();
        NomTotal_type = new DataColumn();
        NomSalle = new DataColumn();
        NomPPN = new DataColumn();
        NomApogée = new DataColumn();
        NomNb_Groupe = new DataColumn();
        NomDurée = new DataColumn();
        Titulaires = new DataColumn();
        Vacataires = new DataColumn();
      }
      //------------------------------------
      #endregion méthodes de NomFicheMatière
    }
    //-------------------------------
    #endregion classe NomFicheMatière

    //--------------------------
    #endregion Class ServiceGeii
  }
}
