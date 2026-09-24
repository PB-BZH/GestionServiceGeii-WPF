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
║  Nom de fichier : FormationGEII.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
namespace GestionServiceGeii.Shared.Librairie_GEII {
  public class Promotion {
    private int nbEtudiants;
    private int nbGroupeTD;
    private int nbGroupeTP;
    private string? cours;

    internal string? Cours {
      get => cours;
      set => cours = value ?? cours;
    }

    internal int NbEtudiants {
      get => nbEtudiants;
      set => nbEtudiants = value;
    }
    internal int NbGroupeTD {
      get => nbGroupeTD;
      set => nbGroupeTD = value;
    }
    internal int NbGroupeTP {
      get => nbGroupeTP;
      set => nbGroupeTP = value;
    }
  }

  public class FormationGEII {
    private static Promotion fi_geii_1 = new();
    private static Promotion fi_geii_2 = new();
    private static Promotion apprentis_geii_1 = new();
    private static Promotion apprentis_geii_2 = new();
    private static Promotion alternance = new();
    private static Promotion sagema_geii_1 = new();
    private static Promotion sagema_geii_2 = new();

    internal static Promotion FI_Geii_1 {
      get => fi_geii_1;
      set => fi_geii_1 = value ?? fi_geii_1;
    }

    internal static Promotion FI_Geii_2 {
      get => fi_geii_2;
      set => fi_geii_2 = value ?? fi_geii_2;
    }

    internal static Promotion Apprentis_Geii_1 {
      get => apprentis_geii_1;
      set => apprentis_geii_1 = value ?? apprentis_geii_1;
    }
    internal static Promotion Apprentis_Geii_2 {
      get => apprentis_geii_2;
      set => apprentis_geii_2 = value ?? apprentis_geii_2;
    }
    internal static Promotion Aternance {
      get => alternance;
      set => alternance = value ?? alternance;
    }
    internal static Promotion Sagema_Geii_1 {
      get => sagema_geii_1;
      set => sagema_geii_1 = value ?? sagema_geii_1;
    }
    internal static Promotion Sagema_Geii_2 {
      get => sagema_geii_2;
      set => sagema_geii_2 = value ?? sagema_geii_2;
    }
  }
}
