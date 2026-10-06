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
    internal static Promotion Geii_1_FI { get; set; } = new();
    internal static Promotion Geii_2_FI { get; set; } = new();
    internal static Promotion Geii_1_FA { get; set; } = new();
    internal static Promotion Geii_2_FA { get; set; } = new();
  }
}
