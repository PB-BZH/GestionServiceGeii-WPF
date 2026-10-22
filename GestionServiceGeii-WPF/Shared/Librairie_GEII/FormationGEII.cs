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
║  Nom de fichier : FormationGEII.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
namespace GestionServiceGeii.Shared.Librairie_GEII {
  public class Promotion {
    internal int NbEtudiants { get; set; }
    internal int NbGroupeCM { get; set; } = 1;
    internal int NbGroupeTD { get; set; }
    internal int NbGroupeCM_officiel { get; set; }
    internal int NbGroupeTD_officiel { get; set; }
    internal int NbGroupeTP { get; set; }
    internal int NbGroupeTP_Mutel { get; set; }
    internal int NbGroupeTP_officiel { get; set; }
    internal int NbGroupeTPSp { get; set; }
    internal int NbGroupeTPSp_Mutuel { get; set; }
    internal int NbGroupeTPSp_officiel { get; set; }
    internal string? Cours { get; set; }
    internal string?[] NomGroupeTD = [];
    internal string?[] NomGroupeTD_officiel = [];
    internal string?[] NomGroupeTP = [];
    internal string?[] NomGroupeTP_officiel = [];
    internal string?[] NomGroupeTP_Sp = [];
    internal string?[] NomGroupeTP_Sp_officiel = [];
  }

  public class FormationGEII {
    internal static Promotion Geii_1_FI { get; set; } = new();
    internal static Promotion Geii_2_FI { get; set; } = new();
    internal static Promotion Geii_1_FA { get; set; } = new();
    internal static Promotion Geii_2_FA { get; set; } = new();
    internal static Promotion Geii_2_Mutuel { get; set; } = new();
  }
}
