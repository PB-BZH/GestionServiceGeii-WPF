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
║  Le 18/10/2026 - 12:03
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
║  Nom de fichier : MutualisationCours.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
namespace GestionServiceGeii.Core.Profiles {

  public sealed class MutualisationCours {
    public string Module { get; set; } = string.Empty;
    public string Formation { get; set; } = string.Empty;
    public string Semestre { get; set; } = string.Empty;
    public string Parcours { get; set; } = string.Empty;
    public bool? EstMutualise { get; set; } = false;

    public Dictionary<string,List<string>> Groupes { get; set; } = new();

    public Dictionary<string,int> Effectifs { get; set; } = new();
    public static readonly string[] OngletsTD = ["TD1","TD2","TD3"];
    public static readonly string[] OngletsTP = ["TP1","TP2","TP3","TP4"];
    public static readonly string[] OngletsCours = ["CM","TD1","TD2","TD3","TP1","TP2","TP3","TP4"];
  }
}