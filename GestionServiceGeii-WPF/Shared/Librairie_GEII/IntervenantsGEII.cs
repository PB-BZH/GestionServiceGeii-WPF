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
║  Nom de fichier : IntervenantsGEII.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Windows.Controls;

namespace GestionServiceGeii.Shared.Librairie_GEII {
  public class Intervenants {
    private ComboBox? intervenant;

    public ComboBox? Intervenant {
      get => intervenant;
      set => intervenant = value ?? throw new ArgumentNullException(nameof(value));
    }
  }
  public class IntervenantsGEII {
    private static ComboBox intervenant = new();
    public static ComboBox Intervenant {
      get => intervenant;
      set => intervenant = value ?? throw new ArgumentNullException(nameof(value));
    }
  }
}
