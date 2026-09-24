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
║  Nom de fichier : Classe_FichiersService_Selection_Excel_Fichier.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
namespace GestionServiceGeii.Shared.Librairie_Fichier {

  public class FichierDeService {
    private static ClasseExcel? fichierDeService;

    internal static ClasseExcel Service {
      get {
        if (fichierDeService == null)
          fichierDeService = new ClasseExcel();

        return fichierDeService;
      }
      set {
        fichierDeService = value;
      }
    }
  }

  public class FichierDeSelection {
    private static ClasseExcel fichierDeSelection = new ClasseExcel();
    internal static ClasseExcel Selection {
      get => fichierDeSelection;
      set => fichierDeSelection = value ?? fichierDeSelection;
    }

    public FichierDeSelection() {
      Selection.ConStr = null!;
    }
  }
  public class FichierExcel {
    private FichierExcel() { }
  }
}
