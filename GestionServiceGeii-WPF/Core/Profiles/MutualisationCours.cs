namespace GestionServiceGeii.Core.Profiles {

  public sealed class MutualisationCours {
    public string Module { get; set; } = string.Empty;
    public string Formation { get; set; } = string.Empty;
    public string Semestre { get; set; } = string.Empty;
    public string Parcours { get; set; } = string.Empty;

    public Dictionary<string,List<string>> Groupes { get; set; } = new();

    public Dictionary<string,int> Effectifs { get; set; } = new();
    public static readonly string[] OngletsTD = ["TD1","TD2","TD3"];
    public static readonly string[] OngletsTP = ["TP1","TP2","TP3"];
    public static readonly string[] OngletsCours = ["CM","TD1","TD2","TD3","TP1","TP2","TP3"];
  }
}