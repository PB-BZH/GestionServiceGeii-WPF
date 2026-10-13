namespace GestionServiceGeii.Core.Profiles {

  public sealed class MutualisationCours {
    public string Module { get; set; } = string.Empty;
    public string Formation { get; set; } = string.Empty;
    public string Semestre { get; set; } = string.Empty;
    public string Parcours { get; set; } = string.Empty;

    public Dictionary<string,List<string>> Groupes { get; set; } = new();

    public Dictionary<string,int> Effectifs { get; set; } = new();
  }
}