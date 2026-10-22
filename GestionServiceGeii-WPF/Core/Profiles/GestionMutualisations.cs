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
║  Nom de fichier : GestionMutualisations.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
namespace GestionServiceGeii.Core.Profiles {

  public sealed class GestionMutualisations {

    private readonly MutualisationCours _mutualisation;

    public GestionMutualisations(string module,string formation,string semestre) {
      _mutualisation = new MutualisationCours {
        Module = module,
        Formation = formation,
        Semestre = semestre
      };

      InitialiserGroupes();
    }

    public void ChargerGroupes(MutualisationCours? source) {
      _mutualisation.Groupes.Clear();

      if (source == null)
        return;

      foreach (var onglet in source.Groupes)
        _mutualisation.Groupes[onglet.Key] = new List<string>(onglet.Value);
    }

    public MutualisationCours ObtenirMutualisation() {
      return new MutualisationCours {
        Module = _mutualisation.Module,
        Formation = _mutualisation.Formation,
        Semestre = _mutualisation.Semestre,
        Parcours = _mutualisation.Parcours,
        Groupes = _mutualisation.Groupes.ToDictionary(
          groupe => groupe.Key,
          groupe => new List<string>(groupe.Value))
      };
    }

    // Initialisation des onglets disponibles
    private void InitialiserGroupes() {
      string[] onglets = { "CM","TD1","TD2","TD3","TP1","TP2","TP3" };

      foreach (string onglet in onglets) {
        _mutualisation.Groupes[onglet] = new List<string>();
        _mutualisation.Effectifs[onglet] = 0;
      }
    }

    // Ajouter un groupe d'étudiants à un onglet
    public void AjouterGroupe(string onglet,string groupe) {
      if (!_mutualisation.Groupes.TryGetValue(onglet,out List<string>? groupes)) {
        groupes = new List<string>();
        _mutualisation.Groupes[onglet] = groupes;
      }

      if (!groupes.Contains(groupe))
        groupes.Add(groupe);
    }

    // Retirer un groupe d'étudiants d'un onglet
    public void RetirerGroupe(string onglet,string groupe) {
      if (_mutualisation.Groupes.TryGetValue(onglet,out List<string>? groupes))
        groupes.Remove(groupe);
    }

    // Lire les groupes sélectionnés
    public IReadOnlyList<string> ObtenirGroupes(string onglet) {
      if (_mutualisation.Groupes.TryGetValue(onglet,out List<string>? groupes))
        return groupes.AsReadOnly();

      return Array.Empty<string>();
    }

    // Vérifier si un groupe appartient à un onglet
    public bool GroupeSelectionne(string onglet,string groupe) {
      return _mutualisation.Groupes.TryGetValue(onglet,out List<string>? groupes)
             && groupes.Contains(groupe);
    }

    // Accéder à la configuration pour une future sauvegarde JSON
    public MutualisationCours ObtenirConfiguration() {
      return _mutualisation;
    }
  }
}