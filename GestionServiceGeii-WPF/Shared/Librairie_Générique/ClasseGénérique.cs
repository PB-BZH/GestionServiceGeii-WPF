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
║  Nom de fichier : ClasseGénérique.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/

using System.Data;
using System.Windows.Controls;

using Microsoft.Win32;

namespace GestionServiceGeii.Shared.Librairie_Générique;

public class ClasseGénérique {

  public ClasseGénérique() { }

  public static T DialogueExplorateur<T>(string titre) where T : FileDialog, new() {
    return new T {
      Title = titre,
      Filter =
        "Fichier ExcelApp : *.xlsx|*.xlsx|" +
        "Fichier ExcelApp 98 : *.xls|*.xls|" +
        "Fichier ExcelApp avec macro : *.xlsm|*.xlsm|" +
        "Tous les fichiers : *.*|*.*",
      FilterIndex = 1
    };
  }

  public static T RempliComboBox<T>(T liste,DataSet dataSet,string nomListe) where T : ComboBox, new() {
    if (liste == null)
      return liste!;

    liste.ItemsSource = null;
    liste.SelectedIndex = -1;

    if (dataSet == null || !dataSet.Tables.Contains(nomListe))
      return liste;

    DataTable table = dataSet.Tables[nomListe]!;

    string colonneAffichage = TrouverNomColonne(table,liste.Name);

    if (string.IsNullOrWhiteSpace(colonneAffichage))
      return liste;

    DataView vue = new(table);

    liste.DisplayMemberPath = colonneAffichage;
    liste.SelectedValuePath = colonneAffichage;
    liste.ItemsSource = vue;

    if (liste.Items.Count > 0)
      liste.SelectedIndex = 0;

    return liste;
  }

  public static T RempliComboBox<T>(T liste,DataSet dataSet,string nomListe,T critèreDeSelection) where T : ComboBox, new() {

    if (liste == null)
      return liste!;

    liste.ItemsSource = null;
    liste.SelectedIndex = -1;

    if (dataSet == null || !dataSet.Tables.Contains(nomListe))
      return liste;

    DataTable table = dataSet.Tables[nomListe]!;

    string colonneAffichage = TrouverNomColonne(table,liste.Name);
    string colonneCritère = TrouverNomColonne(table,critèreDeSelection.Name);

    if (string.IsNullOrWhiteSpace(colonneAffichage))
      return liste;

    if (string.IsNullOrWhiteSpace(colonneCritère))
      return RempliComboBox(liste,dataSet,nomListe);

    // Valeur réellement sélectionnée dans le ComboBox Semestre
    string valeurCritère = critèreDeSelection.SelectedValue?.ToString()?.Trim() ?? "";

    if (string.IsNullOrWhiteSpace(valeurCritère))
      return liste;

    string valeurFiltre = valeurCritère.Replace("'","''");

    DataView vue = new(table) {
      RowFilter = $"[{colonneCritère}] = '{valeurFiltre}'"
    };

    liste.DisplayMemberPath = colonneAffichage;
    liste.SelectedValuePath = colonneAffichage;
    liste.ItemsSource = vue;

    if (liste.Items.Count > 0)
      liste.SelectedIndex = 0;

    return liste;
  }

  public static T RempliComboBox<T>(T liste,DataSet dataSet,string nomListe,params ComboBox[] criteres) where T : ComboBox, new() {
    if (liste == null)
      return liste!;

    liste.ItemsSource = null;
    liste.SelectedIndex = -1;

    if (dataSet == null || !dataSet.Tables.Contains(nomListe))
      return liste;

    DataTable table = dataSet.Tables[nomListe]!;
    string colonneAffichage = TrouverNomColonne(table,liste.Name);

    if (string.IsNullOrWhiteSpace(colonneAffichage))
      return liste;

    List<string> filtres = new();

    foreach (ComboBox critere in criteres) {
      if (critere == null)
        continue;

      string colonneCritere = TrouverNomColonne(table,critere.Name);
      string valeurCritere = critere.SelectedValue?.ToString()?.Trim() ?? string.Empty;

      if (string.IsNullOrWhiteSpace(colonneCritere) || string.IsNullOrWhiteSpace(valeurCritere))
        continue;

      string valeurFiltre = valeurCritere.Replace("'","''");
      filtres.Add($"[{colonneCritere}] = '{valeurFiltre}'");
    }

    DataView vue = new(table);

    if (filtres.Count > 0)
      vue.RowFilter = string.Join(" AND ",filtres);

    liste.DisplayMemberPath = colonneAffichage;
    liste.SelectedValuePath = colonneAffichage;
    liste.ItemsSource = vue;

    if (liste.Items.Count > 0)
      liste.SelectedIndex = 0;

    return liste;
  }


  private static string TrouverNomColonne(DataTable table,string nomRecherche) {
    foreach (DataColumn colonne in table.Columns) {
      if (string.Equals(colonne.ColumnName,nomRecherche,StringComparison.OrdinalIgnoreCase))
        return colonne.ColumnName;
    }

    return string.Empty;
  }
}
