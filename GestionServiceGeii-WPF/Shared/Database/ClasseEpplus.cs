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
║  Nom de fichier : ClasseEpplus.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Data;
using System.IO;
using GestionServiceGeii.Shared.Librairie_Fichier;
using OfficeOpenXml;

namespace GestionServiceGeii.Shared.Database {
  public class ClasseEpplus {

    #region Suppression
    // ----------------   

    /// <summary>
    /// Supprime les lignes d'une feuille ExcelApp qui correspondent à une condition donnée.
    /// </summary>
    /// <param name="cheminFichier">Chemin complet du fichier ExcelApp.</param>
    /// <param name="nomFeuille">Nom de la feuille à traiter.</param>
    /// <param name="condition">Condition évaluée pour chaque ligne.</param>
    /// <returns>Nombre de lignes supprimées.</returns>
    internal static int SupprimeLigneExcel_Epplus(
        string cheminFichier,
        string nomFeuille,
        Func<ExcelWorksheet,int,bool> condition
    ) {
      int lignesSupprimees =
          0;

      ConfigureEpplusLicense();

      if (!File.Exists(cheminFichier))
        throw new FileNotFoundException(
            $"Fichier introuvable : {cheminFichier}"
        );

      using (ExcelPackage package = new(new FileInfo(cheminFichier))) {
        ExcelWorksheet feuille =
            package.Workbook.Worksheets[nomFeuille] ?? throw new ArgumentException(
              $"Feuille '{nomFeuille}' introuvable."
          );
        if (feuille.Dimension == null)
          return 0;

        for (int ligne = feuille.Dimension.End.Row;ligne >= 2;ligne--) {
          if (condition(feuille,ligne)) {
            feuille.DeleteRow(ligne);
            lignesSupprimees++;
          }
        }

        if (lignesSupprimees > 0)
          package.Save();
      }

      return lignesSupprimees;
    }

    #endregion

    #region Mise à jour ligne / cellule
    // --------------------------------

    internal static List<string> RechercheNomsPlagesExcel_Epplus(
    string cheminFichier
) {
      List<string> nomsPlages = [];

      ConfigureEpplusLicense();

      if (!File.Exists(cheminFichier))
        throw new FileNotFoundException(
            $"Fichier introuvable : {cheminFichier}"
        );

      using (ExcelPackage package = new(new FileInfo(cheminFichier))) {
        foreach (ExcelNamedRange plage in package.Workbook.Names) {
          if (plage == null)
            continue;

          if (string.IsNullOrWhiteSpace(plage.Name))
            continue;

          if (plage.Name.StartsWith("_xlnm",StringComparison.OrdinalIgnoreCase))
            continue;

          if (plage.Worksheet == null)
            continue;

          if (plage.Start == null || plage.End == null)
            continue;

          nomsPlages.Add(
              plage.Name
          );
        }
      }

      return nomsPlages;
    }

    internal static DataSet LireToutesPlagesNommesExcel_Epplus(
    string cheminFichier
) {
      ConfigureEpplusLicense();
      DataSet dataSet = new();

      List<string> nomsPlages =
          RechercheNomsPlagesExcel_Epplus(
              cheminFichier
          );

      foreach (string nomPlage in nomsPlages) {
        DataTable dataTable =
            LirePlageNommeeExcel_Epplus(
                cheminFichier,
                nomPlage
            );

        if (dataSet.Tables.Contains(dataTable.TableName))
          continue;

        dataSet.Tables.Add(
            dataTable
        );
      }

      return dataSet;
    }


    /// <summary>
    /// Met à jour dans le fichier ExcelApp de service la durée et le total type
    /// d'une ligne Table_complete identifiée par son ID.
    /// </summary>
    /// <param name="fichierDialogue">Fichier ExcelApp de service à modifier.</param>
    /// <param name="nomBase">Nom de la feuille ExcelApp concernée.</param>
    /// <param name="id">ID de la ligne à modifier.</param>
    /// <param name="duree">Nouvelle durée.</param>
    /// <param name="totalType">Nouveau total type.</param>
    /// <returns>Nombre de lignes modifiées.</returns>
    internal static int MiseAJourDureeEtTotalTypeParId_Epplus(
        ClasseExcel fichierDialogue,
        string nomBase,
        string id,
        int duree,
        int totalType
    ) {
      if (string.IsNullOrWhiteSpace(id))
        return 0;

      try {
        return ClasseEpplus.MiseAJourCelluleExcel_Epplus(
            cheminFichier: fichierDialogue.CheminFichier,
            nomFeuille: nomBase,
            condition: (feuille,ligneCible) =>
                ClasseEpplus.GetStringByColumnName_Epplus(
                    feuille,
                    ligneCible,
                    ExcelSchemaNames.Columns.ID
                ) == id,
            actionMiseAJour: (feuille,ligneCible) => {
              ClasseEpplus.SetValueByColumnName_Epplus(
                  feuille,
                  ligneCible,
                  ExcelSchemaNames.Columns.Duree,
                  duree
              );

              ClasseEpplus.SetValueByColumnName_Epplus(
                  feuille,
                  ligneCible,
                  ExcelSchemaNames.Columns.TotalType,
                  totalType
              );
            }
        );
      }
      catch (Exception ex) {
        //ClasseBaseDeDonnées.ShowOleDbError(
        //    ex,
        //    "MiseAJourDureeEtTotalTypeParId_Epplus",
        //    "Mise à jour EPPlus par ID : " + id
        //);

        return 0;
      }
    }


    /// <summary>
    /// Met à jour les cellules d'une feuille ExcelApp qui vérifient la condition fournie en exécutant l'action de mise à jour correspondante.
    /// </summary>
    /// <param name="cheminFichier">Chemin complet du fichier ExcelApp.</param>
    /// <param name="nomFeuille">Nom de la feuille à traiter.</param>
    /// <param name="condition">Condition évaluée pour chaque ligne.</param>
    /// <param name="actionMiseAJour">Action exécutée sur chaque ligne correspondante.</param>
    /// <returns>Nombre de lignes modifiées.</returns>
    internal static int MiseAJourCelluleExcel_Epplus(string cheminFichier,string nomFeuille,Func<ExcelWorksheet,int,bool> condition,Action<ExcelWorksheet,int> actionMiseAJour,bool premierResultatSeulement = false) {
      int lignesModifiees = 0;

      ConfigureEpplusLicense();

      if (!File.Exists(cheminFichier))
        throw new FileNotFoundException(
            $"Fichier introuvable : {cheminFichier}"
        );

      using (ExcelPackage package = new(new FileInfo(cheminFichier))) {
        ExcelWorksheet feuille = package.Workbook.Worksheets[nomFeuille] ?? throw new ArgumentException($"Feuille '{nomFeuille}' introuvable.");
        for (int ligne = 2;ligne <= feuille.Dimension.End.Row;ligne++) {
          if (condition(feuille,ligne)) {
            actionMiseAJour(feuille,ligne);
            lignesModifiees++;
            if (premierResultatSeulement)
              break;
          }
        }
        if (lignesModifiees > 0)
          package.Save();
      }

      return lignesModifiees;
    }

    //----------------
    #endregion

    #region Lecture cellules
    //----------------------

    internal static DataTable LirePlageNommeeExcel_Epplus(
      string cheminFichier,
      string nomPlage
    ) {
      ConfigureEpplusLicense();

      if (!File.Exists(cheminFichier))
        throw new FileNotFoundException(
            $"Fichier introuvable : {cheminFichier}"
        );

      DataTable dataTable = new(nomPlage);

      using (ExcelPackage package = new(new FileInfo(cheminFichier))) {
        ExcelNamedRange plageNommee = null!;

        foreach (ExcelNamedRange nom in package.Workbook.Names) {
          if (nom.Name == nomPlage) {
            plageNommee = nom;
            break;
          }
        }

        if (plageNommee == null)
          throw new ArgumentException("Plage nommée introuvable : " + nomPlage);

        ExcelWorksheet feuille =
            plageNommee.Worksheet;

        int startRow =
            plageNommee.Start.Row;

        int endRow =
            plageNommee.End.Row;

        int startColumn =
            plageNommee.Start.Column;

        int endColumn =
            plageNommee.End.Column;

        for (int colonne = startColumn;colonne <= endColumn;colonne++) {
          string nomColonne =
              feuille.Cells[startRow,colonne].Text.Trim();

          if (string.IsNullOrWhiteSpace(nomColonne))
            nomColonne =
                "Column" + colonne;

          if (dataTable.Columns.Contains(nomColonne))
            nomColonne =
                nomColonne + "_" + colonne;

          dataTable.Columns.Add(
              nomColonne
          );
        }

        for (int ligne = startRow + 1;ligne <= endRow;ligne++) {
          DataRow dataRow =
              dataTable.NewRow();

          bool ligneVide =
              true;

          for (int colonne = startColumn;colonne <= endColumn;colonne++) {
            object valeur =
                feuille.Cells[ligne,colonne].Value;

            if (valeur != null)
              ligneVide =
                  false;

            dataRow[colonne - startColumn] =
                valeur ?? DBNull.Value;
          }

          if (!ligneVide)
            dataTable.Rows.Add(dataRow);
        }
      }

      return dataTable;
    }



    /// <summary>
    /// Trouve l'index numérique d'une colonne à partir de son nom dans la première ligne de la feuille ExcelApp.
    /// </summary>
    /// <param name="feuille">Feuille ExcelApp à analyser.</param>
    /// <param name="nomColonne">Nom de la colonne recherchée.</param>
    /// <returns>Index numérique de la colonne.</returns>
    internal static int FindColumnIndexByName_Epplus(ExcelWorksheet feuille,string nomColonne) {
      ArgumentNullException.ThrowIfNull(feuille);
      ConfigureEpplusLicense();
      if (feuille.Dimension == null)
        throw new InvalidOperationException(
            "La feuille ExcelApp est vide."
        );

      if (string.IsNullOrWhiteSpace(nomColonne))
        throw new ArgumentException(
            "Le nom de colonne est vide.",
            nameof(nomColonne)
        );

      for (int colonne = 1;colonne <= feuille.Dimension.End.Column;colonne++) {
        if (feuille.Cells[1,colonne].Text.Equals(
                nomColonne,
                StringComparison.OrdinalIgnoreCase
            )) {
          return colonne;
        }
      }

      throw new ArgumentException(
          $"Colonne '{nomColonne}' introuvable."
      );
    }

    /// <summary>
    /// Lit la valeur texte d'une cellule à partir du nom de colonne.
    /// </summary>
    /// <param name="feuille">Feuille ExcelApp à lire.</param>
    /// <param name="ligne">Numéro de ligne à lire.</param>
    /// <param name="nomColonne">Nom de la colonne à lire.</param>
    /// <returns>Valeur texte de la cellule.</returns>
    internal static string GetStringByColumnName_Epplus(
        ExcelWorksheet feuille,
        int ligne,
        string nomColonne
    ) {
      ConfigureEpplusLicense();
      int colonne =
          FindColumnIndexByName_Epplus(
              feuille,
              nomColonne
          );

      return feuille.Cells[ligne,colonne].Text;
    }

    /// <summary>
    /// Lit la valeur entière d'une cellule à partir du nom de colonne.
    /// Retourne 0 si la cellule est vide ou si la conversion échoue.
    /// </summary>
    /// <param name="feuille">Feuille ExcelApp à lire.</param>
    /// <param name="ligne">Numéro de ligne à lire.</param>
    /// <param name="nomColonne">Nom de la colonne à lire.</param>
    /// <returns>Valeur entière lue, ou 0 si la conversion échoue.</returns>
    internal static int GetIntByColumnName_Epplus(
        ExcelWorksheet feuille,
        int ligne,
        string nomColonne
    ) {
      ConfigureEpplusLicense();
      int colonne =
          FindColumnIndexByName_Epplus(
              feuille,
              nomColonne
          );


      if (int.TryParse(
              feuille.Cells[ligne,colonne].Text,
              out int valeur
          )) {
        return valeur;
      }

      return 0;
    }

    //----------------------
    #endregion

    #region Ecriture cellules
    //-----------------------

    /// <summary>
    /// Ajoute une nouvelle ligne à la fin d'une feuille ExcelApp et exécute 
    /// une action pour remplir les cellules de cette ligne.
    /// </summary>
    /// <param name="cheminFichier"></param>
    /// <param name="nomFeuille"></param>
    /// <param name="actionAjout"></param>
    /// <returns></returns>
    /// <exception cref="FileNotFoundException"></exception>
    /// <exception cref="ArgumentException"></exception>
    internal static int AjouterLigneExcel_Epplus(string cheminFichier,string nomFeuille,Action<ExcelWorksheet,int> actionAjout) {
      ConfigureEpplusLicense();

      if (!File.Exists(cheminFichier))
        throw new FileNotFoundException(
            $"Fichier introuvable : {cheminFichier}"
        );

      using ExcelPackage package = new(new FileInfo(cheminFichier));
      ExcelWorksheet feuille = package.Workbook.Worksheets[nomFeuille] ?? throw new ArgumentException($"Feuille '{nomFeuille}' introuvable.");
      ExcelNamedRange plageNommee = null!;

      foreach (ExcelNamedRange nom in package.Workbook.Names) {
        if (nom.Name == nomFeuille) {
          plageNommee = nom;
          break;
        }
      }

      if (plageNommee == null)
        throw new ArgumentException("Plage nommée introuvable : " + nomFeuille);

      int startRow = plageNommee.Start.Row;
      int startColumn = plageNommee.Start.Column;
      int endRow = plageNommee.End.Row;
      int endColumn = plageNommee.End.Column;
      int nouvelleLigne = endRow + 1;

      feuille.InsertRow(nouvelleLigne,1,endRow);
      actionAjout(feuille,nouvelleLigne);
      string nouvelleAdresse = ExcelCellBase.GetAddress(startRow,startColumn,endRow + 1,endColumn);
      package.Workbook.Names.Remove(plageNommee.Name);
      package.Workbook.Names.Add(nomFeuille,feuille.Cells[nouvelleAdresse]);
      package.Save();
      return 1;
    }


    /// <summary>
    /// Écrit une valeur dans une cellule à partir du nom de colonne.
    /// </summary>
    /// <param name="feuille">Feuille ExcelApp à modifier.</param>
    /// <param name="ligne">Numéro de ligne à modifier.</param>
    /// <param name="nomColonne">Nom de la colonne à modifier.</param>
    /// <param name="valeur">Valeur à écrire.</param>
    internal static void SetValueByColumnName_Epplus(ExcelWorksheet feuille,int ligne,string nomColonne,object valeur) {
      int colonne = FindColumnIndexByName_Epplus(feuille,nomColonne);
      feuille.Cells[ligne,colonne].Value = valeur;
    }

    /// <summary>
    /// Écrit une valeur entière dans une cellule à partir du nom de colonne, en convertissant la chaîne de caractères fournie.
    /// </summary>
    /// <param name="feuille"></param>
    /// <param name="ligne"></param>
    /// <param name="colonne"></param>
    /// <param name="valeur"></param>
    internal static void SetIntValueByColumnName_Epplus(
      OfficeOpenXml.ExcelWorksheet feuille,
          int ligne,
          string colonne,
          string valeur
      ) {

      if (!int.TryParse(valeur,out int intValue))
        intValue = 0;

      ClasseEpplus.SetValueByColumnName_Epplus(
          feuille,
          ligne,
          colonne,
          intValue
      );
    }

    private static string SanitizeWorksheetName_Epplus(string name) {
      if (string.IsNullOrWhiteSpace(name))
        return "Feuil1";

      string result =
          name;

      char[] invalidChars =
          [
          '\\',
          '/',
          '?',
          '*',
          '[',
          ']',
          ':'
          ];

      foreach (char invalidChar in invalidChars) {
        result =
            result.Replace(
                invalidChar,
                '_'
            );
      }

      if (result.Length > 31)
        result =
            result.Substring(
                0,
                31
            );

      return result;
    }

    internal static void CreerClasseurDepuisDataSet_Epplus(string cheminFichier,DataSet dataSet) {
      ConfigureEpplusLicense();

      if (string.IsNullOrWhiteSpace(cheminFichier))
        throw new ArgumentException(
            "Le chemin du fichier Excel est vide."
        );

      ArgumentNullException.ThrowIfNull(dataSet);

      FileInfo fichier = new(cheminFichier);

      if (fichier.Exists)
        fichier.Delete();

      using ExcelPackage package = new();
      foreach (DataTable table in dataSet.Tables) {
        if (table == null)
          continue;

        if (string.IsNullOrWhiteSpace(table.TableName))
          continue;

        string nomTable =
            table.TableName;

        if (nomTable[^1] == '$')
          continue;

        string nomFeuille =
            SanitizeWorksheetName_Epplus(nomTable);

        ExcelWorksheet feuille =
            package.Workbook.Worksheets.Add(
                nomFeuille
            );

        for (int colonne = 0;colonne < table.Columns.Count;colonne++) {
          feuille.Cells[1,colonne + 1].Value =
              table.Columns[colonne].ColumnName;
        }

        for (int ligne = 0;ligne < table.Rows.Count;ligne++) {
          for (int colonne = 0;colonne < table.Columns.Count;colonne++) {
            object valeur =
                table.Rows[ligne][colonne];

            if (valeur == null || valeur == DBNull.Value) {
              feuille.Cells[ligne + 2,colonne + 1].Value =
                  string.Empty;

              continue;
            }

            feuille.Cells[ligne + 2,colonne + 1].Value =
                valeur;
          }
        }

        int derniereLigne =
            table.Rows.Count + 1;

        int derniereColonne =
            table.Columns.Count;

        if (derniereColonne > 0) {
          ExcelRange plage =
              feuille.Cells[
                  1,
                  1,
                  derniereLigne,
                  derniereColonne
              ];

          package.Workbook.Names.Add(
              nomTable,
              plage
          );

          feuille.Cells[
              1,
              1,
              derniereLigne,
              derniereColonne
          ].AutoFitColumns();
        }
      }

      if (package.Workbook.Worksheets.Count == 0) {
        ExcelWorksheet feuille =
            package.Workbook.Worksheets.Add(
                "Feuil1"
            );

        feuille.Cells[1,1].Value =
            string.Empty;
      }

      package.SaveAs(
          fichier
      );
    }

    internal static void MiseAJourTableDepuisDataTable_Epplus(
    string cheminFichier,
    string nomPlage,
    DataTable table
) {
      ConfigureEpplusLicense();

      if (string.IsNullOrWhiteSpace(cheminFichier))
        throw new ArgumentException(
            "Le chemin du fichier Excel est vide."
        );

      if (!File.Exists(cheminFichier))
        throw new FileNotFoundException(
            "Fichier introuvable : " + cheminFichier
        );

      if (string.IsNullOrWhiteSpace(nomPlage))
        throw new ArgumentException(
            "Le nom de la plage est vide."
        );

      ArgumentNullException.ThrowIfNull(table);

      using ExcelPackage package = new(new FileInfo(cheminFichier));
      ExcelNamedRange plageNommee =
          null!;

      foreach (ExcelNamedRange nom in package.Workbook.Names) {
        if (nom.Name == nomPlage) {
          plageNommee =
              nom;

          break;
        }
      }

      if (plageNommee == null)
        throw new ArgumentException(
            "Plage nommée introuvable : " + nomPlage
        );

      ExcelWorksheet feuille =
          plageNommee.Worksheet;

      int startRow =
          plageNommee.Start.Row;

      int startColumn =
          plageNommee.Start.Column;

      int oldEndRow =
          plageNommee.End.Row;

      int oldEndColumn =
          plageNommee.End.Column;

      int newEndRow =
          startRow + table.Rows.Count;

      int newEndColumn =
          startColumn + table.Columns.Count - 1;

      // Nettoyage de l'ancienne zone.
      feuille.Cells[
          startRow,
          startColumn,
          oldEndRow,
          oldEndColumn
      ].Clear();

      // En-têtes.
      for (int colonne = 0;colonne < table.Columns.Count;colonne++) {
        feuille.Cells[startRow,startColumn + colonne].Value =
            table.Columns[colonne].ColumnName;
      }

      // Données.
      for (int ligne = 0;ligne < table.Rows.Count;ligne++) {
        for (int colonne = 0;colonne < table.Columns.Count;colonne++) {
          object valeur =
              table.Rows[ligne][colonne];

          if (valeur == null || valeur == DBNull.Value) {
            feuille.Cells[startRow + ligne + 1,startColumn + colonne].Value =
                string.Empty;
          }
          else {
            feuille.Cells[startRow + ligne + 1,startColumn + colonne].Value =
                valeur;
          }
        }
      }

      // Redimensionnement de la plage nommée.
      string nouvelleAdresse =
          ExcelCellBase.GetAddress(
              startRow,
              startColumn,
              newEndRow,
              newEndColumn
          );

      package.Workbook.Names.Remove(
          plageNommee.Name
      );

      package.Workbook.Names.Add(
          nomPlage,
          feuille.Cells[nouvelleAdresse]
      );

#if DEBUG
      System.Diagnostics.Debug.WriteLine(
          "MiseAJourTableDepuisDataTable_Epplus" +
          " | Plage=" + nomPlage +
          " | Rows=" + table.Rows.Count +
          " | Columns=" + table.Columns.Count +
          " | Adresse=" + nouvelleAdresse
      );
#endif

      package.Save();
    }


    //----------------------
    #endregion

    #region Identifiants
    //-----------------------

    internal static void ConfigureEpplusLicense() {
      ExcelPackage.License.SetNonCommercialPersonal("PB-BZH Concept");
    }

    /// <summary>
    /// Calcule le prochain ID disponible à partir de la valeur maximale présente
    /// dans une colonne ID du fichier ExcelApp réel.
    /// </summary>
    /// <param name="cheminFichier">Chemin complet du fichier ExcelApp.</param>
    /// <param name="nomFeuille">Nom de la feuille à analyser.</param>
    /// <param name="nomColonneId">Nom de la colonne contenant les ID.</param>
    /// <returns>Prochain ID disponible sous forme de texte.</returns>
    internal static string GetNextIdFromWorksheet_Epplus(
        string cheminFichier,
        string nomFeuille,
        string nomColonneId
    ) {
      ConfigureEpplusLicense();

      if (string.IsNullOrWhiteSpace(cheminFichier))
        throw new ArgumentException(
            "Le chemin du fichier ExcelApp est vide.",
            nameof(cheminFichier)
        );

      if (!File.Exists(cheminFichier))
        throw new FileNotFoundException(
            $"Fichier introuvable : {cheminFichier}"
        );

      int maxId =
          0;

      using (ExcelPackage package = new(new FileInfo(cheminFichier))) {
        ExcelWorksheet feuille =
            package.Workbook.Worksheets[nomFeuille] ?? throw new ArgumentException(
              $"Feuille '{nomFeuille}' introuvable.",
              nameof(nomFeuille)
          );
        if (feuille.Dimension == null)
          return "1";

        for (int ligne = 2;ligne <= feuille.Dimension.End.Row;ligne++) {
          string valeur =
              GetStringByColumnName_Epplus(
                  feuille,
                  ligne,
                  nomColonneId
              );


          if (int.TryParse(valeur,out int id)) {
            if (id > maxId)
              maxId =
                  id;
          }
        }
      }

      return (maxId + 1).ToString();
    }

    //----------------------
    #endregion
  }
}