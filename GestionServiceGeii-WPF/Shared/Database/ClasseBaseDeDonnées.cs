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
║  Nom de fichier : ClasseBaseDeDonnées.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using GestionServiceGeii.Core.Excel;
using GestionServiceGeii.Core.Service;
using GestionServiceGeii.Core.Services;
using GestionServiceGeii.Shared.Librairie_Fichier;
using GestionServiceGeii.Shared.Librairie_Générique;
using GestionServiceGeii.UI;
using OfficeOpenXml;
using static GestionServiceGeii.Shared.Librairie_GEII.ServiceGeii;
using DataTable = System.Data.DataTable;
using Label = System.Windows.Controls.Label;
using TextBox = System.Windows.Controls.TextBox;


namespace GestionServiceGeii.Shared.Database {
  class ClasseBaseDeDonnées: ClasseGénérique {

    #region Champs

    private static List<FicheMatière> fiche = [];

    private static DataSet dataSetExcel = new();

    private static DataTable dataTableExcel = new();

    private static DialogueServiceGEII DialogueServiceGeii = new();

    private static ushort compteurCM = new();

    private static ushort compteurTD = new();

    private static ushort compteurTP = new();

    private static readonly Dictionary<string,Dictionary<string,List<string>>> _cacheGroupesSemestres =
    new(StringComparer.OrdinalIgnoreCase);

    private static readonly object _cacheGroupesSemestresLock = new();

    private const int MaxColonnesRechercheGroupes = 300;

    private const int MaxLignesRechercheSemestre = 1200;

    private const int StopColonnesVidesConsecutives = 25;

    #endregion Champs

    #region Propriétés

    internal static DataSet DataSetExcel {
      get => dataSetExcel ??= new DataSet();
      set => dataSetExcel = value ?? dataSetExcel;
    }

    internal static DataTable DataTableExcel {
      get => dataTableExcel ??= new DataTable();
      set => dataTableExcel = value ?? dataTableExcel;
    }

    internal static DialogueServiceGEII DialogueServiceGEII {
      get => DialogueServiceGeii ??= new DialogueServiceGEII();
      set => DialogueServiceGeii = value ?? DialogueServiceGeii;
    }

    internal static ushort CompteurCM {
      get => compteurCM;
      set => compteurCM = value;
    }

    internal static ushort CompteurTD {
      get => compteurTD;
      set => compteurTD = value;
    }

    internal static ushort CompteurTP {
      get => compteurTP;
      set => compteurTP = value;
    }

    #endregion Propriétés

    #region Lecture des classeurs Excel

    internal static int MiseAJourNomDepuisCelluleSourceSemestre_Epplus(
    ClasseExcel fichierDialogue,
    DataRowView row
) {
      if (fichierDialogue == null)
        return 0;

      if (row == null)
        return 0;

      if (string.IsNullOrWhiteSpace(fichierDialogue.CheminFichier))
        return 0;

      string nomFeuille =
          GetBoundRowValue_vue(
              row,
              ExcelSchemaNames.Columns.SourceSheet).Trim();

      string sourceRowText =
          GetBoundRowValue_vue(
              row,
              ExcelSchemaNames.Columns.SourceRow).Trim();

      string sourceColumnText =
          GetBoundRowValue_vue(
              row,
              ExcelSchemaNames.Columns.SourceColumn).Trim();

      string nouveauNom =
          GetBoundRowValue_vue(
              row,
              ExcelSchemaNames.Columns.Noms).Trim();

      if (string.IsNullOrWhiteSpace(nomFeuille)) {
        MessageBox.Show(
            "Impossible de modifier la cellule source : SourceSheet est vide." +
            Environment.NewLine +
            Environment.NewLine +
            "La ligne affichée ne porte pas encore l'adresse de sa cellule source S1/S2.",
            "Mise à jour source Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      if (
          !string.Equals(nomFeuille,"S1",StringComparison.OrdinalIgnoreCase) &&
          !string.Equals(nomFeuille,"S2",StringComparison.OrdinalIgnoreCase)
      ) {
        MessageBox.Show(
            "Modification refusée : la cellule source n'est pas dans S1 ou S2." +
            Environment.NewLine +
            Environment.NewLine +
            "Feuille source trouvée : " + nomFeuille,
            "Mise à jour source Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      if (!int.TryParse(sourceRowText,out int sourceRow) || sourceRow <= 0) {
        MessageBox.Show(
            "Impossible de modifier la cellule source : SourceRow est invalide.",
            "Mise à jour source Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      if (!int.TryParse(sourceColumnText,out int sourceColumn) || sourceColumn <= 0) {
        MessageBox.Show(
            "Impossible de modifier la cellule source : SourceColumn est invalide.",
            "Mise à jour source Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      FileInfo fichier =
          new(fichierDialogue.CheminFichier);

      if (!fichier.Exists)
        return 0;

      using ExcelPackage package =
          new(fichier);

      ExcelWorksheet? feuille =
          package.Workbook.Worksheets[nomFeuille];

      if (feuille == null) {
        MessageBox.Show(
            "Feuille source introuvable : " + nomFeuille,
            "Mise à jour source Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      feuille.Cells[sourceRow,sourceColumn].Value =
          nouveauNom;

      package.Save();

      MessageBox.Show(
          "Cellule source modifiée : " +
          nomFeuille +
          "!" +
          feuille.Cells[sourceRow,sourceColumn].Address,
          "Mise à jour source Excel",
          MessageBoxButton.OK,
          MessageBoxImage.Information);

      return 1;
    }

    public static DataSet LectureFichierDeSélection(ClasseExcel fichierDeSelection) {
      DataSet dataSetSelection = new("FICHIER DE SELECTION");

      if (fichierDeSelection == null)
        return dataSetSelection;

      if (string.IsNullOrWhiteSpace(fichierDeSelection.CheminFichier))
        return dataSetSelection;

      try {
        dataSetSelection =
            ClasseEpplus.LireToutesPlagesNommesExcel_Epplus(
                fichierDeSelection.CheminFichier);

        if (!dataSetSelection.Tables.Contains(ExcelSchemaNames.Tables.Formation)) {
          MessageBox.Show(
              "Le fichier sélectionné ne semble pas être le fichier de sélection.",
              "Lecture fichier de sélection",
              MessageBoxButton.OK,
              MessageBoxImage.Warning);

          return new DataSet("FICHIER DE SELECTION");
        }
      }
      catch (Exception ex) {
        MessageBox.Show(
            "Erreur lors de la lecture du fichier de sélection." +
            Environment.NewLine +
            Environment.NewLine +
            ex.Message,
            "Lecture fichier de sélection",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
      }

      return dataSetSelection;
    }

    /// <summary>
    /// Lit les feuilles utiles du fichier ExcelApp de service via OleDb et les ajoute
    /// au DataSet de service.
    /// </summary>
    /// <param name="fichierDeService">Fichier ExcelApp de service à lire.</param>
    /// <returns>DataSet contenant les tables lues depuis le fichier de service.</returns>
    public static DataSet LectureFichierDeService(ClasseExcel fichierDeService) {
      Debug.WriteLine(">>> ENTREE LectureFichierDeService");
      DataSet serviceDataSet = new("FICHIER DE SERVICE");

      if (fichierDeService == null) {
        MessageBox.Show(
            "Le fichier de service n'est pas initialisé.",
            "Lecture fichier de service",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return serviceDataSet;
      }

      if (string.IsNullOrWhiteSpace(fichierDeService.CheminFichier))
        return serviceDataSet;

      if (!File.Exists(fichierDeService.CheminFichier))
        return serviceDataSet;

      try {
        string cheminFichier = fichierDeService.CheminFichier;

        //if (ServiceWorkbookDiskCache.TryLoadServiceDataSet(
        //        cheminFichier,
        //        out DataSet? serviceDataSetDepuisCache)) {

        //  Debug.WriteLine(">>> SERVICE CHARGE DEPUIS CACHE");

        //  if (serviceDataSetDepuisCache != null) {
        //    DataSetExcel = serviceDataSetDepuisCache;

        //    if (serviceDataSetDepuisCache.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal))
        //      DataTableExcel = serviceDataSetDepuisCache.Tables[ExcelSchemaNames.Tables.NomTableGlobal]!;

        //    return serviceDataSetDepuisCache;
        //  }
        //}

        using (ExcelPackage package = new(new FileInfo(cheminFichier))) {
          ExcelWorksheet? feuilleGlobal = package.Workbook.Worksheets[ExcelSchemaNames.Tables.NomTableGlobal];

          if (feuilleGlobal == null)
            throw new InvalidOperationException("Feuille introuvable : " + ExcelSchemaNames.Tables.NomTableGlobal);

          DataTable tableGlobal = LireFeuilleExcelCommeDataTable_Epplus(feuilleGlobal,ExcelSchemaNames.Tables.NomTableGlobal);
          tableGlobal.TableName = ExcelSchemaNames.Tables.NomTableGlobal;

          NormaliserColonnesGlobalPourInterface(tableGlobal);
          EnrichirCodesGlobalDepuisLibelleCourt(tableGlobal);
          EnrichirGlobalDepuisLignesTechniques(tableGlobal);

          EnrichirGroupesGlobalDepuisFeuilleSemestre_Epplus(cheminFichier,tableGlobal,"S1");
          AjouterLignesManquantesDepuisFeuilleSemestre_Epplus(cheminFichier,tableGlobal,"S1");
          EnrichirInfosGlobalDepuisFeuilleSemestre_Epplus(cheminFichier,tableGlobal,"S1");

          DiagnostiquerLigneServiceFeuilleSemestre(cheminFichier,"S1","R1-10-TP");

          serviceDataSet.Tables.Add(tableGlobal);

          foreach (ServiceWorkbookSheetInfo sheetInfo in ServiceWorkbookSheetCatalog.SheetsToCache) {
            if (!sheetInfo.UseForGroupEnrichment)
              continue;

            if (package.Workbook.Worksheets[sheetInfo.SheetName] == null)
              continue;

            EnrichirGroupesGlobalDepuisFeuilleSemestre_Epplus(cheminFichier,tableGlobal,sheetInfo.SheetName);
          }
        }

        foreach (DataTable table in serviceDataSet.Tables)
          Debug.WriteLine("Table service chargée : " + table.TableName + " | " + table.Rows.Count + " lignes");

        ServiceWorkbookDiskCache.SaveServiceDataSet(cheminFichier,serviceDataSet);

        Debug.WriteLine(">>> SERVICE LU DEPUIS EXCEL");

        return serviceDataSet;
      }
      catch (Exception ex) {
        MessageBox.Show(
            "Le fichier sélectionné ne semble pas être le nouveau fichier de service." +
            Environment.NewLine +
            Environment.NewLine +
            "La feuille attendue '" +
            ExcelSchemaNames.Tables.NomTableGlobal +
            "' est introuvable ou illisible." +
            Environment.NewLine +
            Environment.NewLine +
            ex.Message,
            "Lecture fichier de service",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        DataSetExcel = serviceDataSet;

        if (serviceDataSet.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal))
          DataTableExcel = serviceDataSet.Tables[ExcelSchemaNames.Tables.NomTableGlobal]!;

        Debug.WriteLine(">>> SERVICE LU DEPUIS EXCEL");

        return serviceDataSet;
      }
    }

    private static void RemplacerTablesServiceDansDataSet(
    DataSet serviceDataSet
) {
      if (serviceDataSet == null)
        return;

      DataSetExcel ??= new DataSet();

      foreach (DataTable table in serviceDataSet.Tables) {
        if (string.IsNullOrWhiteSpace(table.TableName))
          continue;

        if (DataSetExcel.Tables.Contains(table.TableName)) {
          DataSetExcel.Tables.Remove(table.TableName);
        }

        DataTable copie =
            table.Copy();

        copie.TableName =
            table.TableName;

        DataSetExcel.Tables.Add(
            copie);
      }

      if (DataSetExcel.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal)) {
        DataTableExcel =
            DataSetExcel.Tables[ExcelSchemaNames.Tables.NomTableGlobal];
      }
    }

    private static DataTable LireFeuilleExcelCommeDataTable_Epplus(
    ExcelWorksheet feuille,
    string nomTable
) {
      DataTable table =
          new DataTable(nomTable);

      if (feuille == null)
        return table;

      if (feuille.Dimension == null)
        return table;

      int rowMax =
          feuille.Dimension.End.Row;

      int colMax =
          feuille.Dimension.End.Column;

      if (!string.Equals(
              nomTable,
              ExcelSchemaNames.Tables.NomTableGlobal,
              StringComparison.OrdinalIgnoreCase)) {
        rowMax =
            Math.Min(rowMax,1500);

        colMax =
            Math.Min(colMax,300);
      }

      int ligneEntetes =
          1;

      for (int col = 1;col <= colMax;col++) {
        string nomColonne =
            feuille.Cells[ligneEntetes,col].Text.Trim();

        if (string.IsNullOrWhiteSpace(nomColonne))
          nomColonne =
              "Colonne_" + col;

        nomColonne =
            CreerNomColonneUnique(
                table,
                nomColonne);

        table.Columns.Add(
            nomColonne,
            typeof(string));
      }

      for (int row = ligneEntetes + 1;row <= rowMax;row++) {
        DataRow ligne =
            table.NewRow();

        bool ligneVide =
            true;

        for (int col = 1;col <= colMax;col++) {
          string valeur =
              feuille.Cells[row,col].Text.Trim();

          if (!string.IsNullOrWhiteSpace(valeur))
            ligneVide = false;

          ligne[col - 1] =
              valeur;
        }

        if (!ligneVide)
          table.Rows.Add(ligne);
      }

      table.AcceptChanges();

      return table;
    }

    #endregion Lecture des classeurs Excel

    #region Normalisation et enrichissement de la feuille Global

    internal static void NormaliserColonnesGlobalPourInterface(DataTable table) {
      if (table == null)
        return;
      ClasseEpplus.ConfigureEpplusLicense();

      RenommerColonneSiExiste(table,"SEMESTRE",ExcelSchemaNames.Columns.Semestre);
      RenommerColonneSiExiste(table,"INTITULE",ExcelSchemaNames.Columns.Module);
      RenommerColonneSiExiste(table,"TYPE",ExcelSchemaNames.Columns.Cours);

      RenommerColonneSiExiste(table,"INTERVENANT",ExcelSchemaNames.Columns.Noms);
      RenommerColonneSiExiste(table,"ENSEIGNANT",ExcelSchemaNames.Columns.Noms);
      RenommerColonneSiExiste(table,"NOMS",ExcelSchemaNames.Columns.Noms);

      RenommerColonneSiExiste(table,"DIPLÔME",ExcelSchemaNames.Columns.Formation);
      RenommerColonneSiExiste(table,"DIPLOME",ExcelSchemaNames.Columns.Formation);
      RenommerColonneSiExiste(table,"FORMATION",ExcelSchemaNames.Columns.Formation);

      RenommerColonneSiExiste(table,"DURÉE",ExcelSchemaNames.Columns.Duree);
      RenommerColonneSiExiste(table,"DUREE",ExcelSchemaNames.Columns.Duree);

      RenommerColonneSiExiste(table,"TOTAL_TYPE",ExcelSchemaNames.Columns.TotalType);
      RenommerColonneSiExiste(table,"TOTAL TYPE",ExcelSchemaNames.Columns.TotalType);

      RenommerColonneSiExiste(table,"GROUPE",ExcelSchemaNames.Columns.Groupe);
      RenommerColonneSiExiste(table,"OSE",ExcelSchemaNames.Columns.OSE);
      RenommerColonneSiExiste(table,"SALLE",ExcelSchemaNames.Columns.Salle);
      RenommerColonneSiExiste(table,"SALLES",ExcelSchemaNames.Columns.Salle);

      RenommerColonneSiExiste(table,ExcelSchemaNames.Columns.LibelleCourt,ExcelSchemaNames.Columns.LibelleCourt);
      RenommerColonneSiExiste(table,ExcelSchemaNames.Columns.Infos,ExcelSchemaNames.Columns.Infos);
      RenommerColonneSiExiste(table,ExcelSchemaNames.Columns.StatutIntervenant,ExcelSchemaNames.Columns.StatutIntervenant);
      RenommerColonneSiExiste(table,ExcelSchemaNames.Columns.PN,ExcelSchemaNames.Columns.PN);

      GarantirColonne(table,ExcelSchemaNames.Columns.ID,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.Semestre,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.Formation,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.Module,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.Cours,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.Noms,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.Groupe,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.Nombre,"1");
      GarantirColonne(table,ExcelSchemaNames.Columns.Duree,"0");
      GarantirColonne(table,ExcelSchemaNames.Columns.TotalType,"0");
      GarantirColonne(table,ExcelSchemaNames.Columns.OSE,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.Salle,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.Libelle,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.PPN,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.LibelleCourt,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.Infos,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.StatutIntervenant,string.Empty);
      GarantirColonne(table,ExcelSchemaNames.Columns.PN,string.Empty);

      int id = 1;

      foreach (DataRow row in table.Rows) {
        if (string.IsNullOrWhiteSpace(row[ExcelSchemaNames.Columns.ID]?.ToString()))
          row[ExcelSchemaNames.Columns.ID] = id.ToString();

        if (string.IsNullOrWhiteSpace(row[ExcelSchemaNames.Columns.Nombre]?.ToString()))
          row[ExcelSchemaNames.Columns.Nombre] = "1";

        if (string.IsNullOrWhiteSpace(row[ExcelSchemaNames.Columns.Duree]?.ToString()))
          row[ExcelSchemaNames.Columns.Duree] = "0";

        if (string.IsNullOrWhiteSpace(row[ExcelSchemaNames.Columns.TotalType]?.ToString()))
          row[ExcelSchemaNames.Columns.TotalType] =
              row[ExcelSchemaNames.Columns.Duree]?.ToString() ?? "0";

        id++;
      }

      table.AcceptChanges();
    }

    private static void RenommerColonneSiExiste(DataTable table,string ancienNom,string nouveauNom) {
      if (table == null || string.IsNullOrWhiteSpace(ancienNom) || string.IsNullOrWhiteSpace(nouveauNom))
        return;

      DataColumn? colonneSource = null;

      foreach (DataColumn colonne in table.Columns) {
        if (string.Equals(colonne.ColumnName,ancienNom,StringComparison.OrdinalIgnoreCase)) {
          colonneSource = colonne;
          break;
        }
      }

      if (colonneSource == null)
        return;

      if (string.Equals(colonneSource.ColumnName,nouveauNom,StringComparison.Ordinal))
        return;

      foreach (DataColumn colonne in table.Columns) {
        if (ReferenceEquals(colonne,colonneSource))
          continue;

        if (string.Equals(colonne.ColumnName,nouveauNom,StringComparison.OrdinalIgnoreCase))
          return;
      }

      colonneSource.ColumnName = nouveauNom;
    }

    private static void GarantirColonne(
    DataTable table,
    string nomColonne,
    string valeurParDefaut
) {
      if (table == null)
        return;

      if (string.IsNullOrWhiteSpace(nomColonne))
        return;

      foreach (DataColumn colonne in table.Columns) {
        if (string.Equals(
                colonne.ColumnName,
                nomColonne,
                StringComparison.OrdinalIgnoreCase))
          return;
      }

      table.Columns.Add(nomColonne,typeof(string));

      foreach (DataRow row in table.Rows)
        row[nomColonne] = valeurParDefaut ?? string.Empty;
    }

    internal static void EnrichirCodesGlobalDepuisLibelleCourt(DataTable table) {
      ClasseEpplus.ConfigureEpplusLicense();
      if (table == null)
        return;

      string colonneLibelleCourt =
          TrouverNomColonneIgnoreCase(
              table,
              ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      if (!table.Columns.Contains(ExcelSchemaNames.Columns.OSE))
        return;

      foreach (DataRow row in table.Rows) {
        string libelleCourt =
            GetRowValueIfColumnExists(
                row,
                colonneLibelleCourt);

        string codeCourtModule =
            ExtraireCodeCourtModuleDepuisLibelleCourt(
                libelleCourt);

        if (string.IsNullOrWhiteSpace(codeCourtModule))
          continue;

        string codeOse =
            ConvertirCodeCourtEnCodeOse(
                codeCourtModule);

        if (string.IsNullOrWhiteSpace(codeOse))
          continue;

        string oseActuel =
            GetRowValueIfColumnExists(
                row,
                ExcelSchemaNames.Columns.OSE);

        if (
            string.IsNullOrWhiteSpace(oseActuel) ||
            oseActuel == "0" ||
            oseActuel.Contains("#N/A",StringComparison.OrdinalIgnoreCase)
        ) {
          row[ExcelSchemaNames.Columns.OSE] =
              codeOse;
        }
      }

      table.AcceptChanges();
    }

    internal static string ExtraireCodeCourtModuleDepuisLibelleCourt(string libelleCourt) {
      if (string.IsNullOrWhiteSpace(libelleCourt))
        return string.Empty;

      string texte =
          libelleCourt.Trim();

      if (
          texte.EndsWith("-CM",StringComparison.OrdinalIgnoreCase) ||
          texte.EndsWith("-TD",StringComparison.OrdinalIgnoreCase) ||
          texte.EndsWith("-TP",StringComparison.OrdinalIgnoreCase) ||
          texte.EndsWith("-DS",StringComparison.OrdinalIgnoreCase)
      ) {
        return texte.Substring(0,texte.Length - 3);
      }

      return texte;
    }

    internal static string ConvertirCodeCourtEnCodeOse(string codeCourt) {
      Match match =
          Regex.Match(
              codeCourt,
              @"^R(\d+)-(\d{2})$",
              RegexOptions.IgnoreCase);

      if (!match.Success)
        return string.Empty;

      string semestre =
          match.Groups[1].Value;

      string numero =
          match.Groups[2].Value;

      return $"F2R{semestre}{numero}";
    }

    internal static void EnrichirGlobalDepuisLignesTechniques(DataTable table) {
      ClasseEpplus.ConfigureEpplusLicense();
      Dictionary<string,string> infosParCle = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);

      if (table == null)
        return;

      if (!table.Columns.Contains(ExcelSchemaNames.Columns.Semestre))
        return;

      if (!table.Columns.Contains(ExcelSchemaNames.Columns.Formation))
        return;

      if (!table.Columns.Contains(ExcelSchemaNames.Columns.Module))
        return;

      if (!table.Columns.Contains(ExcelSchemaNames.Columns.Cours))
        return;

      if (!table.Columns.Contains(ExcelSchemaNames.Columns.Noms))
        return;

      if (!table.Columns.Contains(ExcelSchemaNames.Columns.TotalType))
        return;

      if (!table.Columns.Contains(ExcelSchemaNames.Columns.Duree))
        return;

      string colonneLibelleCourt =
          TrouverNomColonneIgnoreCase(table,"LIBELLE COURT");

      Dictionary<string,int> totalTypeParCle =
          new Dictionary<string,int>(StringComparer.OrdinalIgnoreCase);

      foreach (DataRow row in table.Rows) {
        string noms =
            GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Noms);

        if (!EstLigneTechniqueGlobal(noms))
          continue;

        string cle =
            ConstruireCleGlobal(
                row,
                colonneLibelleCourt);

        if (table.Columns.Contains(ExcelSchemaNames.Columns.Infos)) {
          string infos =
              GetRowValueIfColumnExists(
                  row,
                  ExcelSchemaNames.Columns.Infos);

          if (
              !string.IsNullOrWhiteSpace(infos) &&
              infos != "0" &&
              !infos.Contains("#REF!",StringComparison.OrdinalIgnoreCase)
          ) {
            infosParCle[cle] = infos;
          }
        }


        if (string.IsNullOrWhiteSpace(cle))
          continue;

        int totalType =
            LireEntierDepuisRow(
                row,
                ExcelSchemaNames.Columns.TotalType);

        if (totalType <= 0)
          continue;

        totalTypeParCle[cle] =
            totalType;
      }

      foreach (DataRow row in table.Rows) {
        string noms =
            GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Noms);

        if (EstLigneTechniqueGlobal(noms))
          continue;

        if (string.IsNullOrWhiteSpace(noms))
          continue;

        string cle =
            ConstruireCleGlobal(
                row,
                colonneLibelleCourt);

        if (
          table.Columns.Contains(ExcelSchemaNames.Columns.Infos) &&
          infosParCle.TryGetValue(cle,out string? infos)
) {
          string infosActuelles =
              GetRowValueIfColumnExists(
                  row,
                  ExcelSchemaNames.Columns.Infos);

          if (
              string.IsNullOrWhiteSpace(infosActuelles) ||
              infosActuelles == "0"
          ) {
            row[ExcelSchemaNames.Columns.Infos] =
                infos;
          }
        }


        if (string.IsNullOrWhiteSpace(cle))
          continue;

        if (!totalTypeParCle.TryGetValue(cle,out int totalType))
          continue;

        int dureeActuelle =
            LireEntierDepuisRow(
                row,
                ExcelSchemaNames.Columns.Duree);

        int totalTypeActuel =
            LireEntierDepuisRow(
                row,
                ExcelSchemaNames.Columns.TotalType);

        if (dureeActuelle == 0)
          row[ExcelSchemaNames.Columns.Duree] = totalType;

        if (totalTypeActuel == 0)
          row[ExcelSchemaNames.Columns.TotalType] = totalType;
      }

      table.AcceptChanges();
    }

    internal static void EnrichirInfosGlobalDepuisFeuilleSemestre_Epplus(
    string cheminFichier,
    DataTable tableGlobal,
    string nomFeuille
) {
      if (string.IsNullOrWhiteSpace(cheminFichier))
        return;

      if (!File.Exists(cheminFichier))
        return;

      if (tableGlobal == null)
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Semestre))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Module))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Cours))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Noms))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Groupe))
        return;

      string colonneLibelleCourt =
          TrouverNomColonneIgnoreCase(
              tableGlobal,
              ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      string colonneInfos =
          TrouverNomColonneIgnoreCase(
              tableGlobal,
              ExcelSchemaNames.Columns.Infos);

      if (string.IsNullOrWhiteSpace(colonneInfos)) {
        tableGlobal.Columns.Add(
            ExcelSchemaNames.Columns.Infos,
            typeof(string));

        colonneInfos =
            ExcelSchemaNames.Columns.Infos;
      }

      using (ExcelPackage package =
          new ExcelPackage(new FileInfo(cheminFichier))) {
        ExcelWorksheet? feuille =
            package.Workbook.Worksheets[nomFeuille];

        if (feuille == null || feuille.Dimension == null)
          return;

        int ligneSalles;
        int colonneSalles;

        if (!TrouverCelluleTexteFeuilleSemestre(
                feuille,
                "SALLES",
                out ligneSalles,
                out colonneSalles)) {
          return;
        }

        int colMax =
            TrouverDerniereColonneGroupes(
                feuille,
                ligneSalles,
                ligneSalles + 1,
                colonneSalles + 1);

        if (colMax < colonneSalles + 1)
          return;

        Dictionary<int,string> groupesTd =
            LireGroupesFeuilleSemestre(
                feuille,
                ligneSalles,
                colonneSalles + 1,
                colMax,
                estTp: false);

        Dictionary<int,string> groupesTp =
            LireGroupesFeuilleSemestre(
                feuille,
                ligneSalles + 1,
                colonneSalles + 1,
                colMax,
                estTp: true);

        Regex regexCodeService =
            new Regex(
                @"^(R\d+-\d{2})-(CM|TD|TP|DS)$",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        int rowMax =
            TrouverDerniereLigneUtileSemestre(
                feuille);

        string moduleCourant =
            string.Empty;

        string infosCourantes =
            string.Empty;

        int lignesModifiees =
            0;

        for (int row = 1;row <= rowMax;row++) {
          string celluleA =
              feuille.Cells[row,1].Text.Trim();

          string moduleDetecte =
              ExtraireModuleDepuisCelluleFeuilleSemestre(
                  celluleA);

          if (EstCodeModuleFeuilleSemestreValide(moduleDetecte)) {
            moduleCourant =
                moduleDetecte;

            infosCourantes =
                string.Empty;
          }
          else if (
              !string.IsNullOrWhiteSpace(celluleA) &&
              EstInfoFeuilleSemestreValide(celluleA)
          ) {
            infosCourantes =
                celluleA;
          }

          string codeService =
              feuille.Cells[row,2].Text.Trim();

          Match match =
              regexCodeService.Match(codeService);

          if (!match.Success)
            continue;

          if (string.IsNullOrWhiteSpace(moduleCourant))
            continue;

          string cours =
              match.Groups[2].Value.ToUpperInvariant();

          if (cours != "TD" && cours != "TP")
            continue;

          if (string.IsNullOrWhiteSpace(infosCourantes))
            continue;

          Dictionary<int,string> groupes =
              cours == "TD"
                  ? groupesTd
                  : groupesTp;

          int dureeLigne =
              LireDureeDepuisLigneFeuilleSemestre(
                  feuille,
                  row,
                  colonneSalles);

          foreach (KeyValuePair<int,string> groupe in groupes) {
            string nomIntervenant =
                feuille.Cells[row,groupe.Key].Text.Trim();

            if (!EstNomIntervenantValidePourGroupe(nomIntervenant))
              continue;

            foreach (DataRow globalRow in tableGlobal.Rows) {
              if (!LigneGlobalCorrespondInfoSemestre(
                      globalRow,
                      nomFeuille,
                      moduleCourant,
                      codeService,
                      cours,
                      nomIntervenant,
                      groupe.Value,
                      dureeLigne,
                      colonneLibelleCourt)) {
                continue;
              }

              string infosActuelles =
                  GetRowValueIfColumnExists(
                      globalRow,
                      colonneInfos);

              if (
                  !string.IsNullOrWhiteSpace(infosActuelles) &&
                  infosActuelles != "0"
              ) {
                continue;
              }

              globalRow[colonneInfos] =
                  infosCourantes;

              lignesModifiees++;
            }
          }
        }

        if (lignesModifiees > 0) {
          Debug.WriteLine(
              "INFOS enrichies depuis " +
              nomFeuille +
              " : " +
              lignesModifiees +
              " ligne(s)");
        }
      }

      tableGlobal.AcceptChanges();
    }

    private static bool EstCodeModuleFeuilleSemestreValide(
        string module
    ) {
      if (string.IsNullOrWhiteSpace(module))
        return false;

      return Regex.IsMatch(
          module.Trim(),
          @"^[A-Z]{3,}[A-Z0-9]*\d+(?:-\d+)?$",
          RegexOptions.IgnoreCase);
    }

    private static bool EstInfoFeuilleSemestreValide(
        string valeur
    ) {
      if (string.IsNullOrWhiteSpace(valeur))
        return false;

      string texte =
          valeur.Trim();

      if (texte == "0")
        return false;

      if (texte.Contains("#REF!",StringComparison.OrdinalIgnoreCase))
        return false;

      if (texte.Equals("#N/A",StringComparison.OrdinalIgnoreCase))
        return false;

      if (texte.StartsWith("Total",StringComparison.OrdinalIgnoreCase))
        return false;

      if (texte.StartsWith("Nb",StringComparison.OrdinalIgnoreCase))
        return false;

      if (Regex.IsMatch(
              texte,
              @"^(R\d+-\d{2})-(CM|TD|TP|DS)$",
              RegexOptions.IgnoreCase)) {
        return false;
      }

      return true;
    }

    private static bool LigneGlobalCorrespondInfoSemestre(
        DataRow row,
        string semestre,
        string module,
        string libelleCourt,
        string cours,
        string noms,
        string groupe,
        int duree,
        string colonneLibelleCourt
    ) {
      string semestreGlobal =
          GetRowValueIfColumnExists(
              row,
              ExcelSchemaNames.Columns.Semestre);

      if (!string.Equals(
              semestreGlobal,
              semestre,
              StringComparison.OrdinalIgnoreCase)) {
        return false;
      }

      string moduleGlobal =
          GetRowValueIfColumnExists(
              row,
              ExcelSchemaNames.Columns.Module);

      if (!ModuleCorrespond(
              moduleGlobal,
              module)) {
        return false;
      }

      string libelleCourtGlobal =
          GetRowValueIfColumnExists(
              row,
              colonneLibelleCourt);

      if (!string.Equals(
              libelleCourtGlobal,
              libelleCourt,
              StringComparison.OrdinalIgnoreCase)) {
        return false;
      }

      string coursGlobal =
          GetRowValueIfColumnExists(
              row,
              ExcelSchemaNames.Columns.Cours);

      if (!string.Equals(
              coursGlobal,
              cours,
              StringComparison.OrdinalIgnoreCase)) {
        return false;
      }

      string nomsGlobal =
          GetRowValueIfColumnExists(
              row,
              ExcelSchemaNames.Columns.Noms);

      if (!string.Equals(
              nomsGlobal,
              noms,
              StringComparison.OrdinalIgnoreCase)) {
        return false;
      }

      string groupeGlobal =
          GetRowValueIfColumnExists(
              row,
              ExcelSchemaNames.Columns.Groupe);

      if (!string.Equals(
              groupeGlobal,
              groupe,
              StringComparison.OrdinalIgnoreCase)) {
        return false;
      }

      if (duree > 0) {
        int dureeGlobal =
            LireEntierDepuisRow(
                row,
                ExcelSchemaNames.Columns.Duree);

        if (dureeGlobal > 0 && dureeGlobal != duree)
          return false;
      }

      return true;
    }

    private static string ConstruireCleGlobal(DataRow row,string colonneLibelleCourt) {
      string semestre = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Semestre);
      string formation = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Formation);
      string module = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Module);
      string cours = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Cours);
      string libelleCourt = GetRowValueIfColumnExists(row,colonneLibelleCourt);
      if (string.IsNullOrWhiteSpace(semestre))
        return string.Empty;

      if (string.IsNullOrWhiteSpace(formation))
        return string.Empty;

      if (string.IsNullOrWhiteSpace(module))
        return string.Empty;

      if (string.IsNullOrWhiteSpace(cours))
        return string.Empty;

      return
          semestre.Trim().ToUpperInvariant() + "|" +
          formation.Trim().ToUpperInvariant() + "|" +
          module.Trim().ToUpperInvariant() + "|" +
          cours.Trim().ToUpperInvariant() + "|" +
          libelleCourt.Trim().ToUpperInvariant();
    }

    private static bool EstLigneTechniqueGlobal(string noms) {
      if (string.IsNullOrWhiteSpace(noms))
        return true;

      string texte =
          noms.Trim();

      if (string.Equals(texte,"0",StringComparison.OrdinalIgnoreCase))
        return true;

      if (texte.StartsWith("BUT",StringComparison.OrdinalIgnoreCase))
        return true;

      if (texte.Contains("#REF!",StringComparison.OrdinalIgnoreCase))
        return true;

      return false;
    }

    #endregion Normalisation et enrichissement de la feuille Global

    #region Enrichissement depuis les feuilles semestre

    private sealed class AffectationGroupeSemestre {
      internal string Semestre { get; set; } = string.Empty;
      internal string Module { get; set; } = string.Empty;
      internal string LibelleCourt { get; set; } = string.Empty;
      internal string Cours { get; set; } = string.Empty;
      internal string Noms { get; set; } = string.Empty;
      internal string Groupe { get; set; } = string.Empty;
      internal string SourceSheet { get; set; } = string.Empty;
      internal int SourceRow { get; set; }
      internal int SourceColumn { get; set; }
      internal int Duree { get; set; }
      internal string Infos { get; set; } = string.Empty;
    }

    internal static void EnrichirGroupesGlobalDepuisFeuilleSemestre_Epplus(
        string cheminFichier,
        DataTable tableGlobal,
        string nomFeuille
    ) {
      if (string.IsNullOrWhiteSpace(cheminFichier))
        return;

      if (!File.Exists(cheminFichier))
        return;

      if (tableGlobal == null)
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Semestre))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Module))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Cours))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Noms))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Groupe))
        return;

      string colonneLibelleCourt =
          TrouverNomColonneIgnoreCase(
              tableGlobal,
              ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      Dictionary<string,List<string>> groupesParCle =
          ObtenirGroupesParCleSemestreDepuisCache(
              cheminFichier,
              nomFeuille);

      if (groupesParCle.Count == 0)
        return;

      Dictionary<string,int> indexParCle =
          new(StringComparer.OrdinalIgnoreCase);

      foreach (DataRow row in tableGlobal.Rows) {
        string semestre =
            GetRowValueIfColumnExists(
                row,
                ExcelSchemaNames.Columns.Semestre);

        if (!string.Equals(
                semestre,
                nomFeuille,
                StringComparison.OrdinalIgnoreCase)) {
          continue;
        }

        string cours =
            GetRowValueIfColumnExists(
                row,
                ExcelSchemaNames.Columns.Cours);

        if (
            !string.Equals(cours,"TD",StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(cours,"TP",StringComparison.OrdinalIgnoreCase)
        ) {
          continue;
        }

        string noms =
            GetRowValueIfColumnExists(
                row,
                ExcelSchemaNames.Columns.Noms);

        if (EstLigneTechniqueGlobal(noms))
          continue;

        string module =
            GetRowValueIfColumnExists(
                row,
                ExcelSchemaNames.Columns.Module);

        string libelleCourt =
            GetRowValueIfColumnExists(
                row,
                colonneLibelleCourt);

        string cle =
            ConstruireCleGroupeGlobal(
                nomFeuille,
                module,
                libelleCourt,
                cours,
                noms);

        string cleUtilisee =
            cle;

        if (!groupesParCle.TryGetValue(cle,out List<string>? groupes)) {
          string moduleAvecSuffixeUn =
              AjouterSuffixeModuleUnSiNecessaire(module);

          string cleAvecSuffixeUn =
              ConstruireCleGroupeGlobal(
                  nomFeuille,
                  moduleAvecSuffixeUn,
                  libelleCourt,
                  cours,
                  noms);

          cleUtilisee =
              cleAvecSuffixeUn;

          groupesParCle.TryGetValue(
              cleAvecSuffixeUn,
              out groupes);
        }

        if (groupes == null)
          continue;

        if (!indexParCle.TryGetValue(cleUtilisee,out int index))
          index = 0;

        if (index >= groupes.Count)
          continue;

        row[ExcelSchemaNames.Columns.Groupe] =
            groupes[index];

        indexParCle[cleUtilisee] =
            index + 1;
      }

      tableGlobal.AcceptChanges();
    }

    private static Dictionary<string,List<string>> ObtenirGroupesParCleSemestreDepuisCache(
        string cheminFichier,
        string nomFeuille
    ) {
      string cleCache =
          ConstruireCleCacheGroupesSemestre(
              cheminFichier,
              nomFeuille);

      lock (_cacheGroupesSemestresLock) {
        if (_cacheGroupesSemestres.TryGetValue(
                cleCache,
                out Dictionary<string,List<string>>? groupesDepuisCache)) {
          return groupesDepuisCache;
        }
      }

      Dictionary<string,List<string>> groupes =
          ConstruireGroupesParCleSemestreDepuisClasseur(
              cheminFichier,
              nomFeuille);

      lock (_cacheGroupesSemestresLock) {
        if (!_cacheGroupesSemestres.ContainsKey(cleCache)) {
          _cacheGroupesSemestres.Add(
              cleCache,
              groupes);
        }

        return _cacheGroupesSemestres[cleCache];
      }
    }

    private static string ConstruireCleCacheGroupesSemestre(
        string cheminFichier,
        string nomFeuille
    ) {
      FileInfo fichier =
          new FileInfo(cheminFichier);

      return
          fichier.FullName.ToUpperInvariant() + "|" +
          nomFeuille.ToUpperInvariant() + "|" +
          fichier.LastWriteTimeUtc.Ticks + "|" +
          fichier.Length;
    }

    private static Dictionary<string,List<string>> ConstruireGroupesParCleSemestreDepuisClasseur(
    string cheminFichier,
    string nomFeuille
) {
      Dictionary<string,List<string>> groupesParCle =
          new(StringComparer.OrdinalIgnoreCase);

      using (ExcelPackage package =
          new ExcelPackage(new FileInfo(cheminFichier))) {
        ExcelWorksheet? feuille =
            package.Workbook.Worksheets[nomFeuille];

        if (feuille == null || feuille.Dimension == null)
          return groupesParCle;

        int ligneSalles;
        int colonneSalles;

        if (!TrouverCelluleTexteFeuilleSemestre(
                feuille,
                "SALLES",
                out ligneSalles,
                out colonneSalles)) {
          return groupesParCle;
        }

        int colFinGroupes =
            TrouverDerniereColonneGroupes(
                feuille,
                ligneSalles,
                ligneSalles + 1,
                colonneSalles + 1);

        if (colFinGroupes < colonneSalles + 1)
          return groupesParCle;

        Dictionary<int,string> groupesTd =
            LireGroupesFeuilleSemestre(
                feuille,
                ligneSalles,
                colonneSalles + 1,
                colFinGroupes,
                estTp: false);

        Dictionary<int,string> groupesTp =
            LireGroupesFeuilleSemestre(
                feuille,
                ligneSalles + 1,
                colonneSalles + 1,
                colFinGroupes,
                estTp: true);

        int rowMax =
            TrouverDerniereLigneUtileSemestre(
                feuille);

        groupesParCle =
            ConstruireGroupesParCleDepuisFeuilleSemestre(
                feuille,
                nomFeuille,
                rowMax,
                groupesTd,
                groupesTp);
      }

      return groupesParCle;
    }

    private static Dictionary<string,List<string>> ConstruireGroupesParCleDepuisFeuilleSemestre(
        ExcelWorksheet feuille,
        string nomFeuille,
        int rowMax,
        Dictionary<int,string> groupesTd,
        Dictionary<int,string> groupesTp
    ) {
      Dictionary<string,List<string>> groupesParCle =
          new(StringComparer.OrdinalIgnoreCase);

      Regex regexCodeService =
          new Regex(
              @"^(R\d+-\d{2})-(CM|TD|TP|DS)$",
              RegexOptions.IgnoreCase | RegexOptions.Compiled);

      string moduleCourant =
          string.Empty;

      for (int row = 1;row <= rowMax;row++) {
        string celluleModule =
            feuille.Cells[row,1].Text.Trim();

        string moduleDetecte =
            ExtraireModuleDepuisCelluleFeuilleSemestre(
                celluleModule);

        if (!string.IsNullOrWhiteSpace(moduleDetecte))
          moduleCourant = moduleDetecte;

        string codeService =
            feuille.Cells[row,2].Text.Trim();

        if (string.IsNullOrWhiteSpace(codeService))
          continue;

        Match match =
            regexCodeService.Match(codeService);

        if (!match.Success)
          continue;

        string typeCours =
            match.Groups[2].Value.ToUpperInvariant();

        if (typeCours != "TD" && typeCours != "TP")
          continue;

        Dictionary<int,string> groupes =
            typeCours == "TD"
                ? groupesTd
                : groupesTp;

        foreach (KeyValuePair<int,string> groupe in groupes) {
          int col =
              groupe.Key;

          string nomIntervenant =
              feuille.Cells[row,col].Text.Trim();

          if (!EstNomIntervenantValidePourGroupe(nomIntervenant))
            continue;

          string cle =
              ConstruireCleGroupeGlobal(
                  nomFeuille,
                  moduleCourant,
                  codeService,
                  typeCours,
                  nomIntervenant);

          if (string.IsNullOrWhiteSpace(cle))
            continue;

          if (!groupesParCle.TryGetValue(cle,out List<string>? listeGroupes)) {
            listeGroupes =
                new List<string>();

            groupesParCle.Add(
                cle,
                listeGroupes);
          }

          listeGroupes.Add(
              groupe.Value);
        }
      }

      return groupesParCle;
    }

    internal static void AjouterLignesManquantesDepuisFeuilleSemestre_Epplus(
        string cheminFichier,
        DataTable tableGlobal,
        string nomFeuille
    ) {
      if (string.IsNullOrWhiteSpace(cheminFichier))
        return;

      if (!File.Exists(cheminFichier))
        return;

      if (tableGlobal == null)
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Semestre))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Module))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Cours))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Noms))
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Groupe))
        return;

      string colonneLibelleCourt =
          TrouverNomColonneIgnoreCase(
              tableGlobal,
              ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      List<AffectationGroupeSemestre> affectations =
          LireAffectationsGroupesDepuisFeuilleSemestre_Epplus(
              cheminFichier,
              nomFeuille);

      if (affectations.Count == 0)
        return;

      int prochainId =
          TrouverProchainIdGlobal(tableGlobal);

      AjouterColonnesSourceSemestreSiAbsentes(tableGlobal);

      ReconstruireLignesInfosSpecialesDepuisAffectations(
          tableGlobal,
          colonneLibelleCourt,
          affectations);

      foreach (AffectationGroupeSemestre affectation in affectations) {
        if (string.IsNullOrWhiteSpace(affectation.Noms))
          continue;

        if (EstLigneTechniqueGlobal(affectation.Noms))
          continue;

        if (ExisteDejaDansGlobalAvecMemeActivite(
                tableGlobal,
                colonneLibelleCourt,
                affectation)) {
          continue;
        }

        DataRow? modele =
            TrouverLigneModeleGlobal(
                tableGlobal,
                colonneLibelleCourt,
                affectation);

        if (modele == null)
          continue;

        DataRow nouvelleLigne =
            tableGlobal.NewRow();

        foreach (DataColumn colonne in tableGlobal.Columns) {
          nouvelleLigne[colonne.ColumnName] =
              modele[colonne.ColumnName];
        }

        if (tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.ID)) {
          nouvelleLigne[ExcelSchemaNames.Columns.ID] =
              prochainId.ToString();

          prochainId++;
        }

        nouvelleLigne[ExcelSchemaNames.Columns.Noms] =
            affectation.Noms;

        nouvelleLigne[ExcelSchemaNames.Columns.Groupe] =
            affectation.Groupe;

        nouvelleLigne[ExcelSchemaNames.Columns.SourceSheet] =
          affectation.SourceSheet;

        nouvelleLigne[ExcelSchemaNames.Columns.SourceRow] =
            affectation.SourceRow.ToString();

        nouvelleLigne[ExcelSchemaNames.Columns.SourceColumn] =
            affectation.SourceColumn.ToString();

        nouvelleLigne[ExcelSchemaNames.Columns.Nombre] =
            "1";

        nouvelleLigne[ExcelSchemaNames.Columns.Cours] =
            affectation.Cours;

        nouvelleLigne[ExcelSchemaNames.Columns.Module] =
            modele[ExcelSchemaNames.Columns.Module]?.ToString() ?? affectation.Module;

        nouvelleLigne[colonneLibelleCourt] =
            affectation.LibelleCourt;

        if (affectation.Duree > 0) {
          nouvelleLigne[ExcelSchemaNames.Columns.Duree] =
              affectation.Duree.ToString();

          nouvelleLigne[ExcelSchemaNames.Columns.TotalType] =
              affectation.Duree.ToString();
        }

        if (tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Infos)) {
          nouvelleLigne[ExcelSchemaNames.Columns.Infos] =
              affectation.Infos?.Trim() ?? string.Empty;
        }

        tableGlobal.Rows.Add(
            nouvelleLigne);

        if (string.Equals(
                affectation.Module,
                "ENER1-2",
                StringComparison.OrdinalIgnoreCase)) {
          Debug.WriteLine(
              "Ligne ajoutée ENER1-2 : " +
              affectation.LibelleCourt +
              " | " +
              affectation.Cours +
              " | " +
              affectation.Noms +
              " | " +
              affectation.Groupe +
              " | " +
              affectation.Duree +
              "h" +
              " | ligne Excel " +
              affectation.SourceRow);
        }
      }

      Debug.WriteLine(
          "Ajout lignes manquantes depuis " +
          nomFeuille +
          " terminé.");

      tableGlobal.AcceptChanges();
    }

    private static void AjouterColonnesSourceSemestreSiAbsentes(DataTable tableGlobal) {
      if (tableGlobal == null)
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.SourceSheet))
        tableGlobal.Columns.Add(ExcelSchemaNames.Columns.SourceSheet,typeof(string));

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.SourceRow))
        tableGlobal.Columns.Add(ExcelSchemaNames.Columns.SourceRow,typeof(string));

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.SourceColumn))
        tableGlobal.Columns.Add(ExcelSchemaNames.Columns.SourceColumn,typeof(string));
    }

    private static void ReconstruireLignesInfosSpecialesDepuisAffectations(DataTable tableGlobal,string colonneLibelleCourt,List<AffectationGroupeSemestre> affectations) {
      if (tableGlobal == null)
        return;

      if (affectations == null || affectations.Count == 0)
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Infos))
        return;

      List<AffectationGroupeSemestre> affectationsSpeciales =
          affectations
              .Where(a => !string.IsNullOrWhiteSpace(a.Infos))
              .ToList();

      if (affectationsSpeciales.Count == 0)
        return;

      List<DataRow> lignesASupprimer =
          new List<DataRow>();

      foreach (DataRow row in tableGlobal.Rows) {
        string infosGlobal =
            GetRowValueIfColumnExists(
                row,
                ExcelSchemaNames.Columns.Infos);

        if (string.IsNullOrWhiteSpace(infosGlobal))
          continue;

        foreach (AffectationGroupeSemestre affectation in affectationsSpeciales) {
          if (!string.Equals(
                  infosGlobal,
                  affectation.Infos,
                  StringComparison.OrdinalIgnoreCase)) {
            continue;
          }

          string semestreGlobal =
              GetRowValueIfColumnExists(
                  row,
                  ExcelSchemaNames.Columns.Semestre);

          if (!string.Equals(
                  semestreGlobal,
                  affectation.Semestre,
                  StringComparison.OrdinalIgnoreCase)) {
            continue;
          }

          string moduleGlobal =
              GetRowValueIfColumnExists(
                  row,
                  ExcelSchemaNames.Columns.Module);

          if (!ModuleCorrespond(
                  moduleGlobal,
                  affectation.Module)) {
            continue;
          }

          string coursGlobal =
              GetRowValueIfColumnExists(
                  row,
                  ExcelSchemaNames.Columns.Cours);

          if (!string.Equals(
                  coursGlobal,
                  affectation.Cours,
                  StringComparison.OrdinalIgnoreCase)) {
            continue;
          }

          string libelleCourtGlobal =
              GetRowValueIfColumnExists(
                  row,
                  colonneLibelleCourt);

          if (!string.Equals(
                  libelleCourtGlobal,
                  affectation.LibelleCourt,
                  StringComparison.OrdinalIgnoreCase)) {
            continue;
          }

          lignesASupprimer.Add(
              row);

          break;
        }
      }

      foreach (DataRow row in lignesASupprimer) {
        tableGlobal.Rows.Remove(
            row);
      }

      if (lignesASupprimer.Count > 0) {
        Debug.WriteLine(
            "Lignes INFOS spéciales supprimées avant reconstruction : " +
            lignesASupprimer.Count);
      }
    }

    private static List<AffectationGroupeSemestre> LireAffectationsGroupesDepuisFeuilleSemestre_Epplus(
        string cheminFichier,
        string nomFeuille
    ) {
      List<AffectationGroupeSemestre> affectations =
          new List<AffectationGroupeSemestre>();

      if (string.IsNullOrWhiteSpace(cheminFichier))
        return affectations;

      if (!File.Exists(cheminFichier))
        return affectations;

      using (ExcelPackage package =
          new ExcelPackage(new FileInfo(cheminFichier))) {
        ExcelWorksheet? feuille =
            package.Workbook.Worksheets[nomFeuille];

        if (feuille == null || feuille.Dimension == null)
          return affectations;
        string infosCourantes = string.Empty;
        int ligneSalles;
        int colonneSalles;

        if (!TrouverCelluleTexteFeuilleSemestre(
                feuille,
                "SALLES",
                out ligneSalles,
                out colonneSalles)) {
          return affectations;
        }

        int colMax =
            TrouverDerniereColonneGroupes(
                feuille,
                ligneSalles,
                ligneSalles + 1,
                colonneSalles + 1);

        if (colMax < colonneSalles + 1)
          return affectations;

        Dictionary<int,string> groupesTd =
            LireGroupesFeuilleSemestre(
                feuille,
                ligneSalles,
                colonneSalles + 1,
                colMax,
                estTp: false);

        Dictionary<int,string> groupesTp =
            LireGroupesFeuilleSemestre(
                feuille,
                ligneSalles + 1,
                colonneSalles + 1,
                colMax,
                estTp: true);

        int rowMax =
            TrouverDerniereLigneUtileSemestre(
                feuille);

        Regex regexCodeService =
            new Regex(
                @"^(R\d+-\d{2})-(CM|TD|TP|DS)$",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        string moduleCourant =
            string.Empty;

        string codeServiceCourant =
            string.Empty;

        string coursCourant =
            string.Empty;

        for (int row = 1;row <= rowMax;row++) {
          string celluleModule =
              feuille.Cells[row,1].Text.Trim();

          string codeService =
              feuille.Cells[row,2].Text.Trim();

          Match match =
              regexCodeService.Match(codeService);

          string moduleDetecte =
              ExtraireModuleDepuisCelluleFeuilleSemestre(
                  celluleModule);

          if (!string.IsNullOrWhiteSpace(moduleDetecte)) {
            moduleCourant =
                moduleDetecte;

            infosCourantes =
                string.Empty;
          }
          else if (
              match.Success &&
              !string.IsNullOrWhiteSpace(celluleModule)
          ) {
            infosCourantes =
                celluleModule;
          }

          if (match.Success) {
            codeServiceCourant =
                codeService;

            coursCourant =
                match.Groups[2].Value.ToUpperInvariant();
          }
          else {
            if (!string.IsNullOrWhiteSpace(codeService))
              continue;
          }

          if (string.IsNullOrWhiteSpace(codeServiceCourant))
            continue;

          if (coursCourant != "TD" && coursCourant != "TP")
            continue;

          int dureeLigne =
              LireDureeDepuisLigneFeuilleSemestre(
                  feuille,
                  row,
                  colonneSalles);

          string infosLigne =
              LireInfosDepuisLigneFeuilleSemestre(
                  feuille,
                  row,
                  colonneSalles);

          Dictionary<int,string> groupes =
              coursCourant == "TD"
                  ? groupesTd
                  : groupesTp;

          foreach (KeyValuePair<int,string> groupe in groupes) {
            string nomIntervenant =
                feuille.Cells[row,groupe.Key].Text.Trim();

            if (!EstNomIntervenantValidePourGroupe(nomIntervenant))
              continue;

            AffectationGroupeSemestre affectation =
                new AffectationGroupeSemestre {
                  Semestre = nomFeuille,
                  Module = moduleCourant,
                  LibelleCourt = codeServiceCourant,
                  Cours = coursCourant,
                  Noms = nomIntervenant,
                  Groupe = groupe.Value,
                  Duree = dureeLigne,
                  Infos = infosCourantes,
                  SourceSheet = nomFeuille,
                  SourceRow = row,
                  SourceColumn = groupe.Key
                };

            affectations.Add(
                affectation);
          }
        }
      }

      return affectations;
    }

    private static bool ExisteDejaDansGlobalAvecMemeActivite(DataTable tableGlobal,string colonneLibelleCourt,AffectationGroupeSemestre affectation) {
      if (tableGlobal == null)
        return false;
      if (affectation == null)
        return false;
      foreach (DataRow row in tableGlobal.Rows) {
        string semestre = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Semestre);
        string module = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Module);
        string cours = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Cours);
        string noms = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Noms);
        string groupe = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Groupe);
        string libelleCourt = GetRowValueIfColumnExists(row,colonneLibelleCourt);
        int duree = LireEntierDepuisRow(row,ExcelSchemaNames.Columns.Duree);
        if (!string.Equals(semestre,affectation.Semestre,StringComparison.OrdinalIgnoreCase))
          continue;
        if (!ModuleCorrespond(module,affectation.Module))
          continue;
        if (!string.Equals(cours,affectation.Cours,StringComparison.OrdinalIgnoreCase))
          continue;
        if (!string.Equals(noms,affectation.Noms,StringComparison.OrdinalIgnoreCase))
          continue;
        if (!string.Equals(groupe,affectation.Groupe,StringComparison.OrdinalIgnoreCase))
          continue;
        if (!string.Equals(libelleCourt,affectation.LibelleCourt,StringComparison.OrdinalIgnoreCase))
          continue;
        if (affectation.Duree > 0 && duree != affectation.Duree)
          continue;
        if (row.Table.Columns.Contains(ExcelSchemaNames.Columns.Infos)) {
          row[ExcelSchemaNames.Columns.Infos] = affectation.Infos?.Trim() ?? string.Empty;
        }
        if (row.Table.Columns.Contains(ExcelSchemaNames.Columns.SourceSheet)) {
          row[ExcelSchemaNames.Columns.SourceSheet] =
              affectation.SourceSheet;
        }

        if (row.Table.Columns.Contains(ExcelSchemaNames.Columns.SourceRow)) {
          row[ExcelSchemaNames.Columns.SourceRow] =
              affectation.SourceRow.ToString();
        }

        if (row.Table.Columns.Contains(ExcelSchemaNames.Columns.SourceColumn)) {
          row[ExcelSchemaNames.Columns.SourceColumn] =
              affectation.SourceColumn.ToString();
        }
        return true;
      }
      return false;
    }

    private static DataRow? TrouverLigneModeleGlobal(DataTable tableGlobal,string colonneLibelleCourt,AffectationGroupeSemestre affectation) {
      foreach (DataRow row in tableGlobal.Rows) {
        string semestre = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Semestre);
        string cours = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Cours);
        string libelleCourt = GetRowValueIfColumnExists(row,colonneLibelleCourt);
        string module = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Module);

        if (!string.Equals(semestre,affectation.Semestre,StringComparison.OrdinalIgnoreCase))
          continue;
        if (!string.Equals(cours,affectation.Cours,StringComparison.OrdinalIgnoreCase))
          continue;
        if (!string.Equals(libelleCourt,affectation.LibelleCourt,StringComparison.OrdinalIgnoreCase))
          continue;
        if (
            string.Equals(module,affectation.Module,StringComparison.OrdinalIgnoreCase) ||
            string.Equals(AjouterSuffixeModuleUnSiNecessaire(module),affectation.Module,StringComparison.OrdinalIgnoreCase) ||
            string.Equals(module,AjouterSuffixeModuleUnSiNecessaire(affectation.Module),StringComparison.OrdinalIgnoreCase)
        ) {
          return row;
        }
      }
      return null;
    }

    private static int TrouverProchainIdGlobal(DataTable tableGlobal) {
      int maxId = 0;
      if (tableGlobal == null)
        return 1;
      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.ID))
        return tableGlobal.Rows.Count + 1;
      foreach (DataRow row in tableGlobal.Rows) {
        string valeur = row[ExcelSchemaNames.Columns.ID]?.ToString()?.Trim() ?? string.Empty;
        if (int.TryParse(valeur,out int id)) {
          if (id > maxId)
            maxId = id;
        }
      }
      return maxId + 1;
    }

    private static string LireInfosDepuisLigneFeuilleSemestre(ExcelWorksheet feuille,int row,int colonneSalles) {
      if (feuille == null)
        return string.Empty;

      // Une activité spéciale est indiquée en colonne A
      // sur la même ligne que le code R1-xx-TD/TP en colonne B.
      // Exemple :
      //   Col A = Soutien Promo
      //   Col B = R1-08-TP
      //
      // On ne lit plus les colonnes proches de SALLES,
      // car elles peuvent contenir des commentaires du planning
      // comme "Vend AP" qui ne doivent pas devenir INFOS.

      string valeur = feuille.Cells[row,1].Text.Trim();

      if (string.IsNullOrWhiteSpace(valeur))
        return string.Empty;
      if (valeur.Equals("0",StringComparison.OrdinalIgnoreCase))
        return string.Empty;
      if (valeur.Equals("#N/A",StringComparison.OrdinalIgnoreCase))
        return string.Empty;
      if (valeur.Equals("#REF!",StringComparison.OrdinalIgnoreCase))
        return string.Empty;
      if (valeur.StartsWith("Total",StringComparison.OrdinalIgnoreCase))
        return string.Empty;
      if (valeur.StartsWith("Nb",StringComparison.OrdinalIgnoreCase))
        return string.Empty;

      // Si la colonne A contient un vrai nom de module,
      // ce n'est pas une info spéciale.
      string moduleDetecte = ExtraireModuleDepuisCelluleFeuilleSemestre(valeur);

      if (!string.IsNullOrWhiteSpace(moduleDetecte))
        return string.Empty;
      return valeur;
    }

    private static int LireDureeDepuisLigneFeuilleSemestre(ExcelWorksheet feuille,int row,int colonneSalles) {
      if (feuille == null)
        return 0;
      int[] colonnesCandidates = { colonneSalles - 2,colonneSalles - 1,colonneSalles - 3 };
      foreach (int col in colonnesCandidates) {
        if (col <= 0)
          continue;
        string valeur = feuille.Cells[row,col].Text.Trim();
        if (string.IsNullOrWhiteSpace(valeur))
          continue;
        valeur = valeur.Replace(",",".");
        if (decimal.TryParse(valeur,System.Globalization.NumberStyles.Any,System.Globalization.CultureInfo.InvariantCulture,out decimal resultat)) {
          return Convert.ToInt32(Math.Round(resultat));
        }
      }
      return 0;
    }

    private static int TrouverDerniereColonneGroupes(ExcelWorksheet feuille,int ligneTd,int ligneTp,int colonneDebut) {
      if (feuille == null || feuille.Dimension == null)
        return colonneDebut - 1;
      int colonneLimite = Math.Min(feuille.Dimension.End.Column,colonneDebut + MaxColonnesRechercheGroupes);
      int derniereColonne = colonneDebut - 1;
      int colonnesVidesConsecutives = 0;
      bool groupeTrouve = false;

      for (int col = colonneDebut;col <= colonneLimite;col++) {
        string valeurTd = feuille.Cells[ligneTd,col].Text.Trim();
        string valeurTp = feuille.Cells[ligneTp,col].Text.Trim();
        bool celluleUtile = !string.IsNullOrWhiteSpace(valeurTd) || !string.IsNullOrWhiteSpace(valeurTp);
        if (celluleUtile) {
          groupeTrouve = true;
          derniereColonne = col;
          colonnesVidesConsecutives = 0;
          continue;
        }
        if (groupeTrouve) {
          colonnesVidesConsecutives++;
          if (colonnesVidesConsecutives >= StopColonnesVidesConsecutives)
            break;
        }
      }
      return derniereColonne;
    }

    private static int TrouverDerniereLigneUtileSemestre(ExcelWorksheet feuille) {
      if (feuille == null || feuille.Dimension == null)
        return 1;
      int ligneLimite = Math.Min(feuille.Dimension.End.Row,MaxLignesRechercheSemestre);
      int derniereLigne = 1;
      for (int row = 1;row <= ligneLimite;row++) {
        string colonneA = feuille.Cells[row,1].Text.Trim();
        string colonneB = feuille.Cells[row,2].Text.Trim();
        if (!string.IsNullOrWhiteSpace(colonneA) || !string.IsNullOrWhiteSpace(colonneB)) {
          derniereLigne = row;
        }
      }
      return derniereLigne;
    }

    private static bool TrouverCelluleTexteFeuilleSemestre(ExcelWorksheet feuille,string texteRecherche,out int ligneTrouvee,out int colonneTrouvee) {
      ligneTrouvee = 0;
      colonneTrouvee = 0;
      if (feuille == null || feuille.Dimension == null)
        return false;
      int rowMax = Math.Min(feuille.Dimension.End.Row,40);
      int colMax = feuille.Dimension.End.Column;
      for (int row = 1;row <= rowMax;row++) {
        for (int col = 1;col <= colMax;col++) {
          string valeur = LireTexteCelluleRapide(feuille,row,col);
          if (string.Equals(valeur,texteRecherche,StringComparison.OrdinalIgnoreCase)) {
            ligneTrouvee = row;
            colonneTrouvee = col;
            return true;
          }
        }
      }
      return false;
    }

    private static Dictionary<int,string> LireGroupesFeuilleSemestre(ExcelWorksheet feuille,int row,int colDebut,int colFin,bool estTp) {
      Dictionary<int,string> groupes = new Dictionary<int,string>();
      if (feuille == null)
        return groupes;
      if (row <= 0)
        return groupes;
      for (int col = colDebut;col <= colFin;col++) {
        string valeur = LireTexteCelluleRapide(feuille,row,col);
        if (string.IsNullOrWhiteSpace(valeur))
          continue;
        bool groupeValide = estTp
          ? Regex.IsMatch(valeur,@"^\d{1,2}[A-Z]$",RegexOptions.IgnoreCase)
          : Regex.IsMatch(valeur,@"^\d{1,2}$",RegexOptions.IgnoreCase);
        if (!groupeValide)
          continue;
        groupes[col] = valeur;
      }
      return groupes;
    }

    private static string ExtraireModuleDepuisCelluleFeuilleSemestre(string valeur) {
      if (string.IsNullOrWhiteSpace(valeur))
        return string.Empty;
      string texte = valeur.Trim();
      texte = texte.Replace('–','-');
      if (texte.StartsWith("Total",StringComparison.OrdinalIgnoreCase))
        return string.Empty;
      if (texte.StartsWith("Nb",StringComparison.OrdinalIgnoreCase))
        return string.Empty;
      int indexParenthese = texte.IndexOf('(');
      if (indexParenthese > 0)
        texte = texte.Substring(0,indexParenthese).Trim();
      int indexEspace = texte.IndexOf(' ');
      if (indexEspace > 0)
        texte = texte.Substring(0,indexEspace).Trim();
      texte = texte.Trim();
      if (!Regex.IsMatch(texte,@"^[A-Z]{3,}[A-Z0-9]*\d+(?:-\d+)?$",RegexOptions.IgnoreCase)) {
        return string.Empty;
      }
      return texte;
    }

    private static bool EstNomIntervenantValidePourGroupe(string valeur) {
      if (string.IsNullOrWhiteSpace(valeur))
        return false;
      string texte = valeur.Trim();
      if (texte == "0")
        return false;
      if (texte.Contains("#REF!",StringComparison.OrdinalIgnoreCase))
        return false;
      if (texte.StartsWith("BUT",StringComparison.OrdinalIgnoreCase))
        return false;
      decimal nombre;
      if (decimal.TryParse(texte.Replace(",","."),System.Globalization.NumberStyles.Any,System.Globalization.CultureInfo.InvariantCulture,out nombre)) {
        return false;
      }
      if (texte.Any(char.IsDigit))
        return false;
      if (texte.Contains("(") || texte.Contains(")"))
        return false;
      if (texte.Contains("/"))
        return false;
      return Regex.IsMatch(texte,@"^[A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý \-']+$",RegexOptions.IgnoreCase);
    }

    private static string ConstruireCleGroupeGlobal(string semestre,string module,string libelleCourt,string cours,string noms) {
      if (string.IsNullOrWhiteSpace(semestre))
        return string.Empty;
      if (string.IsNullOrWhiteSpace(module))
        return string.Empty;
      if (string.IsNullOrWhiteSpace(libelleCourt))
        return string.Empty;
      if (string.IsNullOrWhiteSpace(cours))
        return string.Empty;
      if (string.IsNullOrWhiteSpace(noms))
        return string.Empty;
      return
          semestre.Trim().ToUpperInvariant() + "|" +
          module.Trim().ToUpperInvariant() + "|" +
          libelleCourt.Trim().ToUpperInvariant() + "|" +
          cours.Trim().ToUpperInvariant() + "|" +
          noms.Trim().ToUpperInvariant();
    }

    private static bool ModuleCorrespond(string moduleGlobal,string moduleFeuille) {
      if (string.Equals(moduleGlobal,moduleFeuille,StringComparison.OrdinalIgnoreCase)) {
        return true;
      }
      string moduleGlobalSuffixe = AjouterSuffixeModuleUnSiNecessaire(moduleGlobal);
      string moduleFeuilleSuffixe = AjouterSuffixeModuleUnSiNecessaire(moduleFeuille);
      if (string.Equals(moduleGlobalSuffixe,moduleFeuille,StringComparison.OrdinalIgnoreCase)) {
        return true;
      }
      if (string.Equals(moduleGlobal,moduleFeuilleSuffixe,StringComparison.OrdinalIgnoreCase)) {
        return true;
      }
      return false;
    }

    private static string AjouterSuffixeModuleUnSiNecessaire(string module) {
      if (string.IsNullOrWhiteSpace(module))
        return string.Empty;
      string texte = module.Trim();
      if (Regex.IsMatch(texte,@"-\d+$"))
        return texte;
      return texte + "-1";
    }

    private static string LireTexteCelluleRapide(ExcelWorksheet feuille,int row,int col) {
      if (feuille == null)
        return string.Empty;
      object? valeur = feuille.Cells[row,col].Value;
      if (valeur == null)
        return string.Empty;
      return valeur.ToString()?.Trim() ?? string.Empty;
    }

    #endregion Enrichissement depuis les feuilles semestre

    #region Écriture et mise à jour du fichier de service

    private static string GetNextIdFromDataTable(DataTable dataTable,string idColumnName) {
      if (dataTable == null)
        return "1";
      if (!dataTable.Columns.Contains(idColumnName))
        return "1";
      int maxId = 0;
      foreach (DataRow row in dataTable.Rows) {
        string value = row[idColumnName]?.ToString()?.Trim() ?? string.Empty;
        if (int.TryParse(value,out int id)) {
          if (id > maxId)
            maxId = id;
        }
      }
      return (maxId + 1).ToString();
    }

    /// <summary>
    /// Ajoute une ligne dans le fichier ExcelApp de service, en fonction du module sélectionné.
    /// </summary>
    /// <param name="fichierDialogue">Fichier ExcelApp de service à modifier.</param>
    /// <param name="dataSetExcel">DataSet contenant les données actuellement chargées.</param>
    /// <param name="nomBase">Nom de la table ou feuille ExcelApp concernée.</param>
    /// <param name="module">Nom du module utilisé comme référence pour l'ajout.</param>
    /// <returns>DataSet passé en entrée, après traitement.</returns>
    internal static DataSet AjouterLigneFichierExcel(ClasseExcel fichierDialogue,DataSet dataSetExcel,string nomBase,string module,string coursAjoute,string nomAjoute) {
      if (fichierDialogue == null)
        return dataSetExcel;
      if (dataSetExcel == null)
        return dataSetExcel!;
      if (!dataSetExcel.Tables.Contains(nomBase))
        return dataSetExcel;

      // ==================================================
      // Table_complete : traitement EPPlus uniquement
      // AVANT toute ouverture OleDb
      // ==================================================
      if (nomBase == ExcelSchemaNames.Tables.NomTableGlobal) {
        DataTable? dataTable = dataSetExcel.Tables[nomBase];
        DataRow? sourceRow = null;
        if (!string.IsNullOrWhiteSpace(coursAjoute)) {
          sourceRow = FindLastRowByModuleAndCourse(dataTable!,module,coursAjoute.Trim());
        }
        sourceRow ??= FindLastRowByColumnValue(dataTable!,ExcelSchemaNames.Columns.Module,module);
        if (sourceRow == null) {
          MessageBox.Show(
              "Aucune ligne source trouvée pour le module :" +
              Environment.NewLine +
              module,
              "Ajout ligne Table Global",
              MessageBoxButton.OK,
              MessageBoxImage.Warning
          );
          return dataSetExcel;
        }
        string id = GetNextIdFromDataTable(dataTable!,ExcelSchemaNames.Columns.ID);
        int nbLignesAjoutees = AjouterLigneNomTableGlobal_Epplus(
          FichierDeService.Service,
          ExcelSchemaNames.Tables.NomTableGlobal,
          sourceRow,
          id,
          module,
          coursAjoute,
          nomAjoute
          );
        if (nbLignesAjoutees > 0) {
          AjouterLigneNomTableGlobal_DataTable(dataTable!,sourceRow,id,module,coursAjoute,nomAjoute);
        }
        return dataSetExcel;
      }
      return dataSetExcel;
    }

    private static int AjouterLigneNomTableGlobal_Epplus(ClasseExcel fichierDialogue,string nomBase,
      DataRow sourceRow,string id,string module,string coursAjoute = "",string nomAjoute = "") {
      if (fichierDialogue == null)
        return 0;
      if (sourceRow == null)
        return 0;
      if (string.IsNullOrWhiteSpace(id))
        return 0;
      string coursFinal = string.IsNullOrWhiteSpace(coursAjoute)
        ? GetRowValue(sourceRow,ExcelSchemaNames.Columns.Cours)
        : coursAjoute.Trim();
      string nomFinal = string.IsNullOrWhiteSpace(nomAjoute)
        ? GetRowValue(sourceRow,ExcelSchemaNames.Columns.Noms)
        : nomAjoute.Trim();
      string dureeFinale = GetRowValue(sourceRow,ExcelSchemaNames.Columns.Duree);
      string totalTypeFinal = GetRowValue(sourceRow,ExcelSchemaNames.Columns.TotalType);
      if (string.IsNullOrWhiteSpace(totalTypeFinal))
        totalTypeFinal = dureeFinale;
      try {
        return ClasseEpplus.AjouterLigneExcel_Epplus(
            cheminFichier: fichierDialogue.CheminFichier,
            nomFeuille: nomBase,
            actionAjout: (feuille,ligneCible) => {
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.ID,id);
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Module,module);
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Noms,nomFinal);
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Semestre,GetRowValue(sourceRow,ExcelSchemaNames.Columns.Semestre));
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Formation,GetRowValue(sourceRow,ExcelSchemaNames.Columns.Formation));
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Parcours,GetRowValue(sourceRow,ExcelSchemaNames.Columns.Parcours));
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.LibellePpn,GetRowValue(sourceRow,ExcelSchemaNames.Columns.LibellePpn));
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Libelle,GetRowValue(sourceRow,ExcelSchemaNames.Columns.Libelle));
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Commentaires,GetRowValue(sourceRow,ExcelSchemaNames.Columns.Commentaires));
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Cours,coursFinal);
              ClasseEpplus.SetIntValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Duree,dureeFinale);
              ClasseEpplus.SetIntValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Total,GetRowValue(sourceRow,ExcelSchemaNames.Columns.Total));
              ClasseEpplus.SetIntValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Nombre,"1");
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Groupe,string.Empty);
              ClasseEpplus.SetIntValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.TotalType,totalTypeFinal);
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Salle,GetRowValue(sourceRow,ExcelSchemaNames.Columns.Salle));
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.PPN,GetRowValue(sourceRow,ExcelSchemaNames.Columns.PPN));
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Apogee,GetRowValue(sourceRow,ExcelSchemaNames.Columns.Apogee));
              ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.OSE,GetRowValue(sourceRow,ExcelSchemaNames.Columns.OSE));
            });
      }
      catch (Exception ex) {
        ShowOleDbError(ex,"AjouterLigneNomTableGlobal_Epplus","Ajout EPPlus Table Global | ID=" + id + " | Module=" + module + " | Cours=" + coursFinal + " | Nom=" + nomFinal);
        return 0;
      }
    }

    private static void AjouterLigneNomTableGlobal_DataTable(DataTable table,DataRow sourceRow,string id,string module,string coursAjoute,string nomAjoute) {
      if (table == null)
        return;
      if (sourceRow == null)
        return;
      DataRow nouvelleLigne = table.NewRow();
      foreach (DataColumn colonne in table.Columns) {
        nouvelleLigne[colonne.ColumnName] = sourceRow[colonne.ColumnName];
      }
      string coursFinal = string.IsNullOrWhiteSpace(coursAjoute)
        ? GetRowValue(sourceRow,ExcelSchemaNames.Columns.Cours)
        : coursAjoute.Trim();
      string nomFinal = string.IsNullOrWhiteSpace(nomAjoute)
        ? GetRowValue(sourceRow,ExcelSchemaNames.Columns.Noms)
        : nomAjoute.Trim();
      SetRowValue(nouvelleLigne,ExcelSchemaNames.Columns.ID,id);
      SetRowValue(nouvelleLigne,ExcelSchemaNames.Columns.Module,module);
      SetRowValue(nouvelleLigne,ExcelSchemaNames.Columns.Cours,coursFinal);
      SetRowValue(nouvelleLigne,ExcelSchemaNames.Columns.Noms,nomFinal);
      SetRowValue(nouvelleLigne,ExcelSchemaNames.Columns.Nombre,1);
      table.Rows.Add(nouvelleLigne);
      table.AcceptChanges();
    }

    /// <summary>
    /// Supprime une ligne de la table ExcelApp de service en recherchant la ligne correspondante dans le DataSet,
    /// puis en supprimant la ligne ExcelApp par son ID.
    /// </summary>
    /// <param name="fichierDialogue">Fichier ExcelApp de service à modifier.</param>
    /// <param name="dataSetExcel">DataSet contenant les données actuellement chargées.</param>
    /// <param name="nomBase">Nom de la table ou feuille ExcelApp concernée.</param>
    /// <param name="module">Module de la ligne à supprimer.</param>
    /// <param name="cours">Type de cours de la ligne à supprimer.</param>
    /// <param name="noms">Nom de l'intervenant de la ligne à supprimer.</param>
    /// <param name="groupes">Nombre ou groupe associé à la ligne à supprimer.</param>
    /// <returns>DataSet passé en entrée, après traitement.</returns>
    internal static int SuppressionLigneFichierExcelParId(ClasseExcel fichierDialogue,string nomBase,string id) {
      if (string.IsNullOrWhiteSpace(id))
        return 0;
      try {
        return ClasseEpplus.SupprimeLigneExcel_Epplus(
            cheminFichier: fichierDialogue.CheminFichier,
            nomFeuille: nomBase,
            condition: (feuille,ligneCible) =>
                ClasseEpplus.GetStringByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.ID) == id);
      }
      catch (Exception ex) {
        ShowOleDbError(ex,"SuppressionLigneFichierExcelParId","Suppression EPPlus par ID : " + id);
        return 0;
      }
    }

    /// <summary>
    /// Met à jour une seule ligne Table_complete dans le fichier ExcelApp de service,
    /// en utilisant son ID.
    /// Version EPPlus : aucune requête OleDb.
    /// </summary>
    /// <param name="fichierDialogue">Fichier ExcelApp de service à modifier.</param>
    /// <param name="nomBase">Nom de la table ExcelApp concernée.</param>
    /// <param name="id">ID de la ligne à modifier.</param>
    /// <param name="noms">Nom de l'intervenant.</param>
    /// <param name="groupe">Groupe associé à la ligne.</param>
    /// <param name="nombre">Nombre de groupes.</param>
    /// <param name="cours">Type de cours.</param>
    /// <param name="duree">Durée du cours.</param>
    /// <param name="totalType">Total type recalculé.</param>
    /// <returns>Nombre de lignes modifiées.</returns>
    internal static int MiseAJourLigneNomTableGlobalParId(ClasseExcel fichierDialogue,string nomBase,string id,string noms,string groupe,int nombre,string cours,int duree,int totalType) {
      if (fichierDialogue == null)
        return 0;

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
                ).Trim() == id.Trim(),
            actionMiseAJour: (feuille,ligneCible) => {
              ClasseEpplus.SetValueByColumnName_Epplus(
                  feuille,
                  ligneCible,
                  ExcelSchemaNames.Columns.Noms,
                  noms
              );

              ClasseEpplus.SetValueByColumnName_Epplus(
                  feuille,
                  ligneCible,
                  ExcelSchemaNames.Columns.Groupe,
                  groupe
              );

              ClasseEpplus.SetValueByColumnName_Epplus(
                  feuille,
                  ligneCible,
                  ExcelSchemaNames.Columns.Nombre,
                  nombre
              );

              ClasseEpplus.SetValueByColumnName_Epplus(
                  feuille,
                  ligneCible,
                  ExcelSchemaNames.Columns.Cours,
                  cours
              );

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
        ShowOleDbError(
            ex,
            "MiseAJourLigneNomTableGlobalParId",
            "Mise à jour EPPlus par ID : " + id
        );

        return 0;
      }
    }

    /// <summary>
    /// Met à jour les informations liées à un module dans le fichier ExcelApp de service.
    /// Selon la table ciblée, met à jour les volumes CM/TD/TP ou les lignes associées dans Table_complete.
    /// </summary>
    /// <param name="fichierDialogue">Fichier ExcelApp de service à modifier.</param>
    /// <param name="dataSetExcel">DataSet contenant les données actuellement chargées.</param>
    /// <param name="nomBase">Nom de la table ou feuille ExcelApp concernée.</param>
    /// <param name="module">Nom du module à mettre à jour.</param>
    /// <returns>DataSet passé en entrée, après traitement.</returns>
    internal static DataSet MiseAJourLigneModule_fichierDeServiceExcel(ClasseExcel fichierDialogue,DataSet dataSetExcel,string nomBase,string module) {
      if (fichierDialogue == null)
        return dataSetExcel;

      if (dataSetExcel == null) {
        MessageBox.Show(
            "Le DataSet Excel n'est pas initialisé.",
            "Erreur DataSet",
            MessageBoxButton.OK,
            MessageBoxImage.Warning
        );

        return dataSetExcel!;
      }
      if (string.IsNullOrWhiteSpace(nomBase))
        return dataSetExcel;

      if (!dataSetExcel.Tables.Contains(nomBase))
        return dataSetExcel;

      DataTable dataTable = dataSetExcel.Tables[nomBase]!;

      switch (nomBase) {
        case ExcelSchemaNames.Tables.NomTableGlobal:
          MiseAJourNomTableGlobalParModule_Epplus(fichierDialogue,dataTable,nomBase,module);
          break;
        case ExcelSchemaNames.Tables.Module:
          MiseAJourHeuresModuleDepuisDataTable_Epplus(fichierDialogue,dataTable,nomBase,module);

          break;

        default:
          break;
      }

      return dataSetExcel;
    }

    private static int MiseAJourNomTableGlobalParModule_Epplus(ClasseExcel fichierDialogue,DataTable dataTable,string nomBase,string module) {
      if (fichierDialogue == null)
        return 0;
      if (dataTable == null)
        return 0;
      Dictionary<string,(string Noms,string Groupe,int Nombre,string Cours,int Duree,int TotalType)> lignesParId = new();
      foreach (DataRow ligne in dataTable.Rows) {
        if (ligne[ExcelSchemaNames.Columns.Module].ToString() != module)
          continue;
        string? id = ligne[ExcelSchemaNames.Columns.ID].ToString();
        if (string.IsNullOrWhiteSpace(id))
          continue;
        string noms = ligne[ExcelSchemaNames.Columns.Noms].ToString() ?? string.Empty;
        string groupe = ligne[ExcelSchemaNames.Columns.Groupe].ToString() ?? string.Empty;
        string cours = ligne[ExcelSchemaNames.Columns.Cours].ToString() ?? string.Empty;
        string? nombre = ligne[ExcelSchemaNames.Columns.Nombre].ToString();
        string? duree = ligne[ExcelSchemaNames.Columns.Duree].ToString();
        string? totalType = ligne[ExcelSchemaNames.Columns.TotalType].ToString();
        if (!int.TryParse(nombre,out int intNombre))
          continue;
        if (!int.TryParse(duree,out int intDuree))
          continue;
        if (!int.TryParse(totalType,out int intTotalType))
          continue;
        lignesParId[id.Trim()] = (noms,groupe,intNombre,cours,intDuree,intTotalType);
      }
      if (lignesParId.Count == 0)
        return 0;
      return ClasseEpplus.MiseAJourCelluleExcel_Epplus(
          cheminFichier: fichierDialogue.CheminFichier,
          nomFeuille: nomBase,
          condition: (feuille,ligneCible) => {
            string idExcel = ClasseEpplus.GetStringByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.ID).Trim();
            return lignesParId.ContainsKey(idExcel);
          },
          actionMiseAJour: (feuille,ligneCible) => {
            string idExcel = ClasseEpplus.GetStringByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.ID).Trim();
            if (!lignesParId.TryGetValue(idExcel,out var valeurs))
              return;
            ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Noms,valeurs.Noms);
            ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Groupe,valeurs.Groupe);
            ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Nombre,valeurs.Nombre);
            ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Cours,valeurs.Cours);
            ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.Duree,valeurs.Duree);
            ClasseEpplus.SetValueByColumnName_Epplus(feuille,ligneCible,ExcelSchemaNames.Columns.TotalType,valeurs.TotalType);
          }
      );
    }

    private static int MiseAJourHeuresModuleDepuisDataTable_Epplus(
        ClasseExcel fichierDialogue,
        DataTable dataTable,
        string nomBase,
        string module
    ) {
      foreach (DataRow ligne in dataTable.Rows) {
        if (ligne[ExcelSchemaNames.Columns.Module].ToString() != module)
          continue;

        string? semestre = ligne[ExcelSchemaNames.Columns.Semestre].ToString();
        string? cm = ligne[ExcelSchemaNames.Columns.CM].ToString();
        string? td = ligne[ExcelSchemaNames.Columns.TD].ToString();
        string? tp = ligne[ExcelSchemaNames.Columns.TP].ToString();

        if (!int.TryParse(cm,out int intCm))
          return 0;

        if (!int.TryParse(td,out int intTd))
          return 0;

        if (!int.TryParse(tp,out int intTp))
          return 0;

        int lignesModifiees =
            MiseAJourHeuresModuleParModuleEtUe_Epplus(
                fichierDialogue,
                nomBase,
                module,
                semestre!,
                intCm,
                intTd,
                intTp
            );

        return lignesModifiees;
      }

      return 0;
    }

    internal static int MiseAJourHeuresModuleParModuleEtUe_Epplus(
        ClasseExcel fichierDialogue,
        string nomBase,
        string module,
        string semestre,
        int cm,
        int td,
        int tp
    ) {
      if (fichierDialogue == null)
        return 0;

      if (string.IsNullOrWhiteSpace(module))
        return 0;

      if (string.IsNullOrWhiteSpace(semestre))
        return 0;

      try {
        return ClasseEpplus.MiseAJourCelluleExcel_Epplus(
            cheminFichier: fichierDialogue.CheminFichier,
            nomFeuille: nomBase,
            condition: (feuille,ligneCible) =>
                ClasseEpplus.GetStringByColumnName_Epplus(
                    feuille,
                    ligneCible,
                    ExcelSchemaNames.Columns.Module
                ) == module &&
                ClasseEpplus.GetStringByColumnName_Epplus(
                    feuille,
                    ligneCible,
                    ExcelSchemaNames.Columns.Semestre
                ) == semestre,
            actionMiseAJour: (feuille,ligneCible) => {
              ClasseEpplus.SetValueByColumnName_Epplus(
                  feuille,
                  ligneCible,
                  ExcelSchemaNames.Columns.CM,
                  cm
              );

              ClasseEpplus.SetValueByColumnName_Epplus(
                  feuille,
                  ligneCible,
                  ExcelSchemaNames.Columns.TD,
                  td
              );

              ClasseEpplus.SetValueByColumnName_Epplus(
                  feuille,
                  ligneCible,
                  ExcelSchemaNames.Columns.TP,
                  tp
              );
            }
        );
      }
      catch (Exception ex) {
        ShowOleDbError(
            ex,
            "MiseAJourHeuresModuleParModuleEtUe_Epplus",
            "Mise à jour EPPlus Module + UE : " + module + " / " + semestre
        );

        return 0;
      }
    }

    private static CellReferenceTarget EcrireValeurDansCelluleSourceRecursive(
    string cheminFichierExcel,
    string feuilleDepart,
    string adresseDepart,
    string nouvelleValeur
) {
      FileInfo fichier =
          new FileInfo(cheminFichierExcel);

      using ExcelPackage package =
          new ExcelPackage(fichier);

      CellReferenceTarget cible =
          ExcelReferenceResolver.ResolveFinalReferencedCell(
              package.Workbook,
              feuilleDepart,
              adresseDepart);

      if (
          !string.Equals(cible.WorksheetName,"S1",StringComparison.OrdinalIgnoreCase) &&
          !string.Equals(cible.WorksheetName,"S2",StringComparison.OrdinalIgnoreCase)
      ) {
        throw new InvalidOperationException(
            "Modification refusée : la cellule source finale n'est pas dans S1 ou S2. Cellule trouvée : " +
            cible.WorksheetName + "!" + cible.Address);
      }

      ExcelWorksheet feuilleCible =
          package.Workbook.Worksheets[cible.WorksheetName];

      feuilleCible.Cells[cible.Address].Value =
          nouvelleValeur;

      package.Workbook.Calculate();

      package.Save();

      return cible;
    }

    private sealed class CelluleSourceExcel {
      internal string NomFeuille { get; set; } =
          string.Empty;

      internal string Adresse { get; set; } =
          string.Empty;

      internal List<string> Chaine { get; } =
          new();
    }

    private static CelluleSourceExcel ResoudreCelluleSourceRecursive(
        ExcelWorkbook workbook,
        string nomFeuilleDepart,
        string adresseDepart
    ) {
      string nomFeuille =
          nomFeuilleDepart.Trim();

      string adresse =
          adresseDepart.Replace("$",string.Empty).Trim();

      CelluleSourceExcel resultat =
          new();

      HashSet<string> cellulesVisitees =
          new(StringComparer.OrdinalIgnoreCase);

      for (int profondeur = 0;profondeur < 30;profondeur++) {
        string cle =
            nomFeuille + "!" + adresse;

        if (!cellulesVisitees.Add(cle)) {
          throw new InvalidOperationException(
              "Référence circulaire détectée : " + cle);
        }

        resultat.Chaine.Add(
            cle);

        ExcelWorksheet feuille =
            TrouverFeuilleExcel(
                workbook,
                nomFeuille);

        if (feuille == null) {
          throw new InvalidOperationException(
              "Feuille introuvable : " + nomFeuille);
        }

        ExcelRange cellule =
            feuille.Cells[adresse];

        string formule =
            cellule.Formula?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(formule)) {
          resultat.NomFeuille =
              feuille.Name;

          resultat.Adresse =
              adresse;

          return resultat;
        }

        if (
            !EssayerLireReferenceCelluleDirecte(
                formule,
                feuille.Name,
                out string prochaineFeuille,
                out string prochaineAdresse)
        ) {
          throw new InvalidOperationException(
              "La formule n'est pas une référence directe : " +
              cle +
              " = " +
              formule);
        }

        nomFeuille =
            prochaineFeuille;

        adresse =
            prochaineAdresse;
      }

      throw new InvalidOperationException(
          "Trop de niveaux de références Excel.");
    }

    private static bool EssayerLireReferenceCelluleDirecte(
        string formule,
        string nomFeuilleCourante,
        out string nomFeuille,
        out string adresse
    ) {
      nomFeuille =
          nomFeuilleCourante;

      adresse =
          string.Empty;

      if (string.IsNullOrWhiteSpace(formule))
        return false;

      string texte =
          formule.Trim();

      if (texte.StartsWith("=",StringComparison.Ordinal))
        texte =
            texte.Substring(1).Trim();

      texte =
          texte.Replace("$",string.Empty);

      Match match =
          Regex.Match(
              texte,
              @"^(?:(?:'(?<sheetq>[^']+)'|(?<sheet>[^'!]+))!)?(?<col>[A-Z]{1,3})(?<row>[0-9]+)$",
              RegexOptions.IgnoreCase);

      if (!match.Success)
        return false;

      if (match.Groups["sheetq"].Success) {
        nomFeuille =
            match.Groups["sheetq"].Value.Trim();
      }
      else if (match.Groups["sheet"].Success) {
        nomFeuille =
            match.Groups["sheet"].Value.Trim();
      }

      adresse =
          match.Groups["col"].Value.ToUpperInvariant() +
          match.Groups["row"].Value;

      return true;
    }

    private static ExcelWorksheet TrouverFeuilleExcel(
        ExcelWorkbook workbook,
        string nomFeuille
    ) {
      foreach (ExcelWorksheet feuille in workbook.Worksheets) {
        if (
            string.Equals(
                feuille.Name,
                nomFeuille,
                StringComparison.OrdinalIgnoreCase)
        ) {
          return feuille;
        }
      }

      return null!;
    }

    private static int TrouverLigneEnteteExcel(
        ExcelWorksheet feuille,
        string nomColonneReference
    ) {
      if (feuille.Dimension == null)
        return -1;

      int maxRow =
          Math.Min(
              feuille.Dimension.End.Row,
              30);

      for (int row = feuille.Dimension.Start.Row;row <= maxRow;row++) {
        for (int col = feuille.Dimension.Start.Column;col <= feuille.Dimension.End.Column;col++) {
          string texte =
              feuille.Cells[row,col].Text.Trim();

          if (
              string.Equals(
                  texte,
                  nomColonneReference,
                  StringComparison.OrdinalIgnoreCase)
          ) {
            return row;
          }
        }
      }

      return -1;
    }

    private static int TrouverIndexColonneExcel(
        ExcelWorksheet feuille,
        int ligneEntete,
        string nomColonne
    ) {
      if (feuille.Dimension == null)
        return -1;

      for (int col = feuille.Dimension.Start.Column;col <= feuille.Dimension.End.Column;col++) {
        string texte =
            feuille.Cells[ligneEntete,col].Text.Trim();

        if (
            string.Equals(
                texte,
                nomColonne,
                StringComparison.OrdinalIgnoreCase)
        ) {
          return col;
        }
      }

      return -1;
    }

    private static int TrouverLigneGlobalParId(
        ExcelWorksheet feuille,
        int ligneEntete,
        int colonneId,
        string id
    ) {
      if (feuille.Dimension == null)
        return -1;

      if (colonneId <= 0)
        return -1;

      for (int row = ligneEntete + 1;row <= feuille.Dimension.End.Row;row++) {
        string idLigne =
            feuille.Cells[row,colonneId].Text.Trim();

        if (
            string.Equals(
                idLigne,
                id,
                StringComparison.OrdinalIgnoreCase)
        ) {
          return row;
        }
      }

      return -1;
    }

    private static string LireValeurDepuisLigneVue(DataRowView row,string colonne) {
      return GetBoundRowValue_vue(row,colonne).Trim();
    }

    internal static int MiseAJourCellulesSourcesReferenceesNomTableGlobal_Epplus(
    ClasseExcel fichierDialogue,
    DataRowView row,
    params string[] colonnes
) {
      if (fichierDialogue == null)
        return 0;

      if (row == null)
        return 0;

      if (string.IsNullOrWhiteSpace(fichierDialogue.CheminFichier))
        return 0;

      if (colonnes == null || colonnes.Length == 0)
        return 0;

      string id =
          GetBoundRowValue_vue(
              row,
              ExcelSchemaNames.Columns.ID).Trim();

      if (string.IsNullOrWhiteSpace(id)) {
        MessageBox.Show(
            "Impossible de mettre à jour Excel : l'ID de la ligne est absent.",
            "Mise à jour source Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      FileInfo fichier =
          new(fichierDialogue.CheminFichier);

      if (!fichier.Exists) {
        MessageBox.Show(
            "Le fichier Excel de service est introuvable :" +
            Environment.NewLine +
            fichierDialogue.CheminFichier,
            "Mise à jour source Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      int cellulesModifiees =
          0;

      using ExcelPackage package =
          new(fichier);

      ExcelWorksheet? feuilleGlobal =
          TrouverFeuille(
              package.Workbook,
              ExcelSchemaNames.Tables.NomTableGlobal);

      if (feuilleGlobal == null) {
        MessageBox.Show(
            "La feuille Global est introuvable dans le fichier de service.",
            "Mise à jour source Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      int ligneEntete =
          TrouverLigneEntete(
              feuilleGlobal,
              ExcelSchemaNames.Columns.ID);

      if (ligneEntete <= 0) {
        MessageBox.Show(
            "La ligne d'entête de Global est introuvable.",
            "Mise à jour source Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      int colonneId =
          TrouverColonne(
              feuilleGlobal,
              ligneEntete,
              ExcelSchemaNames.Columns.ID);

      if (colonneId <= 0) {
        MessageBox.Show(
            "La colonne ID est introuvable dans Global.",
            "Mise à jour source Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      int ligneGlobal =
          TrouverLigneParId(
              feuilleGlobal,
              ligneEntete,
              colonneId,
              id);

      if (ligneGlobal <= 0) {
        MessageBox.Show(
            "Impossible de retrouver la ligne Global correspondant à l'ID : " + id,
            "Mise à jour source Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      foreach (string colonne in colonnes) {
        if (string.IsNullOrWhiteSpace(colonne))
          continue;

        int colonneGlobal =
            TrouverColonne(
                feuilleGlobal,
                ligneEntete,
                colonne);

        if (colonneGlobal <= 0)
          continue;

        string nouvelleValeur =
            LireValeurDepuisLigneVue(
                row,
                colonne);

        ExcelRange celluleGlobal =
            feuilleGlobal.Cells[ligneGlobal,colonneGlobal];

        CelluleSourceReference celluleSource =
            ResoudreCelluleSource(
                package.Workbook,
                feuilleGlobal.Name,
                celluleGlobal.Address);

        if (
            !string.Equals(celluleSource.NomFeuille,"S1",StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(celluleSource.NomFeuille,"S2",StringComparison.OrdinalIgnoreCase)
        ) {
          MessageBox.Show(
              "Modification refusée." +
              Environment.NewLine +
              Environment.NewLine +
              "La cellule source finale n'est pas dans S1 ou S2." +
              Environment.NewLine +
              "Cellule trouvée : " + celluleSource.NomFeuille + "!" + celluleSource.Adresse +
              Environment.NewLine +
              Environment.NewLine +
              "Chaîne de références :" +
              Environment.NewLine +
              string.Join(" -> ",celluleSource.Chaine),
              "Mise à jour source Excel",
              MessageBoxButton.OK,
              MessageBoxImage.Warning);

          continue;
        }

        ExcelWorksheet? feuilleSource =
            TrouverFeuille(
                package.Workbook,
                celluleSource.NomFeuille);

        if (feuilleSource == null)
          continue;

        feuilleSource.Cells[celluleSource.Adresse].Value =
            nouvelleValeur;

        cellulesModifiees++;
      }

      if (cellulesModifiees > 0) {
        try {
          package.Workbook.Calculate();
        }
        catch {
          // EPPlus ne recalcule pas toujours toutes les formules Excel.
          // La cellule source est quand même bien modifiée.
        }

        package.Save();
      }

      return cellulesModifiees;

      static ExcelWorksheet? TrouverFeuille(
          ExcelWorkbook workbook,
          string nomFeuille
      ) {
        foreach (ExcelWorksheet feuille in workbook.Worksheets) {
          if (
              string.Equals(
                  feuille.Name,
                  nomFeuille,
                  StringComparison.OrdinalIgnoreCase)
          ) {
            return feuille;
          }
        }

        return null;
      }

      static int TrouverLigneEntete(
          ExcelWorksheet feuille,
          string nomColonneReference
      ) {
        if (feuille == null)
          return -1;

        if (feuille.Dimension == null)
          return -1;

        // 1) Priorité aux tables Excel structurées
        foreach (var table in feuille.Tables) {
          foreach (var colonne in table.Columns) {
            if (
                string.Equals(
                    colonne.Name?.Trim(),
                    nomColonneReference,
                    StringComparison.OrdinalIgnoreCase)
            ) {
              return table.Address.Start.Row;
            }
          }
        }

        // 2) Sinon recherche dans toute la zone utilisée de la feuille
        for (int ligne = feuille.Dimension.Start.Row;ligne <= feuille.Dimension.End.Row;ligne++) {
          for (int colonne = feuille.Dimension.Start.Column;colonne <= feuille.Dimension.End.Column;colonne++) {
            string texte =
                feuille.Cells[ligne,colonne].Text.Trim();

            if (
                string.Equals(
                    texte,
                    nomColonneReference,
                    StringComparison.OrdinalIgnoreCase)
            ) {
              return ligne;
            }
          }
        }

        return -1;
      }

      static int TrouverColonne(
          ExcelWorksheet feuille,
          int ligneEntete,
          string nomColonne
      ) {
        if (feuille == null)
          return -1;

        if (feuille.Dimension == null)
          return -1;

        // 1) Priorité aux tables Excel structurées
        foreach (var table in feuille.Tables) {
          foreach (var colonne in table.Columns) {
            if (
                string.Equals(
                    colonne.Name?.Trim(),
                    nomColonne,
                    StringComparison.OrdinalIgnoreCase)
            ) {
              return table.Address.Start.Column + colonne.Position;
            }
          }
        }

        // 2) Sinon recherche classique sur la ligne d'entête trouvée
        for (int colonne = feuille.Dimension.Start.Column;colonne <= feuille.Dimension.End.Column;colonne++) {
          string texte =
              feuille.Cells[ligneEntete,colonne].Text.Trim();

          if (
              string.Equals(
                  texte,
                  nomColonne,
                  StringComparison.OrdinalIgnoreCase)
          ) {
            return colonne;
          }
        }

        return -1;
      }
      static int TrouverLigneParId(
          ExcelWorksheet feuille,
          int ligneEntete,
          int colonneId,
          string id
      ) {
        if (feuille.Dimension == null)
          return -1;

        for (int ligne = ligneEntete + 1;ligne <= feuille.Dimension.End.Row;ligne++) {
          string idLigne =
              feuille.Cells[ligne,colonneId].Text.Trim();

          if (
              string.Equals(
                  idLigne,
                  id,
                  StringComparison.OrdinalIgnoreCase)
          ) {
            return ligne;
          }
        }

        return -1;
      }

      static string LireValeurDepuisLigneVue(DataRowView row,string colonne) {
        return GetBoundRowValue_vue(row,colonne).Trim();
      }

      static CelluleSourceReference ResoudreCelluleSource(
          ExcelWorkbook workbook,
          string nomFeuilleDepart,
          string adresseDepart
      ) {
        string nomFeuille =
            nomFeuilleDepart.Trim();

        string adresse =
            adresseDepart.Replace("$",string.Empty).Trim();

        CelluleSourceReference resultat =
            new();

        HashSet<string> cellulesVisitees =
            new(StringComparer.OrdinalIgnoreCase);

        for (int profondeur = 0;profondeur < 30;profondeur++) {
          string cle =
              nomFeuille + "!" + adresse;

          if (!cellulesVisitees.Add(cle)) {
            throw new InvalidOperationException(
                "Référence circulaire détectée : " + cle);
          }

          resultat.Chaine.Add(
              cle);

          ExcelWorksheet? feuille =
              TrouverFeuille(
                  workbook,
                  nomFeuille);

          if (feuille == null) {
            throw new InvalidOperationException(
                "Feuille introuvable : " + nomFeuille);
          }

          ExcelRange cellule =
              feuille.Cells[adresse];

          string formule =
              cellule.Formula?.Trim() ?? string.Empty;

          if (string.IsNullOrWhiteSpace(formule)) {
            resultat.NomFeuille =
                feuille.Name;

            resultat.Adresse =
                adresse;

            return resultat;
          }

          if (
              !EssayerLireReferenceDirecte(
                  formule,
                  feuille.Name,
                  out string prochaineFeuille,
                  out string prochaineAdresse)
          ) {
            throw new InvalidOperationException(
                "La formule n'est pas une référence directe : " +
                cle +
                " = " +
                formule);
          }

          nomFeuille =
              prochaineFeuille;

          adresse =
              prochaineAdresse;
        }

        throw new InvalidOperationException(
            "Trop de niveaux de références Excel.");
      }

      static bool EssayerLireReferenceDirecte(
          string formule,
          string nomFeuilleCourante,
          out string nomFeuille,
          out string adresse
      ) {
        nomFeuille =
            nomFeuilleCourante;

        adresse =
            string.Empty;

        if (string.IsNullOrWhiteSpace(formule))
          return false;

        string texte =
            formule.Trim();

        if (texte.StartsWith("=",StringComparison.Ordinal))
          texte =
              texte.Substring(1).Trim();

        texte =
            texte.Replace("$",string.Empty);

        Match match =
            Regex.Match(
                texte,
                @"^(?:(?:'(?<sheetq>[^']+)'|(?<sheet>[^'!]+))!)?(?<col>[A-Z]{1,3})(?<row>[0-9]+)$",
                RegexOptions.IgnoreCase);

        if (!match.Success)
          return false;

        if (match.Groups["sheetq"].Success) {
          nomFeuille =
              match.Groups["sheetq"].Value.Trim();
        }
        else if (match.Groups["sheet"].Success) {
          nomFeuille =
              match.Groups["sheet"].Value.Trim();
        }

        adresse =
            match.Groups["col"].Value.ToUpperInvariant() +
            match.Groups["row"].Value;

        return true;
      }
    }

    private sealed class CelluleSourceReference {
      internal string NomFeuille { get; set; } =
          string.Empty;

      internal string Adresse { get; set; } =
          string.Empty;

      internal List<string> Chaine { get; } =
          new();
    }


    /// <summary>
    /// Modifie dans une table la valeur d'une colonne pour la première ligne
    /// contenant le texte du module indiqué.
    /// </summary>
    /// <param name="texteMatière">Label contenant le texte du module recherché.</param>
    /// <param name="entêteColonne">Nom de la colonne à modifier.</param>
    /// <param name="tableDeDonnée">Nom de la table à modifier dans le DataSet de service.</param>
    /// <param name="nouvelleValeur">Nouvelle valeur à appliquer.</param>
    /// <returns>Table modifiée.</returns>
    internal static int MiseAJourCellulesCourantesNomTableGlobal_Epplus(params string[] colonnes) {
      if (DialogueServiceGEII.VisualisationDonnées.SelectedItem is not DataRowView row)
        return 0;

      string id =
          ClasseBaseDeDonnées.GetBoundRowValue_vue(
              row,
              ExcelSchemaNames.Columns.ID);

      if (string.IsNullOrWhiteSpace(id)) {
        MessageBox.Show(
            DialogueServiceGEII,
            "Impossible de mettre à jour Excel : l'ID de la ligne est absent.",
            "Mise à jour Excel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return 0;
      }

      return ClasseEpplus.MiseAJourCelluleExcel_Epplus(
          cheminFichier: FichierDeService.Service.CheminFichier,
          nomFeuille: ExcelSchemaNames.Tables.NomTableGlobal,
          condition: (feuille,ligne) =>
              ClasseEpplus.GetStringByColumnName_Epplus(
                  feuille,
                  ligne,
                  ExcelSchemaNames.Columns.ID) == id,
          actionMiseAJour: (feuille,ligne) => {
            foreach (string colonne in colonnes) {
              string valeur =
                  ClasseBaseDeDonnées.GetBoundRowValue_vue(
                      row,
                      colonne);

              ClasseEpplus.SetValueByColumnName_Epplus(
                  feuille,
                  ligne,
                  colonne,
                  valeur);
            }
          });
    }

    /// <summary>
    /// Met à jour dans la table Module les valeurs CM, TD et TP d'une fiche module.
    /// </summary>
    /// <param name="texteMatière">Label contenant le module recherché.</param>
    /// <param name="nouvelleValeurCM">Nouvelle valeur CM.</param>
    /// <param name="nouvelleValeurTD">Nouvelle valeur TD.</param>
    /// <param name="nouvelleValeurTP">Nouvelle valeur TP.</param>
    /// <returns>Table Module modifiée.</returns>
    internal protected static DataTable ChangementBaseDeDonnéesFicheModule(
        Label texteMatière,
        TextBox nouvelleValeurCM,
        TextBox nouvelleValeurTD,
        TextBox nouvelleValeurTP
    ) {
      DataTable? tableDeDonnées =
          DataSetExcel.Tables[ExcelSchemaNames.Tables.Module];

      if (tableDeDonnées == null)
        return null!;

      string module =
          texteMatière.Content?.ToString()?.Trim() ?? string.Empty;

      if (string.IsNullOrWhiteSpace(module))
        return tableDeDonnées;

      tableDeDonnées.BeginLoadData();

      try {
        foreach (DataRow ligne in tableDeDonnées.Rows) {
          string moduleLigne =
              ligne[ExcelSchemaNames.Columns.Module]?.ToString()?.Trim()
              ?? string.Empty;

          if (!string.Equals(moduleLigne,module,StringComparison.OrdinalIgnoreCase))
            continue;

          ligne[ExcelSchemaNames.Columns.CM] =
              nouvelleValeurCM.Text.Trim();

          ligne[ExcelSchemaNames.Columns.TD] =
              nouvelleValeurTD.Text.Trim();

          ligne[ExcelSchemaNames.Columns.TP] =
              nouvelleValeurTP.Text.Trim();

          break;
        }

        tableDeDonnées.AcceptChanges();
      }
      finally {
        tableDeDonnées.EndLoadData();
      }

      return tableDeDonnées;
    }

    #endregion Écriture et mise à jour du fichier de service

    #region Lecture, filtrage et affichage dans l'interface

    /// <summary>
    /// Construit un filtre à partir des ComboBox renseignées puis applique ce filtre
    /// à la table demandée pour alimenter le DataGrid.
    /// </summary>
    /// <param name="table">DataGrid à alimenter.</param>
    /// <param name="liste1">Première ComboBox utilisée pour le filtre.</param>
    /// <param name="liste2">Deuxième ComboBox optionnelle utilisée pour le filtre.</param>
    /// <param name="liste3">Troisième ComboBox optionnelle utilisée pour le filtre.</param>
    /// <param name="liste4">Quatrième ComboBox optionnelle utilisée pour le filtre.</param>
    /// <param name="liste5">Cinquième ComboBox optionnelle utilisée pour le filtre.</param>
    /// <param name="liste6">Sixième ComboBox optionnelle utilisée pour le filtre.</param>
    /// <param name="tableDeDonnée">Nom de la table DataSet à filtrer. Par défaut : Table_complete.</param>
    public static void LireDonnées(
        DataGrid table,
        ComboBox liste1,
        ComboBox liste2 = null!,
        ComboBox liste3 = null!,
        ComboBox liste4 = null!,
        ComboBox liste5 = null!,
        ComboBox liste6 = null!,
        string tableDeDonnée = "") {

      if (string.IsNullOrWhiteSpace(tableDeDonnée))
        tableDeDonnée = ExcelSchemaNames.Tables.NomTableGlobal;

      DataTable? tableSource = DataSetExcel.Tables.Contains(tableDeDonnée) ? DataSetExcel.Tables[tableDeDonnée] : null;

      if (tableSource == null)
        return;

      ComboBox[] listes = [liste1,liste2,liste3,liste4,liste5,liste6];

      bool filtreInfosActif = listes.Any(liste =>
          liste != null &&
          string.Equals(liste.Name,ExcelSchemaNames.Columns.Infos,StringComparison.OrdinalIgnoreCase) &&
          !string.IsNullOrWhiteSpace(liste.Text));

      List<string> filtres = [];

      if (tableDeDonnée == ExcelSchemaNames.Tables.NomTableGlobal &&
          !filtreInfosActif &&
          tableSource.Columns.Contains(ExcelSchemaNames.Columns.Infos)) {
        filtres.Add("([INFOS] IS NULL OR [INFOS] = '' OR [INFOS] = '0')");
      }

      foreach (ComboBox liste in listes) {
        if (liste == null || string.IsNullOrWhiteSpace(liste.Text))
          continue;

        string nomColonne = TrouverNomColonneIgnoreCase(tableSource,liste.Name);

        if (string.IsNullOrWhiteSpace(nomColonne))
          continue;

        filtres.Add(ExcelQueryBuilder.BuildEqualsFilter(nomColonne,liste.Text));
      }

      string filtreListe = string.Join(" AND ",filtres);

      FiltreAffichage(table,tableDeDonnée,filtreListe);
    }

    public static void LireDonnées(DataGrid table,string tableDeDonnée) {
      LireDonnées(table,null!,null!,null!,null!,null!,null!,tableDeDonnée);
    }

    public static void LireDonnées(DataGrid table,ComboBox liste1,string tableDeDonnée) {
      LireDonnées(table,liste1,null!,null!,null!,null!,null!,tableDeDonnée);
    }

    public static void LireDonnées(DataGrid table,ComboBox liste1,ComboBox liste2,string tableDeDonnée) {
      LireDonnées(table,liste1,liste2,null!,null!,null!,null!,tableDeDonnée);
    }

    public static void LireDonnées(DataGrid table,ComboBox liste1,ComboBox liste2,ComboBox liste3,string tableDeDonnée) {
      LireDonnées(table,liste1,liste2,liste3,null!,null!,null!,tableDeDonnée);
    }

    public static void LireDonnées(DataGrid table,ComboBox liste1,ComboBox liste2,ComboBox liste3,ComboBox liste4,string tableDeDonnée) {
      LireDonnées(table,liste1,liste2,liste3,liste4,null!,null!,tableDeDonnée);
    }

    public static void LireDonnées(DataGrid table,ComboBox liste1,ComboBox liste2,ComboBox liste3,ComboBox liste4,ComboBox liste5,string tableDeDonnée) {
      LireDonnées(table,liste1,liste2,liste3,liste4,liste5,null!,tableDeDonnée);
    }

    /// <summary>
    /// Applique un filtre DataView sur une table du DataSet et affecte le résultat
    /// au DataGrid fourni.
    /// </summary>
    /// <param name="table">DataGrid à alimenter.</param>
    /// <param name="tableDeDonnée">Nom de la table DataSet à utiliser.</param>
    /// <param name="filtreListe">Expression de filtre DataView à appliquer.</param>
    private static void FiltreAffichage(DataGrid table,string tableDeDonnée,string filtreListe) {
      if (table == null) return;
      if (string.IsNullOrWhiteSpace(tableDeDonnée)) return;
      if (DataSetExcel == null) return;
      if (!DataSetExcel.Tables.Contains(tableDeDonnée)) return;

      DataTable? tableSource = DataSetExcel.Tables[tableDeDonnée];
      if (tableSource == null) return;

      if (tableDeDonnée == ExcelSchemaNames.Tables.NomTableGlobal)
        filtreListe = AjouterFiltreMetierGlobal(filtreListe,tableSource);

      DataView vue = new(tableSource);

      if (!string.IsNullOrWhiteSpace(filtreListe)) {
        try {
          vue.RowFilter = filtreListe;
        }
        catch (Exception ex) {
          ShowFilterError(ex,tableDeDonnée,filtreListe,tableSource);
          return;
        }
      }

      if (tableSource.Columns.Contains(ExcelSchemaNames.Columns.Cours)) {
        try {
          vue.Sort = ExcelSchemaNames.Columns.Cours;
        }
        catch (Exception ex) {
          ShowFilterError(ex,tableDeDonnée,"Sort = " + ExcelSchemaNames.Columns.Cours,tableSource);
          return;
        }
      }

      table.ItemsSource = vue;

      if (vue.Count > 0)
        table.SelectedIndex = 0;
      else
        table.SelectedIndex = -1;

      CompterLesTypesDeCours(table);
    }

    internal static int AppliquerModificationsCellulesSourcesSemestre_Epplus(
    ClasseExcel fichierDialogue,
    IEnumerable<ExcelSourceCellChange> modifications
) {
      ClasseEpplus.ConfigureEpplusLicense();

      if (fichierDialogue == null)
        return 0;

      if (string.IsNullOrWhiteSpace(fichierDialogue.CheminFichier))
        return 0;

      if (modifications == null)
        return 0;

      FileInfo fichier =
          new(fichierDialogue.CheminFichier);

      if (!fichier.Exists)
        return 0;

      Dictionary<string,ExcelSourceCellChange> modificationsUniques =
          new(StringComparer.OrdinalIgnoreCase);

      foreach (ExcelSourceCellChange modification in modifications) {
        if (modification == null)
          continue;

        if (string.IsNullOrWhiteSpace(modification.SourceSheet))
          continue;

        if (
            !string.Equals(modification.SourceSheet,"S1",StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(modification.SourceSheet,"S2",StringComparison.OrdinalIgnoreCase)
        ) {
          continue;
        }

        if (modification.SourceRow <= 0)
          continue;

        if (modification.SourceColumn <= 0)
          continue;

        string cle =
            modification.SourceSheet.Trim() +
            "!" +
            modification.SourceRow +
            ":" +
            modification.SourceColumn;

        modificationsUniques[cle] =
            modification;
      }

      if (modificationsUniques.Count == 0)
        return 0;

      int cellulesModifiees =
          0;

      using ExcelPackage package =
          new(fichier);

      foreach (ExcelSourceCellChange modification in modificationsUniques.Values) {
        ExcelWorksheet? feuille =
            package.Workbook.Worksheets[modification.SourceSheet];

        if (feuille == null)
          continue;

        feuille.Cells[modification.SourceRow,modification.SourceColumn].Value =
            modification.NewValue;

        cellulesModifiees++;
      }

      if (cellulesModifiees > 0) {
        try {
          package.Workbook.Calculate();
        }
        catch {
          // EPPlus ne sait pas toujours recalculer toutes les formules Excel.
          // Les cellules sources sont quand même modifiées.
        }

        package.Save();
      }

      return cellulesModifiees;
    }

    private static string AjouterFiltreMetierGlobal(
        string filtreExistant,
        DataTable tableSource
    ) {
      List<string> filtres = [];

      if (!string.IsNullOrWhiteSpace(filtreExistant))
        filtres.Add("(" + filtreExistant + ")");

      if (tableSource.Columns.Contains(ExcelSchemaNames.Columns.Noms)) {
        filtres.Add("[" + ExcelSchemaNames.Columns.Noms + "] <> ''");
        filtres.Add("[" + ExcelSchemaNames.Columns.Noms + "] <> '0'");
        filtres.Add("NOT ([" + ExcelSchemaNames.Columns.Noms + "] LIKE 'BUT*')");
        filtres.Add("NOT ([" + ExcelSchemaNames.Columns.Noms + "] LIKE '*#REF!*')");
      }

      if (tableSource.Columns.Contains(ExcelSchemaNames.Columns.Cours)) {
        filtres.Add("[" + ExcelSchemaNames.Columns.Cours + "] <> 'DS'");
      }

      return string.Join(" AND ",filtres);
    }

    internal static int ReecrireLignesSourcesGroupesSemestre_Epplus(
    ClasseExcel fichierDialogue,
    DataTable tableGlobal,
    IEnumerable<ExcelSourceRowRewrite> reecritures
) {
      if (fichierDialogue == null)
        return 0;

      if (tableGlobal == null)
        return 0;

      if (reecritures == null)
        return 0;

      if (string.IsNullOrWhiteSpace(fichierDialogue.CheminFichier))
        return 0;

      FileInfo fichier =
          new(fichierDialogue.CheminFichier);

      if (!fichier.Exists)
        return 0;

      List<ExcelSourceRowRewrite> reecrituresValides =
          reecritures
              .Where(r =>
                  r != null &&
                  !string.IsNullOrWhiteSpace(r.SourceSheet) &&
                  r.SourceRow > 0 &&
                  (
                      string.Equals(r.Cours,"TD",StringComparison.OrdinalIgnoreCase) ||
                      string.Equals(r.Cours,"TP",StringComparison.OrdinalIgnoreCase)
                  ) &&
                  (
                      string.Equals(r.SourceSheet,"S1",StringComparison.OrdinalIgnoreCase) ||
                      string.Equals(r.SourceSheet,"S2",StringComparison.OrdinalIgnoreCase)
                  ))
              .GroupBy(
                  r => r.SourceSheet.Trim() + "!" + r.SourceRow + "!" + r.Cours.ToUpperInvariant(),
                  StringComparer.OrdinalIgnoreCase)
              .Select(g => g.Last())
              .ToList();

      if (reecrituresValides.Count == 0)
        return 0;

      int lignesReecrites =
          0;

      using ExcelPackage package =
          new(fichier);

      foreach (ExcelSourceRowRewrite reecriture in reecrituresValides) {
        ExcelWorksheet? feuille =
            package.Workbook.Worksheets[reecriture.SourceSheet];

        if (feuille == null || feuille.Dimension == null)
          continue;

        if (
            !TrouverCelluleTexteFeuilleSemestre(
                feuille,
                "SALLES",
                out int ligneSalles,
                out int colonneSalles)
        ) {
          continue;
        }

        int colonneMax =
            TrouverDerniereColonneGroupes(
                feuille,
                ligneSalles,
                ligneSalles + 1,
                colonneSalles + 1);

        if (colonneMax < colonneSalles + 1)
          continue;

        Dictionary<int,string> groupes =
            string.Equals(reecriture.Cours,"TD",StringComparison.OrdinalIgnoreCase)
                ? LireGroupesFeuilleSemestre(
                    feuille,
                    ligneSalles,
                    colonneSalles + 1,
                    colonneMax,
                    estTp: false)
                : LireGroupesFeuilleSemestre(
                    feuille,
                    ligneSalles + 1,
                    colonneSalles + 1,
                    colonneMax,
                    estTp: true);

        if (groupes.Count == 0)
          continue;

        Dictionary<string,int> colonneParGroupe =
            new(StringComparer.OrdinalIgnoreCase);

        foreach (KeyValuePair<int,string> groupe in groupes) {
          string nomGroupe =
              groupe.Value?.Trim() ?? string.Empty;

          if (string.IsNullOrWhiteSpace(nomGroupe))
            continue;

          if (!colonneParGroupe.ContainsKey(nomGroupe)) {
            colonneParGroupe.Add(
                nomGroupe,
                groupe.Key);
          }
        }

        List<DataRow> lignesData =
            tableGlobal.Rows
                .Cast<DataRow>()
                .Where(row =>
                    string.Equals(
                        LireValeurRow(row,ExcelSchemaNames.Columns.SourceSheet),
                        reecriture.SourceSheet,
                        StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(
                        LireValeurRow(row,ExcelSchemaNames.Columns.SourceRow),
                        reecriture.SourceRow.ToString(),
                        StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(
                        LireValeurRow(row,ExcelSchemaNames.Columns.Cours),
                        reecriture.Cours,
                        StringComparison.OrdinalIgnoreCase))
                .ToList();

        if (lignesData.Count == 0)
          continue;

        Dictionary<int,string> nomsParColonne =
            new();

        bool erreur =
            false;

        foreach (DataRow row in lignesData) {
          string nom =
              LireValeurRow(
                  row,
                  ExcelSchemaNames.Columns.Noms);

          string groupe =
              LireValeurRow(
                  row,
                  ExcelSchemaNames.Columns.Groupe);

          if (string.IsNullOrWhiteSpace(nom))
            continue;

          if (string.IsNullOrWhiteSpace(groupe)) {
            MessageBox.Show(
                "Impossible de réécrire la ligne source : un groupe est vide." +
                Environment.NewLine +
                Environment.NewLine +
                "Feuille : " + reecriture.SourceSheet +
                Environment.NewLine +
                "Ligne : " + reecriture.SourceRow +
                Environment.NewLine +
                "Intervenant : " + nom,
                "Mise à jour source",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

            erreur =
                true;

            break;
          }

          if (!colonneParGroupe.TryGetValue(groupe,out int colonneCible)) {
            MessageBox.Show(
                "Impossible de réécrire la ligne source : groupe introuvable dans la feuille semestre." +
                Environment.NewLine +
                Environment.NewLine +
                "Feuille : " + reecriture.SourceSheet +
                Environment.NewLine +
                "Ligne : " + reecriture.SourceRow +
                Environment.NewLine +
                "Cours : " + reecriture.Cours +
                Environment.NewLine +
                "Groupe : " + groupe,
                "Mise à jour source",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

            erreur =
                true;

            break;
          }

          if (nomsParColonne.ContainsKey(colonneCible)) {
            MessageBox.Show(
                "Impossible de réécrire la ligne source : deux intervenants ciblent le même groupe." +
                Environment.NewLine +
                Environment.NewLine +
                "Feuille : " + reecriture.SourceSheet +
                Environment.NewLine +
                "Ligne : " + reecriture.SourceRow +
                Environment.NewLine +
                "Groupe : " + groupe,
                "Mise à jour source",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

            erreur =
                true;

            break;
          }

          nomsParColonne[colonneCible] =
              nom;
        }

        if (erreur)
          continue;

        foreach (int colonneGroupe in groupes.Keys) {
          feuille.Cells[reecriture.SourceRow,colonneGroupe].Value =
              null;
        }

        foreach (KeyValuePair<int,string> affectation in nomsParColonne) {
          feuille.Cells[reecriture.SourceRow,affectation.Key].Value =
              affectation.Value;
        }

        foreach (DataRow row in lignesData) {
          string groupe =
              LireValeurRow(
                  row,
                  ExcelSchemaNames.Columns.Groupe);

          if (
              colonneParGroupe.TryGetValue(
                  groupe,
                  out int nouvelleColonneSource)
          ) {
            if (row.Table.Columns.Contains(ExcelSchemaNames.Columns.SourceColumn)) {
              row[ExcelSchemaNames.Columns.SourceColumn] =
                  nouvelleColonneSource.ToString();
            }
          }
        }

        lignesReecrites++;
      }

      if (lignesReecrites > 0) {
        try {
          package.Workbook.Calculate();
        }
        catch {
          // EPPlus ne recalcule pas toujours toutes les formules.
          // Les cellules sources sont quand même correctement modifiées.
        }

        package.Save();
      }

      return lignesReecrites;

      static string LireValeurRow(
          DataRow row,
          string colonne
      ) {
        if (row == null)
          return string.Empty;

        if (row.Table == null)
          return string.Empty;

        if (!row.Table.Columns.Contains(colonne))
          return string.Empty;

        return row[colonne]?.ToString()?.Trim() ?? string.Empty;
      }
    }


    /// <summary>
    /// Récupère la valeur d'une colonne depuis la DataRow liée à une ligne de DataGrid.
    /// Utile lorsque la colonne existe dans la source de données mais n'est pas affichée dans la vue.
    /// </summary>
    /// <param name="gridRow">Ligne du DataGrid.</param>
    /// <param name="columnName">Nom de la colonne à lire dans la DataRow liée.</param>
    /// <returns>Valeur trouvée, ou chaîne vide si elle est inaccessible.</returns>
    internal static string GetBoundRowValue_vue(DataRowView rowView,string columnName) {
      if (rowView == null)
        return string.Empty;

      if (rowView.Row == null || rowView.Row.Table == null)
        return string.Empty;

      if (!rowView.Row.Table.Columns.Contains(columnName))
        return string.Empty;

      return rowView.Row[columnName]?.ToString() ?? string.Empty;
    }

    #endregion Lecture, filtrage et affichage dans l'interface

    #region Recherche, calculs et accès aux données

    private static DataRow? FindLastRowByModuleAndCourse(
         DataTable table,
         string module,
         string cours
     ) {
      if (table == null)
        return null;

      for (int i = table.Rows.Count - 1;i >= 0;i--) {
        DataRow row = table.Rows[i];

        string rowModule =
            GetRowValue(row,ExcelSchemaNames.Columns.Module);

        string rowCours =
            GetRowValue(row,ExcelSchemaNames.Columns.Cours);

        bool isMatch =
            string.Equals(rowModule,module,StringComparison.OrdinalIgnoreCase) &&
            string.Equals(rowCours,cours,StringComparison.OrdinalIgnoreCase);

        if (isMatch)
          return row;
      }

      return null;
    }

    /// <summary>
    /// Recherche la dernière ligne d'une table dont la colonne indiquée correspond
    /// à la valeur attendue.
    /// </summary>
    /// <param name="table">Table dans laquelle effectuer la recherche.</param>
    /// <param name="columnName">Nom de la colonne à comparer.</param>
    /// <param name="expectedValue">Valeur attendue dans la colonne.</param>
    /// <returns>
    /// Dernière ligne correspondante, ou null si aucune correspondance n'est trouvée.
    /// </returns>
    private static DataRow FindLastRowByColumnValue(
        DataTable table,
        string columnName,
        string expectedValue
    ) {
      if (table == null)
        return null!;

      if (string.IsNullOrWhiteSpace(columnName))
        return null!;

      if (!table.Columns.Contains(columnName))
        return null!;

      DataRow foundRow = null!;

      foreach (DataRow row in table.Rows) {
        string? value = row[columnName].ToString();

        if (string.Equals(value,expectedValue,StringComparison.OrdinalIgnoreCase)) {
          foundRow = row;
        }
      }

      return foundRow;
    }

    /// <summary>
    /// Compte les volumes CM, TD et TP affichés dans le DataGrid fourni.
    /// </summary>
    /// <param name="table">DataGrid contenant les lignes à comptabiliser.</param>
    public static void CompterLesTypesDeCours(DataGrid table) {
      if (table == null)
        return;

      CompteurCM = 0;
      CompteurTD = 0;
      CompteurTP = 0;

      foreach (DataRowView ligne in table.Items.OfType<DataRowView>()) {
        string typeCours = GetBoundRowValue_vue(ligne,ExcelSchemaNames.Columns.Cours).Trim();
        string groupe = GetBoundRowValue_vue(ligne,ExcelSchemaNames.Columns.Nombre).Trim();

        if (!ushort.TryParse(groupe,out ushort valeur))
          continue;

        switch (typeCours) {
          case "CM":
            CompteurCM += valeur;
            break;
          case "TD":
            CompteurTD += valeur;
            break;
          case "TP":
            CompteurTP += valeur;
            break;
        }
      }
    }

    /// <summary>
    /// Recherche la colonne dans la vue DataGrid.
    /// </summary>
    /// <param name="table"></param>
    /// <param name="nomColonne"></param>
    /// <returns></returns>
    internal static ushort RechercheColonne_vue(DataGrid table,string nomColonne) {
      if (table == null || string.IsNullOrWhiteSpace(nomColonne))
        return 0;

      for (int i = 0;i < table.Columns.Count;i++) {
        DataGridColumn colonne = table.Columns[i];
        string entete = colonne.Header?.ToString() ?? string.Empty;
        string membreTri = colonne.SortMemberPath ?? string.Empty;
        string cheminBinding = string.Empty;

        if (colonne is DataGridBoundColumn colonneLiee && colonneLiee.Binding is Binding binding)
          cheminBinding = binding.Path?.Path ?? string.Empty;

        if (string.Equals(entete,nomColonne,StringComparison.OrdinalIgnoreCase) ||
            string.Equals(membreTri,nomColonne,StringComparison.OrdinalIgnoreCase) ||
            string.Equals(cheminBinding,nomColonne,StringComparison.OrdinalIgnoreCase))
          return (ushort)i;
      }

      return 0;
    }

    /// <summary>
    /// Recherche une valeur dans une table en trouvant d'abord une ligne contenant le module indiqué,
    /// puis en retournant la valeur de la colonne demandée.
    /// </summary>
    /// <param name="nomModule">Nom du module recherché.</param>
    /// <param name="entêteColonne">Nom de la colonne dont la valeur doit être retournée.</param>
    /// <param name="tableDeDonnée">Nom de la table à parcourir.</param>
    /// <returns>Valeur trouvée, ou chaîne vide si aucune correspondance n'est trouvée.</returns>
    internal static string RechercheValeur(
        string nomModule,
        string entêteColonne,
        string tableDeDonnée) {

      return RechercheValeur(DataSetExcel,nomModule,entêteColonne,tableDeDonnée);
    }

    #endregion Recherche, calculs et accès aux données

    #region Gestion des erreurs et diagnostics

    /// <summary>
    /// Affiche une MessageBox détaillant une erreur survenue lors d'une opération OleDb.
    /// </summary>
    /// <param name="ex">Exception interceptée.</param>
    /// <param name="context">Nom de la méthode ou du contexte dans lequel l'erreur est survenue.</param>
    /// <param name="query">Requête OleDb exécutée au moment de l'erreur.</param>
    internal static void ShowOleDbError(
        Exception ex,
        string context,
        string query
    ) {
      string displayedQuery =
          string.IsNullOrWhiteSpace(query)
              ? "(aucune requête disponible)"
              : query;

      MessageBox.Show(
          "Erreur OleDb" +
          Environment.NewLine +
          Environment.NewLine +
          "Contexte : " + context +
          Environment.NewLine +
          Environment.NewLine +
          "Message :" +
          Environment.NewLine +
          ex.Message +
          Environment.NewLine +
          Environment.NewLine +
          "Requête :" +
          Environment.NewLine +
          displayedQuery,
          "Erreur base de données",
          MessageBoxButton.OK,
          MessageBoxImage.Error
      );
    }

    internal static string RechercheValeur(DataSet dataSetSource,string nomModule,string entêteColonne,string tableDeDonnée) {
      if (dataSetSource == null || string.IsNullOrWhiteSpace(tableDeDonnée))
        return string.Empty;

      if (!dataSetSource.Tables.Contains(tableDeDonnée))
        return string.Empty;

      DataTable table = dataSetSource.Tables[tableDeDonnée]!;

      string colonneModule = TrouverNomColonneIgnoreCase(table,ExcelSchemaNames.Columns.Module);
      string colonneValeur = TrouverNomColonneIgnoreCase(table,entêteColonne);

      if (string.IsNullOrWhiteSpace(colonneModule) || string.IsNullOrWhiteSpace(colonneValeur))
        return string.Empty;

      foreach (DataRow ligne in table.Rows) {
        string module = ligne[colonneModule]?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(module,nomModule,StringComparison.OrdinalIgnoreCase))
          return ligne[colonneValeur]?.ToString()?.Trim() ?? string.Empty;
      }

      return string.Empty;
    }

    /// <summary>
    /// Affiche une MessageBox détaillant une erreur survenue lors de l'application
    /// d'un filtre DataView.
    /// </summary>
    /// <param name="ex">Exception interceptée.</param>
    /// <param name="tableName">Nom de la table concernée par le filtre.</param>
    /// <param name="filter">Filtre DataView appliqué au moment de l'erreur.</param>
    /// <param name="dataTable">Table contenant les colonnes disponibles pour le filtre.</param>
    private static void ShowFilterError(Exception ex,string tableName,string filter,DataTable dataTable) {
      StringBuilder columnsBuilder = new();

      if (dataTable != null) {
        foreach (DataColumn column in dataTable.Columns) {
          columnsBuilder.AppendLine("- " + column.ColumnName);
        }
      }

      string displayedFilter = string.IsNullOrWhiteSpace(filter) ? "(aucun filtre disponible)" : filter;
      string displayedColumns = columnsBuilder.Length == 0 ? "(aucune colonne disponible)" : columnsBuilder.ToString();

      MessageBox.Show(
          "Erreur pendant l'application du filtre." +
          Environment.NewLine +
          Environment.NewLine +
          "Table :" +
          Environment.NewLine +
          tableName +
          Environment.NewLine +
          Environment.NewLine +
          "Filtre :" +
          Environment.NewLine +
          displayedFilter +
          Environment.NewLine +
          Environment.NewLine +
          "Message :" +
          Environment.NewLine +
          ex.Message +
          Environment.NewLine +
          Environment.NewLine +
          "Colonnes disponibles :" +
          Environment.NewLine +
          displayedColumns,
          "Erreur filtre",
          MessageBoxButton.OK,
          MessageBoxImage.Warning
      );
    }

    internal static void DiagnostiquerLigneServiceFeuilleSemestre(
    string cheminFichier,
    string nomFeuille,
    string codeServiceRecherche
) {
      if (string.IsNullOrWhiteSpace(cheminFichier))
        return;

      if (!File.Exists(cheminFichier))
        return;

      using (ExcelPackage package =
          new ExcelPackage(new FileInfo(cheminFichier))) {
        ExcelWorksheet? feuille =
            package.Workbook.Worksheets[nomFeuille];

        if (feuille == null || feuille.Dimension == null)
          return;

        Debug.WriteLine("===== Diagnostic ligne service =====");
        Debug.WriteLine("Feuille : " + nomFeuille);
        Debug.WriteLine("Code recherché : " + codeServiceRecherche);

        int rowMax =
            Math.Min(feuille.Dimension.End.Row,300);

        int colMax =
            Math.Min(feuille.Dimension.End.Column,250);

        for (int row = 1;row <= rowMax;row++) {
          string codeService =
              feuille.Cells[row,2].Text.Trim();

          if (!string.Equals(
                  codeService,
                  codeServiceRecherche,
                  StringComparison.OrdinalIgnoreCase)) {
            continue;
          }

          Debug.WriteLine("-----------------------------------");
          Debug.WriteLine("Ligne trouvée : " + row);
          Debug.WriteLine("-----------------------------------");

          for (int ligne = row - 2;ligne <= row + 5;ligne++) {
            if (ligne < 1)
              continue;

            Debug.WriteLine("---- Ligne Excel " + ligne + " ----");

            for (int col = 1;col <= colMax;col++) {
              string valeur =
                  feuille.Cells[ligne,col].Text.Trim();

              if (string.IsNullOrWhiteSpace(valeur))
                continue;

              Debug.WriteLine(
                  "L" + ligne +
                  " C" + col +
                  " = [" + valeur + "]");
            }
          }

          Debug.WriteLine("===== Fin diagnostic ligne service =====");
          return;
        }

        Debug.WriteLine("Code service non trouvé.");
        Debug.WriteLine("===== Fin diagnostic ligne service =====");
      }
    }

    #endregion Gestion des erreurs et diagnostics

    #region Utilitaires DataTable et DataRow

    private static void SetRowValue(
    DataRow row,
    string columnName,
    object value
) {
      if (!row.Table.Columns.Contains(columnName))
        return;

      row[columnName] =
          value ?? DBNull.Value;
    }

    private static string TrouverNomColonneIgnoreCase(
        DataTable table,
        string nomRecherche
    ) {
      if (table == null)
        return string.Empty;

      foreach (DataColumn colonne in table.Columns) {
        if (string.Equals(
                colonne.ColumnName,
                nomRecherche,
                StringComparison.OrdinalIgnoreCase)) {
          return colonne.ColumnName;
        }
      }

      return string.Empty;
    }

    private static string GetRowValueIfColumnExists(
    DataRow row,
    string columnName
) {
      if (row == null)
        return string.Empty;

      if (string.IsNullOrWhiteSpace(columnName))
        return string.Empty;

      if (row.Table == null)
        return string.Empty;

      if (!row.Table.Columns.Contains(columnName))
        return string.Empty;

      return row[columnName]?.ToString()?.Trim() ?? string.Empty;
    }

    private static int LireEntierDepuisRow(
    DataRow row,
    string columnName
) {
      string valeur =
          GetRowValueIfColumnExists(
              row,
              columnName);

      if (string.IsNullOrWhiteSpace(valeur))
        return 0;

      valeur =
          valeur.Replace(",",".");

      if (decimal.TryParse(
              valeur,
              System.Globalization.NumberStyles.Any,
              System.Globalization.CultureInfo.InvariantCulture,
              out decimal resultat)) {
        return Convert.ToInt32(
            Math.Round(resultat));
      }
      return 0;
    }

    private static string CreerNomColonneUnique(DataTable table,string nomColonne) {
      string nomBase =
          string.IsNullOrWhiteSpace(nomColonne)
              ? "Colonne"
              : nomColonne.Trim();

      string nomFinal =
          nomBase;

      int index =
          2;

      while (table.Columns.Contains(nomFinal)) {
        nomFinal =
            nomBase + "_" + index;

        index++;
      }

      return nomFinal;
    }

    /// <summary>
    /// Récupère sous forme de texte la valeur d'une colonne dans une ligne DataRow.
    /// </summary>
    /// <param name="row">Ligne source à lire.</param>
    /// <param name="columnName">Nom de la colonne à lire.</param>
    /// <returns>
    /// Valeur de la colonne sous forme de chaîne, ou une chaîne vide si la ligne est nulle
    /// ou si la colonne n'existe pas.
    /// </returns>
    private static string GetRowValue(
        DataRow row,
        string columnName
    ) {
      if (row == null)
        return string.Empty;

      if (!row.Table.Columns.Contains(columnName))
        return string.Empty;

      return row[columnName].ToString()!;
    }

    #endregion Utilitaires DataTable et DataRow

    internal sealed class ExcelSourceCellChange {
      internal string SourceSheet { get; set; } =
          string.Empty;

      internal int SourceRow { get; set; }

      internal int SourceColumn { get; set; }

      internal string ColumnName { get; set; } =
          string.Empty;

      internal string NewValue { get; set; } =
          string.Empty;
    }

    internal sealed class ExcelSourceRowRewrite {
      internal string SourceSheet { get; set; } =
          string.Empty;

      internal int SourceRow { get; set; }

      internal string Cours { get; set; } =
          string.Empty;
    }
  }

}
