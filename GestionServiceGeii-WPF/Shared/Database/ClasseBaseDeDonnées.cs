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
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using GestionServiceGeii.Core.Excel;
using GestionServiceGeii.Core.Profiles;
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
    private readonly ServiceManagerProfile _profile;

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

    public ClasseBaseDeDonnées(ServiceManagerProfile profile) {
      _profile = profile;
    }

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

    private static DataSet ConstruireDataSetServiceDepuisClasseur(string cheminFichier) {
      Stopwatch chronoTotal = Stopwatch.StartNew();
      Stopwatch chrono = Stopwatch.StartNew();

      DataSet dataSet = new("FICHIER DE SERVICE");

      if (string.IsNullOrWhiteSpace(cheminFichier) || !File.Exists(cheminFichier))
        return dataSet;

      string connexion =
          $"Provider=Microsoft.ACE.OLEDB.12.0;" +
          $"Data Source={cheminFichier};" +
          $"Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1\";";

      using OleDbConnection connection = new(connexion);

      chrono.Restart();
      connection.Open();
      Debug.WriteLine($"[CHRONO CACHE OLEDB] Ouverture connexion : {chrono.ElapsedMilliseconds} ms");

      chrono.Restart();
      DataTable tableGlobal = LireGlobal_OleDb(connection);
      Debug.WriteLine($"[CHRONO CACHE OLEDB] Lecture Global : {chrono.ElapsedMilliseconds} ms | {tableGlobal.Rows.Count} ligne(s)");

      chrono.Restart();
      DataTable grilleS1 = LireFeuilleSemestreCommeGrille_OleDb(connection,"S1");
      Debug.WriteLine($"[CHRONO CACHE OLEDB] Lecture S1 : {chrono.ElapsedMilliseconds} ms | {grilleS1.Rows.Count} ligne(s) x {grilleS1.Columns.Count} colonne(s)");

      chrono.Restart();
      DataTable grilleS2 = LireFeuilleSemestreCommeGrille_OleDb(connection,"S2");
      Debug.WriteLine($"[CHRONO CACHE OLEDB] Lecture S2 : {chrono.ElapsedMilliseconds} ms | {grilleS2.Rows.Count} ligne(s) x {grilleS2.Columns.Count} colonne(s)");

      chrono.Restart();
      DataTable grilleS3 = LireFeuilleSemestreCommeGrille_OleDb(connection,"S3 FIFA");
      Debug.WriteLine($"[CHRONO CACHE OLEDB] Lecture S3 FIFA : {chrono.ElapsedMilliseconds} ms | {grilleS3.Rows.Count} ligne(s) x {grilleS3.Columns.Count} colonne(s)");

      List<AffectationGroupeSemestre> affectationsS3 = LireAffectationsGroupesDepuisFeuilleS3Fifa_OleDb(grilleS3);

      Debug.WriteLine("===== S3 FIFA : DIAGNOSTIC R3-08 =====");

      for (int i = 0;i < grilleS3.Rows.Count;i++) {
        string code = Convert.ToString(grilleS3.Rows[i][4])?.Trim() ?? string.Empty;

        if (!code.Contains("R308",StringComparison.OrdinalIgnoreCase))
          continue;

        string contenu = string.Join(" | ",grilleS3.Rows[i].ItemArray
            .Select((valeur,index) => $"C{index}={valeur}"));

        Debug.WriteLine($"Ligne DataTable {i} : {contenu}");
      }

      NormaliserColonnesGlobalPourInterface(tableGlobal);
      EnrichirCodesGlobalDepuisLibelleCourt(tableGlobal);

      Debug.WriteLine($"[CHRONO CACHE OLEDB] Normalisation Global : {chrono.ElapsedMilliseconds} ms");

      chrono.Restart();

      ReconstruireNomsGlobalDepuisAffectations_OleDb(tableGlobal,grilleS1,grilleS2);
      ReconstruireNomsGlobalDepuisAffectationsS3_OleDb(tableGlobal,grilleS3);

      Debug.WriteLine($"[CHRONO CACHE OLEDB] Reconstruction NOMS : {chrono.ElapsedMilliseconds} ms");

      chrono.Restart();

      EnrichirGlobalDepuisLignesTechniques(tableGlobal);

      Debug.WriteLine($"[CHRONO CACHE OLEDB] Enrichissement lignes techniques : {chrono.ElapsedMilliseconds} ms");

      chrono.Restart();

      EnrichirGroupesGlobalDepuisFeuilleSemestre_OleDb(
          grilleS1,
          tableGlobal);

      EnrichirGroupesGlobalDepuisFeuilleSemestre_OleDb(
          grilleS2,
          tableGlobal);

      Debug.WriteLine($"[CHRONO CACHE OLEDB] Groupes S1/S2 : {chrono.ElapsedMilliseconds} ms");

      chrono.Restart();

      AjouterLignesManquantesDepuisFeuilleSemestre_OleDb(
          grilleS1,
          tableGlobal);

      AjouterLignesManquantesDepuisFeuilleSemestre_OleDb(
          grilleS2,
          tableGlobal);

      Debug.WriteLine($"[CHRONO CACHE OLEDB] Lignes manquantes S1/S2 : {chrono.ElapsedMilliseconds} ms");

      chrono.Restart();

      EnrichirInfosGlobalDepuisFeuilleSemestre_OleDb(
          grilleS1,
          tableGlobal);

      EnrichirInfosGlobalDepuisFeuilleSemestre_OleDb(
          grilleS2,
          tableGlobal);

      Debug.WriteLine($"[CHRONO CACHE OLEDB] INFOS S1/S2 : {chrono.ElapsedMilliseconds} ms");

      chrono.Restart();

      tableGlobal.AcceptChanges();
      dataSet.Tables.Add(tableGlobal);

      Debug.WriteLine($"[CHRONO CACHE OLEDB] Ajout Global au DataSet : {chrono.ElapsedMilliseconds} ms");
      Debug.WriteLine($"[CHRONO CACHE OLEDB] TOTAL : {chronoTotal.ElapsedMilliseconds} ms | Global={tableGlobal.Rows.Count} ligne(s)");

      return dataSet;
    }

    internal static Task ReconstruireCacheFichierServiceAsync(ClasseExcel fichierDeService) {
      if (fichierDeService == null ||
          string.IsNullOrWhiteSpace(fichierDeService.CheminFichier) ||
          !File.Exists(fichierDeService.CheminFichier))
        return Task.CompletedTask;

      string cheminFichier = fichierDeService.CheminFichier;

      return Task.Run(() => {
        DataSet dataSet = ConstruireDataSetServiceDepuisClasseur(cheminFichier);
        ServiceWorkbookDiskCache.SaveServiceDataSet(cheminFichier,dataSet);
      });
    }

    private static readonly SemaphoreSlim _verrouReconstructionCacheService = new(1,1);

    internal static void ReconstruireCacheFichierService(ClasseExcel fichierDeService) {
      if (fichierDeService == null ||
          string.IsNullOrWhiteSpace(fichierDeService.CheminFichier) ||
          !File.Exists(fichierDeService.CheminFichier))
        return;

      DataSet dataSet = ConstruireDataSetServiceDepuisClasseur(
          fichierDeService.CheminFichier);

      ServiceWorkbookDiskCache.SaveServiceDataSet(
          fichierDeService.CheminFichier,
          dataSet);
    }

    /// <summary>
    /// Lit les feuilles utiles du fichier ExcelApp de service via OleDb et les ajoute
    /// au DataSet de service.
    /// </summary>
    /// <param name="fichierDeService">Fichier ExcelApp de service à lire.</param>
    /// <returns>DataSet contenant les tables lues depuis le fichier de service.</returns>
    public static DataSet LectureFichierDeService(ClasseExcel fichierDeService) {
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

        // ------------------------------------------------------------
        // 1. Tentative de chargement depuis le cache
        // ------------------------------------------------------------
        if (ServiceWorkbookDiskCache.TryLoadServiceDataSet(cheminFichier,out DataSet? serviceDataSetDepuisCache)) {

          if (serviceDataSetDepuisCache != null) {
            if (serviceDataSetDepuisCache.Tables.Contains(ExcelSchemaNames.Tables.NomTableGlobal)) {
              DataTableExcel = serviceDataSetDepuisCache.Tables[ExcelSchemaNames.Tables.NomTableGlobal]!;
              InitialiserReferencesSources(DataTableExcel);
            }

            Debug.WriteLine(">>> SERVICE CHARGE DEPUIS CACHE");

            return serviceDataSetDepuisCache;
          }
        }

        // ------------------------------------------------------------
        // 2. Cache absent ou invalide :
        //    reconstruction depuis le classeur Excel
        // ------------------------------------------------------------
        serviceDataSet = ConstruireDataSetServiceDepuisClasseur(cheminFichier);

        foreach (DataTable table in serviceDataSet.Tables) {
          Debug.WriteLine(
              "Table service chargée : " +
              table.TableName +
              " | " +
              table.Rows.Count +
              " lignes");
        }

        // ------------------------------------------------------------
        // 3. Sauvegarde du DataSet dans le cache
        //
        // IMPORTANT :
        // avant InitialiserReferencesSources(), car les métadonnées
        // runtime ajoutées ensuite ne doivent pas être sérialisées.
        // ------------------------------------------------------------
        ServiceWorkbookDiskCache.SaveServiceDataSet(
            cheminFichier,
            serviceDataSet);

        // ------------------------------------------------------------
        // 4. Initialisation des références runtime
        // ------------------------------------------------------------
        if (serviceDataSet.Tables.Contains(
                ExcelSchemaNames.Tables.NomTableGlobal)) {

          DataTableExcel =
              serviceDataSet.Tables[
                  ExcelSchemaNames.Tables.NomTableGlobal]!;

          InitialiserReferencesSources(DataTableExcel);
        }

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

        return serviceDataSet;
      }
    }
    private static void DiagnostiquerEner11(DataTable tableGlobal,string etape) {
      Debug.WriteLine("===== " + etape + " =====");

      foreach (DataRow row in tableGlobal.Rows) {
        string module = row.Table.Columns.Contains("Module") ? row["Module"]?.ToString() ?? "" : "";

        if (module != "ENER1-1")
          continue;

        string cours = row.Table.Columns.Contains("Cours") ? row["Cours"]?.ToString() ?? "" : "";

        if (cours != "TD" && cours != "TP")
          continue;

        string noms = row.Table.Columns.Contains("Noms") ? row["Noms"]?.ToString() ?? "" : "";
        string groupe = row.Table.Columns.Contains("Groupe") ? row["Groupe"]?.ToString() ?? "" : "";

        Debug.WriteLine($"{cours} | {noms} | {groupe}");
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

    private static DataTable LireFeuilleSemestreCommeGrille_OleDb(OleDbConnection connection,string nomFeuille) {
      DataTable grille = new(nomFeuille);

      if (connection == null || connection.State != ConnectionState.Open || string.IsNullOrWhiteSpace(nomFeuille))
        return grille;

      using OleDbCommand commande = new($"SELECT * FROM [{nomFeuille}$]",connection);
      using OleDbDataReader? lecteur = commande.ExecuteReader();

      if (lecteur == null)
        return grille;

      for (int col = 0;col < lecteur.FieldCount;col++)
        grille.Columns.Add("C" + (col + 1),typeof(string));

      while (lecteur.Read()) {
        DataRow ligne = grille.NewRow();

        for (int col = 0;col < lecteur.FieldCount;col++)
          ligne[col] = lecteur.IsDBNull(col) ? string.Empty : Convert.ToString(lecteur.GetValue(col))?.Trim() ?? string.Empty;

        grille.Rows.Add(ligne);
      }

      grille.AcceptChanges();
      return grille;
    }

    private static string LireCelluleGrille(DataTable grille,int row,int col) {
      if (grille == null || row <= 0 || col <= 0)
        return string.Empty;

      int indexRow = row - 1;
      int indexCol = col - 1;

      if (indexRow >= grille.Rows.Count || indexCol >= grille.Columns.Count)
        return string.Empty;

      object valeur = grille.Rows[indexRow][indexCol];

      if (valeur == null || valeur == DBNull.Value)
        return string.Empty;

      return Convert.ToString(valeur)?.Trim() ?? string.Empty;
    }

    private static bool TrouverCelluleTexteFeuilleSemestre(DataTable grille,string texteRecherche,out int ligneTrouvee,out int colonneTrouvee) {
      ligneTrouvee = 0;
      colonneTrouvee = 0;

      if (grille == null || string.IsNullOrWhiteSpace(texteRecherche))
        return false;

      int rowMax = Math.Min(grille.Rows.Count,40);
      int colMax = Math.Min(grille.Columns.Count,MaxColonnesRechercheGroupes);

      for (int row = 1;row <= rowMax;row++) {
        for (int col = 1;col <= colMax;col++) {
          string valeur = LireCelluleGrille(grille,row,col);

          if (!string.Equals(valeur,texteRecherche,StringComparison.OrdinalIgnoreCase))
            continue;

          ligneTrouvee = row;
          colonneTrouvee = col;
          return true;
        }
      }

      return false;
    }


    private static DataTable LireFeuilleExcelCommeDataTable_OleDb(string cheminFichier,string nomTable) {
      DataTable table = new(nomTable);

      if (string.IsNullOrWhiteSpace(cheminFichier) || !File.Exists(cheminFichier))
        return table;

      string connexion =
          $"Provider=Microsoft.ACE.OLEDB.12.0;" +
          $"Data Source={cheminFichier};" +
          $"Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";";

      using OleDbConnection connection = new(connexion);
      connection.Open();

      using OleDbCommand commande = new($"SELECT * FROM [{nomTable}$]",connection);
      using OleDbDataReader lecteur = commande.ExecuteReader();

      if (lecteur == null)
        return table;

      for (int col = 0;col < lecteur.FieldCount;col++) {
        string nomColonne = lecteur.GetName(col).Trim();

        if (string.IsNullOrWhiteSpace(nomColonne))
          nomColonne = "Colonne_" + (col + 1);

        nomColonne = CreerNomColonneUnique(table,nomColonne);
        table.Columns.Add(nomColonne,typeof(string));
      }

      while (lecteur.Read()) {
        DataRow ligne = table.NewRow();
        bool ligneVide = true;

        for (int col = 0;col < lecteur.FieldCount;col++) {
          string valeur;

          if (lecteur.IsDBNull(col))
            valeur = string.Empty;
          else
            valeur = Convert.ToString(lecteur.GetValue(col))?.Trim() ?? string.Empty;

          if (!string.IsNullOrWhiteSpace(valeur))
            ligneVide = false;

          ligne[col] = valeur;
        }

        if (!ligneVide)
          table.Rows.Add(ligne);
      }

      table.AcceptChanges();
      return table;
    }

    private static DataTable LireFeuilleExcelCommeDataTable_Epplus(ExcelWorksheet feuille,string nomTable) {
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

      RenommerColonneSiExiste(table,"DUREE_ATTENDUE",ExcelSchemaNames.Columns.Duree);

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

    internal static void MiseAJourDureesAttenduesGlobal_DataTable(DataSet dataSetExcelService,string module,string valeurCM,string valeurTD,string valeurTP) {
      DataTable? table = dataSetExcelService.Tables[ExcelSchemaNames.Tables.NomTableGlobal];

      if (table == null)
        return;

      foreach (DataRow row in table.Rows) {
        string moduleLigne = row[ExcelSchemaNames.Columns.Module]?.ToString()?.Trim() ?? string.Empty;

        if (!string.Equals(moduleLigne,module,StringComparison.OrdinalIgnoreCase))
          continue;

        string cours = row[ExcelSchemaNames.Columns.Cours]?.ToString()?.Trim().ToUpperInvariant() ?? string.Empty;

        string valeur = cours switch {
          "CM" => valeurCM,
          "TD" => valeurTD,
          "TP" => valeurTP,
          _ => string.Empty
        };

        if (string.IsNullOrWhiteSpace(valeur))
          continue;

        row[ExcelSchemaNames.Columns.Duree] = valeur;
      }
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

    internal static void EnrichirInfosGlobalDepuisFeuilleSemestre_Epplus(ExcelWorksheet feuille,DataTable tableGlobal) {
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

      if (feuille.Dimension == null)
        return;

      string nomFeuille = feuille.Name;

      string colonneLibelleCourt = TrouverNomColonneIgnoreCase(tableGlobal,ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      string colonneInfos = TrouverNomColonneIgnoreCase(tableGlobal,ExcelSchemaNames.Columns.Infos);

      if (string.IsNullOrWhiteSpace(colonneInfos)) {
        tableGlobal.Columns.Add(ExcelSchemaNames.Columns.Infos,typeof(string));
        colonneInfos = ExcelSchemaNames.Columns.Infos;
      }

      int ligneSalles;
      int colonneSalles;

      if (!TrouverCelluleTexteFeuilleSemestre(feuille,"SALLES",out ligneSalles,out colonneSalles)) {
        return;
      }

      int colMax = TrouverDerniereColonneGroupes(feuille,ligneSalles,ligneSalles + 1,colonneSalles + 1);

      if (colMax < colonneSalles + 1)
        return;

      Dictionary<int,string> groupesTd = LireGroupesFeuilleSemestre(feuille,ligneSalles,colonneSalles + 1,colMax,estTp: false);

      Dictionary<int,string> groupesTp = LireGroupesFeuilleSemestre(feuille,ligneSalles + 1,colonneSalles + 1,colMax,estTp: true);

      Regex regexCodeService = new Regex(@"^(R\d+-\d{2})-(CM|TD|TP|DS)$",RegexOptions.IgnoreCase | RegexOptions.Compiled);

      int rowMax = TrouverDerniereLigneUtileSemestre(feuille);

      string moduleCourant = string.Empty;
      string infosCourantes = string.Empty;
      int lignesModifiees = 0;

      for (int row = 1;row <= rowMax;row++) {
        string celluleA = feuille.Cells[row,1].Text.Trim();
        string moduleDetecte = ExtraireModuleDepuisCelluleFeuilleSemestre(celluleA);
        if (EstCodeModuleFeuilleSemestreValide(moduleDetecte)) {
          moduleCourant = moduleDetecte;
          infosCourantes = string.Empty;
        }
        else if (!string.IsNullOrWhiteSpace(celluleA) && EstInfoFeuilleSemestreValide(celluleA)) {
          infosCourantes = celluleA;
        }

        string codeService = feuille.Cells[row,2].Text.Trim();
        Match match = regexCodeService.Match(codeService);

        if (!match.Success)
          continue;

        if (string.IsNullOrWhiteSpace(moduleCourant))
          continue;

        string cours = match.Groups[2].Value.ToUpperInvariant();

        if (cours != "TD" && cours != "TP")
          continue;

        if (string.IsNullOrWhiteSpace(infosCourantes))
          continue;

        Dictionary<int,string> groupes = cours == "TD" ? groupesTd : groupesTp;

        int dureeLigne = LireDureeDepuisLigneFeuilleSemestre(feuille,row,colonneSalles);

        foreach (KeyValuePair<int,string> groupe in groupes) {
          string nomIntervenant = feuille.Cells[row,groupe.Key].Text.Trim();
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

            string infosActuelles = GetRowValueIfColumnExists(globalRow,colonneInfos);
            if (
                !string.IsNullOrWhiteSpace(infosActuelles) && infosActuelles != "0") {
              continue;
            }

            globalRow[colonneInfos] = infosCourantes;
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

      tableGlobal.AcceptChanges();
    }

    private static bool EstCodeModuleFeuilleSemestreValide(
        string module
    ) {
      if (string.IsNullOrWhiteSpace(module))
        return false;

      return Regex.IsMatch(
          module.Trim(),
          @"^[A-Z]{2,}[A-Z0-9]*\d+(?:-\d+)?$",
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

    private static string NormaliserLibelleCourtS3(string libelle) {
      if (string.IsNullOrWhiteSpace(libelle))
        return string.Empty;

      Match match = Regex.Match(libelle.Trim(),@"^R3(\d{2})[A-Z]?-(CM|TD|TP|DS)$",RegexOptions.IgnoreCase);

      if (!match.Success)
        return libelle.Trim();

      return $"R3-{match.Groups[1].Value}-{match.Groups[2].Value}".ToUpperInvariant();
    }

    internal static void EnrichirGroupesGlobalDepuisFeuilleSemestre_Epplus(
        string cheminFichier,
        DataTable tableGlobal,
        ExcelWorksheet feuille) {

      if (string.IsNullOrWhiteSpace(cheminFichier))
        return;

      if (!File.Exists(cheminFichier))
        return;

      if (tableGlobal == null)
        return;

      if (feuille == null || feuille.Dimension == null)
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

      string nomFeuille = feuille.Name;

      string colonneLibelleCourt =
          TrouverNomColonneIgnoreCase(
              tableGlobal,
              ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      Dictionary<string,List<string>> groupesParCle =
          ObtenirGroupesParCleSemestreDepuisCache(
              cheminFichier,
              feuille);

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

        string cleUtilisee = cle;

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

          cleUtilisee = cleAvecSuffixeUn;

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

    private static Dictionary<string,List<string>> ObtenirGroupesParCleSemestreDepuisCache(string cheminFichier,string nomFeuille) {
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

    private static Dictionary<string,List<string>> ConstruireGroupesParCleSemestreDepuisClasseur(string cheminFichier,string nomFeuille) {
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

    private static Dictionary<string,List<string>> ObtenirGroupesParCleSemestreDepuisCache(string cheminFichier,ExcelWorksheet feuille) {
      string nomFeuille = feuille.Name;

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
          ConstruireGroupesParCleSemestreDepuisClasseur(feuille);

      lock (_cacheGroupesSemestresLock) {
        if (!_cacheGroupesSemestres.ContainsKey(cleCache))
          _cacheGroupesSemestres.Add(cleCache,groupes);

        return _cacheGroupesSemestres[cleCache];
      }
    }

    private static Dictionary<string,List<string>> ConstruireGroupesParCleSemestreDepuisClasseur(ExcelWorksheet feuille) {
      Dictionary<string,List<string>> groupesParCle = new(StringComparer.OrdinalIgnoreCase);

      if (feuille.Dimension == null)
        return groupesParCle;

      string nomFeuille = feuille.Name;

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

    internal static void AjouterLignesManquantesDepuisFeuilleSemestre_OleDb(DataTable grille,DataTable tableGlobal) {
      if (grille == null || tableGlobal == null)
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

      string nomFeuille = grille.TableName;

      string colonneLibelleCourt = TrouverNomColonneIgnoreCase(tableGlobal,ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      Stopwatch chrono = Stopwatch.StartNew();

      List<AffectationGroupeSemestre> affectations = LireAffectationsGroupesDepuisFeuilleSemestre_OleDb(grille);

      if (affectations.Count == 0)
        return;

      int prochainId = TrouverProchainIdGlobal(tableGlobal);
      AjouterColonnesSourceSemestreSiAbsentes(tableGlobal);

      ReconstruireLignesInfosSpecialesDepuisAffectations(tableGlobal,colonneLibelleCourt,affectations);

      Dictionary<string,DataRow> indexActivites = new(StringComparer.OrdinalIgnoreCase);
      Dictionary<string,DataRow> indexModeles = new(StringComparer.OrdinalIgnoreCase);

      foreach (DataRow row in tableGlobal.Rows) {
        string semestre = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Semestre);
        string module = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Module);
        string cours = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Cours);
        string noms = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Noms);
        string groupe = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Groupe);
        string libelleCourt = GetRowValueIfColumnExists(row,colonneLibelleCourt);

        string cleActivite = ConstruireCleActiviteGlobal(semestre,module,cours,noms,groupe,libelleCourt);

        if (!indexActivites.ContainsKey(cleActivite))
          indexActivites.Add(cleActivite,row);

        string cleModele = ConstruireCleModeleGlobal(semestre,module,cours,libelleCourt);

        if (!indexModeles.ContainsKey(cleModele))
          indexModeles.Add(cleModele,row);

      }

      int nbExistantes = 0;
      int nbAjoutees = 0;
      int nbSansModele = 0;

      foreach (AffectationGroupeSemestre affectation in affectations) {
        if (string.IsNullOrWhiteSpace(affectation.Noms))
          continue;

        if (EstLigneTechniqueGlobal(affectation.Noms))
          continue;

        string cleActivite = ConstruireCleActiviteGlobal(
            affectation.Semestre,
            affectation.Module,
            affectation.Cours,
            affectation.Noms,
            affectation.Groupe,
            affectation.LibelleCourt);

        if (indexActivites.TryGetValue(cleActivite,out DataRow? ligneExistante)) {
          if (ligneExistante.Table.Columns.Contains(ExcelSchemaNames.Columns.Infos))
            ligneExistante[ExcelSchemaNames.Columns.Infos] = affectation.Infos?.Trim() ?? string.Empty;

          if (ligneExistante.Table.Columns.Contains(ExcelSchemaNames.Columns.SourceSheet))
            ligneExistante[ExcelSchemaNames.Columns.SourceSheet] = affectation.SourceSheet;

          if (ligneExistante.Table.Columns.Contains(ExcelSchemaNames.Columns.SourceRow))
            ligneExistante[ExcelSchemaNames.Columns.SourceRow] = affectation.SourceRow.ToString();

          if (ligneExistante.Table.Columns.Contains(ExcelSchemaNames.Columns.SourceColumn))
            ligneExistante[ExcelSchemaNames.Columns.SourceColumn] = affectation.SourceColumn.ToString();


          nbExistantes++;
          continue;
        }

        string cleModele = ConstruireCleModeleGlobal(
            affectation.Semestre,
            affectation.Module,
            affectation.Cours,
            affectation.LibelleCourt);

        indexModeles.TryGetValue(cleModele,out DataRow? modele);

        if (modele == null) {
          nbSansModele++;
          continue;
        }

        DataRow nouvelleLigne = tableGlobal.NewRow();

        foreach (DataColumn colonne in tableGlobal.Columns)
          nouvelleLigne[colonne.ColumnName] = modele[colonne.ColumnName];

        if (tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.ID)) {
          nouvelleLigne[ExcelSchemaNames.Columns.ID] = prochainId.ToString();
          prochainId++;
        }

        nouvelleLigne[ExcelSchemaNames.Columns.Noms] = affectation.Noms;
        nouvelleLigne[ExcelSchemaNames.Columns.Groupe] = affectation.Groupe;
        nouvelleLigne[ExcelSchemaNames.Columns.SourceSheet] = affectation.SourceSheet;
        nouvelleLigne[ExcelSchemaNames.Columns.SourceRow] = affectation.SourceRow.ToString();
        nouvelleLigne[ExcelSchemaNames.Columns.SourceColumn] = affectation.SourceColumn.ToString();
        nouvelleLigne[ExcelSchemaNames.Columns.Nombre] = "1";
        nouvelleLigne[ExcelSchemaNames.Columns.Cours] = affectation.Cours;
        nouvelleLigne[ExcelSchemaNames.Columns.Module] = modele[ExcelSchemaNames.Columns.Module]?.ToString() ?? affectation.Module;
        nouvelleLigne[colonneLibelleCourt] = affectation.LibelleCourt;

        if (affectation.Duree > 0) {
          nouvelleLigne[ExcelSchemaNames.Columns.Duree] = affectation.Duree.ToString();
          nouvelleLigne[ExcelSchemaNames.Columns.TotalType] = affectation.Duree.ToString();
        }

        if (tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Infos))
          nouvelleLigne[ExcelSchemaNames.Columns.Infos] = affectation.Infos?.Trim() ?? string.Empty;

        tableGlobal.Rows.Add(nouvelleLigne);
        indexActivites[cleActivite] = nouvelleLigne;
        nbAjoutees++;

        if (string.Equals(affectation.Module,"ENER1-2",StringComparison.OrdinalIgnoreCase)) {
          Debug.WriteLine(
              "Ligne ajoutée ENER1-2 : " +
              affectation.LibelleCourt + " | " +
              affectation.Cours + " | " +
              affectation.Noms + " | " +
              affectation.Groupe + " | " +
              affectation.Duree + "h | ligne Excel " +
              affectation.SourceRow);
        }
      }

      tableGlobal.AcceptChanges();
    }

    internal static void EnrichirGroupesGlobalDepuisFeuilleSemestre_OleDb(DataTable grille,DataTable tableGlobal) {
      if (grille == null || tableGlobal == null)
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

      string nomFeuille = grille.TableName;
      string colonneLibelleCourt = TrouverNomColonneIgnoreCase(tableGlobal,ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      List<AffectationGroupeSemestre> affectations = LireAffectationsGroupesDepuisFeuilleSemestre_OleDb(grille);

      if (affectations.Count == 0)
        return;

      Dictionary<string,List<string>> groupesParCle = new(StringComparer.OrdinalIgnoreCase);

      foreach (AffectationGroupeSemestre affectation in affectations) {
        if (string.IsNullOrWhiteSpace(affectation.Noms))
          continue;

        if (EstLigneTechniqueGlobal(affectation.Noms))
          continue;

        string cle = ConstruireCleGroupeGlobal(
            affectation.Semestre,
            affectation.Module,
            affectation.LibelleCourt,
            affectation.Cours,
            affectation.Noms);

        if (!groupesParCle.TryGetValue(cle,out List<string>? groupes)) {
          groupes = new List<string>();
          groupesParCle.Add(cle,groupes);
        }

        groupes.Add(affectation.Groupe);
      }

      if (groupesParCle.Count == 0)
        return;

      Dictionary<string,int> indexParCle = new(StringComparer.OrdinalIgnoreCase);

      foreach (DataRow row in tableGlobal.Rows) {
        string semestre = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Semestre);

        if (!string.Equals(semestre,nomFeuille,StringComparison.OrdinalIgnoreCase))
          continue;

        string cours = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Cours);

        if (!string.Equals(cours,"TD",StringComparison.OrdinalIgnoreCase) && !string.Equals(cours,"TP",StringComparison.OrdinalIgnoreCase))
          continue;

        string noms = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Noms);

        if (EstLigneTechniqueGlobal(noms))
          continue;

        string module = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Module);
        string libelleCourt = GetRowValueIfColumnExists(row,colonneLibelleCourt);

        string cle = ConstruireCleGroupeGlobal(nomFeuille,module,libelleCourt,cours,noms);
        string cleUtilisee = cle;

        if (!groupesParCle.TryGetValue(cle,out List<string>? groupes)) {
          string moduleAvecSuffixeUn = AjouterSuffixeModuleUnSiNecessaire(module);
          string cleAvecSuffixeUn = ConstruireCleGroupeGlobal(nomFeuille,moduleAvecSuffixeUn,libelleCourt,cours,noms);

          cleUtilisee = cleAvecSuffixeUn;
          groupesParCle.TryGetValue(cleAvecSuffixeUn,out groupes);
        }

        if (groupes == null)
          continue;

        if (!indexParCle.TryGetValue(cleUtilisee,out int index))
          index = 0;

        if (index >= groupes.Count)
          continue;

        row[ExcelSchemaNames.Columns.Groupe] = groupes[index];
        indexParCle[cleUtilisee] = index + 1;
      }

      tableGlobal.AcceptChanges();
    }

    internal static void EnrichirInfosGlobalDepuisFeuilleSemestre_OleDb(DataTable grille,DataTable tableGlobal) {
      if (grille == null || tableGlobal == null)
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

      string nomFeuille = grille.TableName;
      string colonneLibelleCourt = TrouverNomColonneIgnoreCase(tableGlobal,ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      string colonneInfos = TrouverNomColonneIgnoreCase(tableGlobal,ExcelSchemaNames.Columns.Infos);

      if (string.IsNullOrWhiteSpace(colonneInfos)) {
        tableGlobal.Columns.Add(ExcelSchemaNames.Columns.Infos,typeof(string));
        colonneInfos = ExcelSchemaNames.Columns.Infos;
      }

      if (!TrouverCelluleTexteFeuilleSemestre(grille,"SALLES",out int ligneSalles,out int colonneSalles))
        return;

      int colMax = TrouverDerniereColonneGroupes(grille,ligneSalles,ligneSalles + 1,colonneSalles + 1);

      if (colMax < colonneSalles + 1)
        return;

      Dictionary<int,string> groupesTd = LireGroupesFeuilleSemestre(grille,ligneSalles,colonneSalles + 1,colMax,estTp: false);
      Dictionary<int,string> groupesTp = LireGroupesFeuilleSemestre(grille,ligneSalles + 1,colonneSalles + 1,colMax,estTp: true);

      Dictionary<string,List<DataRow>> lignesGlobalParIntervenant = new(StringComparer.OrdinalIgnoreCase);

      foreach (DataRow globalRow in tableGlobal.Rows) {
        string nomIntervenant = GetRowValueIfColumnExists(globalRow,ExcelSchemaNames.Columns.Noms).Trim();

        if (string.IsNullOrWhiteSpace(nomIntervenant))
          continue;

        if (!lignesGlobalParIntervenant.TryGetValue(nomIntervenant,out List<DataRow>? lignes)) {
          lignes = [];
          lignesGlobalParIntervenant.Add(nomIntervenant,lignes);
        }

        lignes.Add(globalRow);
      }

      Regex regexCodeService = new(@"^(R\d+-\d{2})-(CM|TD|TP|DS)$",RegexOptions.IgnoreCase | RegexOptions.Compiled);

      int rowMax = TrouverDerniereLigneUtileSemestre(grille);

      string moduleCourant = string.Empty;
      string infosCourantes = string.Empty;
      int lignesModifiees = 0;

      for (int row = 1;row <= rowMax;row++) {
        string celluleA = LireCelluleGrille(grille,row,1);
        string moduleDetecte = ExtraireModuleDepuisCelluleFeuilleSemestre(celluleA);

        if (EstCodeModuleFeuilleSemestreValide(moduleDetecte)) {
          moduleCourant = moduleDetecte;
          infosCourantes = string.Empty;
        }
        else if (!string.IsNullOrWhiteSpace(celluleA) && EstInfoFeuilleSemestreValide(celluleA)) {
          infosCourantes = celluleA;
        }

        string codeService = LireCelluleGrille(grille,row,2);
        Match match = regexCodeService.Match(codeService);

        if (!match.Success)
          continue;

        if (string.IsNullOrWhiteSpace(moduleCourant))
          continue;

        string cours = match.Groups[2].Value.ToUpperInvariant();

        if (cours != "TD" && cours != "TP")
          continue;

        if (string.IsNullOrWhiteSpace(infosCourantes))
          continue;

        Dictionary<int,string> groupes = cours == "TD" ? groupesTd : groupesTp;
        int dureeLigne = LireDureeDepuisLigneFeuilleSemestre(grille,row,colonneSalles);

        foreach (KeyValuePair<int,string> groupe in groupes) {
          string nomIntervenant = LireCelluleGrille(grille,row,groupe.Key);

          if (!EstNomIntervenantValidePourGroupe(nomIntervenant))
            continue;

          if (!lignesGlobalParIntervenant.TryGetValue(nomIntervenant,out List<DataRow>? lignesIntervenant))
            continue;

          foreach (DataRow globalRow in lignesIntervenant) {
            if (!LigneGlobalCorrespondInfoSemestre(
                    globalRow,
                    nomFeuille,
                    moduleCourant,
                    codeService,
                    cours,
                    nomIntervenant,
                    groupe.Value,
                    dureeLigne,
                    colonneLibelleCourt))
              continue;

            string infosActuelles = GetRowValueIfColumnExists(globalRow,colonneInfos);

            if (!string.IsNullOrWhiteSpace(infosActuelles) && infosActuelles != "0")
              continue;

            globalRow[colonneInfos] = infosCourantes;
            lignesModifiees++;
          }
        }
      }

      if (lignesModifiees > 0)
        Debug.WriteLine("INFOS enrichies OleDb depuis " + nomFeuille + " : " + lignesModifiees + " ligne(s)");

      tableGlobal.AcceptChanges();
    }

    private static DataTable LireGlobal_OleDb(OleDbConnection connection) {
      DataTable table = new(ExcelSchemaNames.Tables.NomTableGlobal);

      if (connection == null || connection.State != ConnectionState.Open)
        return table;

      using OleDbCommand commande = new($"SELECT * FROM [{ExcelSchemaNames.Tables.NomTableGlobal}$]",connection);
      using OleDbDataReader? lecteur = commande.ExecuteReader();

      if (lecteur == null || lecteur.FieldCount == 0)
        return table;

      if (!lecteur.Read())
        return table;

      for (int col = 0;col < lecteur.FieldCount;col++) {
        string nomColonne = lecteur.IsDBNull(col) ? string.Empty : Convert.ToString(lecteur.GetValue(col))?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(nomColonne))
          nomColonne = "Colonne_" + (col + 1);

        string nomUnique = nomColonne;
        int suffixe = 2;

        while (table.Columns.Contains(nomUnique))
          nomUnique = nomColonne + "_" + suffixe++;

        table.Columns.Add(nomUnique,typeof(string));
      }

      while (lecteur.Read()) {
        bool ligneVide = true;
        DataRow ligne = table.NewRow();

        for (int col = 0;col < lecteur.FieldCount;col++) {
          string valeur = lecteur.IsDBNull(col) ? string.Empty : Convert.ToString(lecteur.GetValue(col))?.Trim() ?? string.Empty;
          ligne[col] = valeur;

          if (!string.IsNullOrWhiteSpace(valeur))
            ligneVide = false;
        }

        if (!ligneVide)
          table.Rows.Add(ligne);
      }

      table.AcceptChanges();
      return table;
    }

    internal static void AjouterLignesManquantesDepuisFeuilleSemestre_Epplus(ExcelWorksheet feuille,DataTable tableGlobal) {
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

      string nomFeuille = feuille.Name;

      string colonneLibelleCourt =
          TrouverNomColonneIgnoreCase(
              tableGlobal,
              ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      Stopwatch chrono = Stopwatch.StartNew();

      List<AffectationGroupeSemestre> affectations =
          LireAffectationsGroupesDepuisFeuilleSemestre_Epplus(feuille);

      if (affectations.Count == 0)
        return;

      chrono.Restart();

      int prochainId =
          TrouverProchainIdGlobal(tableGlobal);

      AjouterColonnesSourceSemestreSiAbsentes(tableGlobal);

      ReconstruireLignesInfosSpecialesDepuisAffectations(
          tableGlobal,
          colonneLibelleCourt,
          affectations);

      Dictionary<string,DataRow> indexActivites =
          new(StringComparer.OrdinalIgnoreCase);

      Dictionary<string,DataRow> indexModeles =
          new(StringComparer.OrdinalIgnoreCase);

      foreach (DataRow row in tableGlobal.Rows) {
        string semestre = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Semestre);
        string module = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Module);
        string cours = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Cours);
        string noms = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Noms);
        string groupe = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Groupe);
        string libelleCourt = GetRowValueIfColumnExists(row,colonneLibelleCourt);

        string cleActivite =
            ConstruireCleActiviteGlobal(
                semestre,module,cours,noms,groupe,libelleCourt);

        if (!indexActivites.ContainsKey(cleActivite))
          indexActivites.Add(cleActivite,row);

        string cleModele =
            ConstruireCleModeleGlobal(
                semestre,module,cours,libelleCourt);

        if (!indexModeles.ContainsKey(cleModele))
          indexModeles.Add(cleModele,row);
      }

      int nbExistantes = 0;
      int nbAjoutees = 0;
      int nbSansModele = 0;


      foreach (AffectationGroupeSemestre affectation in affectations) {
        if (string.IsNullOrWhiteSpace(affectation.Noms))
          continue;

        if (EstLigneTechniqueGlobal(affectation.Noms))
          continue;

        string cleActivite =
            ConstruireCleActiviteGlobal(
                affectation.Semestre,
                affectation.Module,
                affectation.Cours,
                affectation.Noms,
                affectation.Groupe,
                affectation.LibelleCourt);

        if (indexActivites.TryGetValue(cleActivite,out DataRow? ligneExistante)) {
          if (ligneExistante.Table.Columns.Contains(ExcelSchemaNames.Columns.Infos))
            ligneExistante[ExcelSchemaNames.Columns.Infos] = affectation.Infos?.Trim() ?? string.Empty;

          if (ligneExistante.Table.Columns.Contains(ExcelSchemaNames.Columns.SourceSheet))
            ligneExistante[ExcelSchemaNames.Columns.SourceSheet] = affectation.SourceSheet;

          if (ligneExistante.Table.Columns.Contains(ExcelSchemaNames.Columns.SourceRow))
            ligneExistante[ExcelSchemaNames.Columns.SourceRow] = affectation.SourceRow.ToString();

          if (ligneExistante.Table.Columns.Contains(ExcelSchemaNames.Columns.SourceColumn))
            ligneExistante[ExcelSchemaNames.Columns.SourceColumn] = affectation.SourceColumn.ToString();

          nbExistantes++;
          continue;
        }

        string cleModele =
            ConstruireCleModeleGlobal(
                affectation.Semestre,
                affectation.Module,
                affectation.Cours,
                affectation.LibelleCourt);

        indexModeles.TryGetValue(cleModele,out DataRow? modele);

        if (modele == null) {
          nbSansModele++;
          continue;
        }
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

        tableGlobal.Rows.Add(nouvelleLigne);
        indexActivites[cleActivite] = nouvelleLigne;
        nbAjoutees++;
      }

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

      List<AffectationGroupeSemestre> affectationsSpeciales = affectations.Where(a => !string.IsNullOrWhiteSpace(a.Infos)).ToList();

      if (affectationsSpeciales.Count == 0)
        return;

      List<DataRow> lignesASupprimer = new List<DataRow>();

      foreach (DataRow row in tableGlobal.Rows) {
        string infosGlobal = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Infos);

        if (string.IsNullOrWhiteSpace(infosGlobal))
          continue;

        foreach (AffectationGroupeSemestre affectation in affectationsSpeciales) {
          if (!string.Equals(infosGlobal,affectation.Infos,StringComparison.OrdinalIgnoreCase)) {
            continue;
          }

          string semestreGlobal =
              GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Semestre);

          if (!string.Equals(semestreGlobal,affectation.Semestre,StringComparison.OrdinalIgnoreCase)) {
            continue;
          }

          string moduleGlobal = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Module);

          if (!ModuleCorrespond(moduleGlobal,affectation.Module)) {
            continue;
          }

          string coursGlobal = GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Cours);

          if (!string.Equals(coursGlobal,affectation.Cours,StringComparison.OrdinalIgnoreCase)) {
            continue;
          }

          string libelleCourtGlobal = GetRowValueIfColumnExists(row,colonneLibelleCourt);

          if (!string.Equals(libelleCourtGlobal,affectation.LibelleCourt,StringComparison.OrdinalIgnoreCase)) {
            continue;
          }

          lignesASupprimer.Add(row);
          break;
        }
      }

      foreach (DataRow row in lignesASupprimer) {
        tableGlobal.Rows.Remove(row);
      }
    }

    private static void ReconstruireNomsGlobalDepuisAffectationsS3_OleDb(DataTable tableGlobal,DataTable grilleS3) {
      if (tableGlobal == null || grilleS3 == null || grilleS3.Rows.Count == 0)
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Semestre) ||
          !tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Formation) ||
          !tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Cours) ||
          !tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Noms))
        return;

      string colonneLibelleCourt = TrouverNomColonneIgnoreCase(tableGlobal,ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      AjouterColonnesSourceSemestreSiAbsentes(tableGlobal);

      List<AffectationGroupeSemestre> affectations = LireAffectationsGroupesDepuisFeuilleS3Fifa_OleDb(grilleS3);

      if (affectations.Count == 0)
        return;

      Dictionary<string,List<AffectationGroupeSemestre>> affectationsParCle = new(StringComparer.OrdinalIgnoreCase);

      foreach (AffectationGroupeSemestre affectation in affectations) {
        if (string.IsNullOrWhiteSpace(affectation.Noms))
          continue;

        bool estFA = Regex.IsMatch(affectation.LibelleCourt,@"^R3\d{2}A-",RegexOptions.IgnoreCase);
        string formation = estFA ? "FA" : "FI";
        string libelleGlobal = NormaliserLibelleCourtS3(affectation.LibelleCourt);
        string cle = $"S3|{formation}|{affectation.Cours}|{libelleGlobal}";

        if (!affectationsParCle.TryGetValue(cle,out List<AffectationGroupeSemestre>? liste)) {
          liste = new List<AffectationGroupeSemestre>();
          affectationsParCle.Add(cle,liste);
        }

        liste.Add(affectation);
      }

      foreach (KeyValuePair<string,List<AffectationGroupeSemestre>> entree in affectationsParCle) {
        List<AffectationGroupeSemestre> listeAffectations = entree.Value;

        if (listeAffectations.Count == 0)
          continue;

        AffectationGroupeSemestre premiereAffectation = listeAffectations[0];

        bool estFA = Regex.IsMatch(premiereAffectation.LibelleCourt,@"^R3\d{2}A-",RegexOptions.IgnoreCase);
        string formation = estFA ? "FA" : "FI";
        string libelleGlobal = NormaliserLibelleCourtS3(premiereAffectation.LibelleCourt);

        List<DataRow> lignesGlobal = tableGlobal.Rows.Cast<DataRow>()
            .Where(row =>
                string.Equals(GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Semestre),"S3",StringComparison.OrdinalIgnoreCase) &&
                string.Equals(GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Formation),formation,StringComparison.OrdinalIgnoreCase) &&
                string.Equals(GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Cours),premiereAffectation.Cours,StringComparison.OrdinalIgnoreCase) &&
                string.Equals(GetRowValueIfColumnExists(row,colonneLibelleCourt),libelleGlobal,StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (lignesGlobal.Count == 0) {
          Debug.WriteLine($"[S3 FIFA] Aucun modèle Global : {formation} | {premiereAffectation.Cours} | {libelleGlobal}");
          continue;
        }

        DataRow ligneModele = lignesGlobal[0];

        for (int i = 0;i < listeAffectations.Count;i++) {
          AffectationGroupeSemestre affectation = listeAffectations[i];
          DataRow ligne;

          if (i < lignesGlobal.Count) {
            ligne = lignesGlobal[i];
          }
          else {
            ligne = tableGlobal.NewRow();
            ligne.ItemArray = (object[])ligneModele.ItemArray.Clone();
            tableGlobal.Rows.Add(ligne);
          }

          ligne[ExcelSchemaNames.Columns.Noms] = affectation.Noms;

          if (tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Groupe))
            ligne[ExcelSchemaNames.Columns.Groupe] = affectation.Groupe;

          ligne[ExcelSchemaNames.Columns.SourceSheet] = affectation.SourceSheet;
          ligne[ExcelSchemaNames.Columns.SourceRow] = affectation.SourceRow.ToString();
          ligne[ExcelSchemaNames.Columns.SourceColumn] = affectation.SourceColumn.ToString();
        }

        for (int i = lignesGlobal.Count - 1;i >= listeAffectations.Count;i--)
          tableGlobal.Rows.Remove(lignesGlobal[i]);
      }

      tableGlobal.AcceptChanges();
    }

    private static void ReconstruireNomsGlobalDepuisAffectations_OleDb(DataTable tableGlobal,params DataTable[] grillesSemestres) {
      if (tableGlobal == null || grillesSemestres == null || grillesSemestres.Length == 0)
        return;

      if (!tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Semestre) ||
          !tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Module) ||
          !tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Cours) ||
          !tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Noms))
        return;

      string colonneLibelleCourt = TrouverNomColonneIgnoreCase(tableGlobal,ExcelSchemaNames.Columns.LibelleCourt);

      if (string.IsNullOrWhiteSpace(colonneLibelleCourt))
        return;

      AjouterColonnesSourceSemestreSiAbsentes(tableGlobal);

      List<AffectationGroupeSemestre> affectations = new();

      foreach (DataTable grille in grillesSemestres) {
        if (grille == null || grille.Rows.Count == 0)
          continue;

        affectations.AddRange(LireAffectationsGroupesDepuisFeuilleSemestre_OleDb(grille));
      }

      if (affectations.Count == 0)
        return;

      Dictionary<string,List<AffectationGroupeSemestre>> affectationsParCle = new(StringComparer.OrdinalIgnoreCase);

      foreach (AffectationGroupeSemestre affectation in affectations) {
        if (string.IsNullOrWhiteSpace(affectation.Noms))
          continue;

        string cle = ConstruireCleModeleGlobal(
            affectation.Semestre,
            affectation.Module,
            affectation.Cours,
            affectation.LibelleCourt);

        if (!affectationsParCle.TryGetValue(cle,out List<AffectationGroupeSemestre>? liste)) {
          liste = new List<AffectationGroupeSemestre>();
          affectationsParCle.Add(cle,liste);
        }

        liste.Add(affectation);
      }

      int activitesReconstruites = 0;
      int lignesReconstruites = 0;
      int lignesAjoutees = 0;
      int lignesSupprimees = 0;
      int activitesSansModele = 0;

      foreach (KeyValuePair<string,List<AffectationGroupeSemestre>> entree in affectationsParCle) {
        List<AffectationGroupeSemestre> listeAffectations = entree.Value;

        if (listeAffectations.Count == 0)
          continue;

        AffectationGroupeSemestre premiereAffectation = listeAffectations[0];

        List<DataRow> lignesGlobal = tableGlobal.Rows.Cast<DataRow>()
            .Where(row =>
                string.Equals(
                    GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Semestre),
                    premiereAffectation.Semestre,
                    StringComparison.OrdinalIgnoreCase) &&
                ModuleCorrespond(
                    GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Module),
                    premiereAffectation.Module) &&
                string.Equals(
                    GetRowValueIfColumnExists(row,ExcelSchemaNames.Columns.Cours),
                    premiereAffectation.Cours,
                    StringComparison.OrdinalIgnoreCase) &&
                string.Equals(
                    GetRowValueIfColumnExists(row,colonneLibelleCourt),
                    premiereAffectation.LibelleCourt,
                    StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (lignesGlobal.Count == 0) {
          activitesSansModele++;
          continue;
        }

        DataRow ligneModele = lignesGlobal[0];

        for (int i = 0;i < listeAffectations.Count;i++) {
          AffectationGroupeSemestre affectation = listeAffectations[i];
          DataRow ligne;

          if (i < lignesGlobal.Count) {
            ligne = lignesGlobal[i];
          }
          else {
            ligne = tableGlobal.NewRow();
            ligne.ItemArray = (object[])ligneModele.ItemArray.Clone();
            tableGlobal.Rows.Add(ligne);
            lignesAjoutees++;
          }

          ligne[ExcelSchemaNames.Columns.Noms] = affectation.Noms;

          if (tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Groupe))
            ligne[ExcelSchemaNames.Columns.Groupe] = affectation.Groupe;

          ligne[ExcelSchemaNames.Columns.SourceSheet] = affectation.SourceSheet;
          ligne[ExcelSchemaNames.Columns.SourceRow] = affectation.SourceRow.ToString();
          ligne[ExcelSchemaNames.Columns.SourceColumn] = affectation.SourceColumn.ToString();

          if (tableGlobal.Columns.Contains(ExcelSchemaNames.Columns.Infos))
            ligne[ExcelSchemaNames.Columns.Infos] = affectation.Infos?.Trim() ?? string.Empty;

          lignesReconstruites++;
        }

        int nbASupprimer = lignesGlobal.Count - listeAffectations.Count;

        for (int i = lignesGlobal.Count - 1;i >= listeAffectations.Count;i--) {
          tableGlobal.Rows.Remove(lignesGlobal[i]);
          lignesSupprimees++;
        }

        activitesReconstruites++;
      }

      tableGlobal.AcceptChanges();
    }

    private static List<AffectationGroupeSemestre> LireAffectationsGroupesDepuisFeuilleS3Fifa_OleDb(DataTable grille) {
      List<AffectationGroupeSemestre> affectations = new();

      if (grille == null || grille.Rows.Count == 0 || grille.Columns.Count == 0)
        return affectations;

      string nomFeuille = grille.TableName;

      if (!TrouverCelluleTexteFeuilleSemestre(grille,"SALLES",out int ligneSalles,out int colonneSalles))
        return affectations;

      Dictionary<int,string> groupes = new();

      for (int col = colonneSalles + 1;col <= grille.Columns.Count;col++) {
        string entete = LireCelluleGrille(grille,ligneSalles + 1,col);

        if (Regex.IsMatch(entete,@"^(TD|TP)\d+$",RegexOptions.IgnoreCase))
          groupes[col] = entete.ToUpperInvariant();
      }

      if (groupes.Count == 0)
        return affectations;

      int rowMax = TrouverDerniereLigneUtileSemestre(grille);
      Regex regexCodeService = new(@"^(R3\d{2})(A?)-(CM|TD|TP|DS)$",RegexOptions.IgnoreCase | RegexOptions.Compiled);

      for (int row = 1;row <= rowMax;row++) {
        string celluleModule = LireCelluleGrille(grille,row,1);
        string codeService = LireCelluleGrille(grille,row,5);

        Match match = regexCodeService.Match(codeService);

        if (!match.Success)
          continue;

        string cours = match.Groups[3].Value.ToUpperInvariant();

        if (cours != "TD" && cours != "TP")
          continue;

        int dureeLigne = LireDureeDepuisLigneFeuilleSemestre(grille,row,colonneSalles);

        foreach (KeyValuePair<int,string> groupe in groupes) {
          if (cours == "TD" && !groupe.Value.StartsWith("TD",StringComparison.OrdinalIgnoreCase))
            continue;

          if (cours == "TP" && !groupe.Value.StartsWith("TP",StringComparison.OrdinalIgnoreCase))
            continue;

          string nomIntervenant = LireCelluleGrille(grille,row,groupe.Key);

          if (!EstNomIntervenantValidePourGroupe(nomIntervenant))
            continue;

          affectations.Add(new AffectationGroupeSemestre {
            Semestre = nomFeuille,
            Module = celluleModule,
            LibelleCourt = codeService,
            Cours = cours,
            Noms = nomIntervenant,
            Groupe = groupe.Value,
            Duree = dureeLigne,
            Infos = string.Empty,
            SourceSheet = nomFeuille,
            SourceRow = row,
            SourceColumn = groupe.Key
          });
        }
      }

      return affectations;
    }

    private static List<AffectationGroupeSemestre> LireAffectationsGroupesDepuisFeuilleSemestre_OleDb(DataTable grille) {
      List<AffectationGroupeSemestre> affectations = new();

      if (grille == null || grille.Rows.Count == 0 || grille.Columns.Count == 0)
        return affectations;

      string nomFeuille = grille.TableName;
      string infosCourantes = string.Empty;

      if (!TrouverCelluleTexteFeuilleSemestre(grille,"SALLES",out int ligneSalles,out int colonneSalles))
        return affectations;

      int colMax = TrouverDerniereColonneGroupes(grille,ligneSalles,ligneSalles + 1,colonneSalles + 1);

      if (colMax < colonneSalles + 1)
        return affectations;

      Dictionary<int,string> groupesTd = LireGroupesFeuilleSemestre(grille,ligneSalles,colonneSalles + 1,colMax,estTp: false);
      Dictionary<int,string> groupesTp = LireGroupesFeuilleSemestre(grille,ligneSalles + 1,colonneSalles + 1,colMax,estTp: true);

      int colonnePromo = 0;

      for (int col = colonneSalles + 1;col <= colMax;col++) {
        string groupe = LireCelluleGrille(grille,ligneSalles,col);

        if (string.Equals(groupe,"Promo",StringComparison.OrdinalIgnoreCase)) {
          colonnePromo = col;
          break;
        }
      }

      int rowMax = TrouverDerniereLigneUtileSemestre(grille);
      Regex regexCodeService = new(@"^(R\d+-\d{2})-(CM|TD|TP|DS)$",RegexOptions.IgnoreCase | RegexOptions.Compiled);

      string moduleCourant = string.Empty;
      string codeServiceCourant = string.Empty;
      string coursCourant = string.Empty;

      for (int row = 1;row <= rowMax;row++) {
        string celluleModule = LireCelluleGrille(grille,row,1);
        string codeService = LireCelluleGrille(grille,row,2);
        Match match = regexCodeService.Match(codeService);

        string moduleDetecte = ExtraireModuleDepuisCelluleFeuilleSemestre(celluleModule);
        if (!string.IsNullOrWhiteSpace(moduleDetecte)) {
          moduleCourant = moduleDetecte;
          infosCourantes = string.Empty;
        }
        else if (match.Success && !string.IsNullOrWhiteSpace(celluleModule)) {
          infosCourantes = celluleModule;
        }
        if (match.Success) {
          codeServiceCourant = codeService;
          coursCourant = match.Groups[2].Value.ToUpperInvariant();
        }
        else {
          if (!string.IsNullOrWhiteSpace(codeService))
            continue;
        }

        if (string.IsNullOrWhiteSpace(codeServiceCourant) || string.IsNullOrWhiteSpace(moduleCourant))
          continue;

        int dureeLigne = LireDureeDepuisLigneFeuilleSemestre(grille,row,colonneSalles);

        if (coursCourant == "CM" || coursCourant == "DS") {
          if (colonnePromo <= 0)
            continue;

          string nomIntervenant = LireCelluleGrille(grille,row,colonnePromo);

          if (!EstNomIntervenantValidePourGroupe(nomIntervenant))
            continue;

          affectations.Add(new AffectationGroupeSemestre {
            Semestre = nomFeuille,
            Module = moduleCourant,
            LibelleCourt = codeServiceCourant,
            Cours = coursCourant,
            Noms = nomIntervenant,
            Groupe = "Promo",
            Duree = dureeLigne,
            Infos = infosCourantes,
            SourceSheet = nomFeuille,
            SourceRow = row,
            SourceColumn = colonnePromo
          });

          continue;
        }

        if (coursCourant != "TD" && coursCourant != "TP")
          continue;

        Dictionary<int,string> groupes = coursCourant == "TD" ? groupesTd : groupesTp;

        foreach (KeyValuePair<int,string> groupe in groupes) {
          string nomIntervenant = LireCelluleGrille(grille,row,groupe.Key);

          if (!EstNomIntervenantValidePourGroupe(nomIntervenant))
            continue;

          affectations.Add(new AffectationGroupeSemestre {
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
          });
        }
      }

      return affectations;
    }


    private static List<AffectationGroupeSemestre> LireAffectationsGroupesDepuisFeuilleSemestre_Epplus(ExcelWorksheet feuille) {
      List<AffectationGroupeSemestre> affectations = new();

      if (feuille.Dimension == null)
        return affectations;

      string nomFeuille = feuille.Name;
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
          new(
              @"^(R\d+-\d{2})-(CM|TD|TP|DS)$",
              RegexOptions.IgnoreCase | RegexOptions.Compiled);

      string moduleCourant = string.Empty;
      string codeServiceCourant = string.Empty;
      string coursCourant = string.Empty;

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
              new() {
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

          affectations.Add(affectation);
        }
      }

      foreach (AffectationGroupeSemestre a in affectations.Where(a =>
                   a.Module.Equals("AN1",StringComparison.OrdinalIgnoreCase))) {
        Debug.WriteLine($"[AN1 SOURCE] {a.SourceSheet}!R{a.SourceRow}C{a.SourceColumn} | Module={a.Module} | Libelle={a.LibelleCourt} | Cours={a.Cours} | Groupe={a.Groupe} | Nom={a.Noms}");
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

    private static string ConstruireCleActiviteGlobal(
    string semestre,
    string module,
    string cours,
    string noms,
    string groupe,
    string libelleCourt) {

      return string.Join("|",
          semestre.Trim(),
          AjouterSuffixeModuleUnSiNecessaire(module.Trim()),
          cours.Trim(),
          noms.Trim(),
          groupe.Trim(),
          libelleCourt.Trim());
    }

    private static string ConstruireCleModeleGlobal(
        string semestre,
        string module,
        string cours,
        string libelleCourt) {

      return string.Join("|",
          semestre.Trim(),
          AjouterSuffixeModuleUnSiNecessaire(module.Trim()),
          cours.Trim(),
          libelleCourt.Trim());
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

    private static string LireInfosDepuisLigneFeuilleSemestre(DataTable grille,int row,int colonneSalles) {
      if (grille == null)
        return string.Empty;

      string valeur = LireCelluleGrille(grille,row,1);

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

      string moduleDetecte = ExtraireModuleDepuisCelluleFeuilleSemestre(valeur);

      if (!string.IsNullOrWhiteSpace(moduleDetecte))
        return string.Empty;

      return valeur;
    }


    private static int LireDureeDepuisLigneFeuilleSemestre(DataTable grille,int row,int colonneSalles) {
      if (grille == null)
        return 0;

      int[] colonnesCandidates = { colonneSalles - 2,colonneSalles - 1,colonneSalles - 3 };

      foreach (int col in colonnesCandidates) {
        if (col <= 0)
          continue;

        string valeur = LireCelluleGrille(grille,row,col);

        if (string.IsNullOrWhiteSpace(valeur))
          continue;

        valeur = valeur.Replace(",",".");

        if (decimal.TryParse(valeur,System.Globalization.NumberStyles.Any,System.Globalization.CultureInfo.InvariantCulture,out decimal resultat))
          return Convert.ToInt32(Math.Round(resultat));
      }

      return 0;
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

    internal static void MiseAJourDureesAttenduesGlobal_OleDb(
        ClasseExcel fichierDeService,
        string codeCourtModule,
        string valeurCM,
        string valeurTD,
        string valeurTP) {

      if (fichierDeService == null || string.IsNullOrWhiteSpace(fichierDeService.CheminFichier) || !File.Exists(fichierDeService.CheminFichier))
        return;

      Stopwatch chronoTotal = Stopwatch.StartNew();
      Stopwatch chrono = Stopwatch.StartNew();
      int lignesModifiees = 0;

      string connexion =
        $"Provider=Microsoft.ACE.OLEDB.12.0;" +
        $"Data Source={fichierDeService.CheminFichier};" +
        $"Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=0\";";

      using (OleDbConnection connection = new(connexion)) {
        connection.Open();

        Debug.WriteLine($"[CHRONO OLEDB] Ouverture connexion : {chrono.ElapsedMilliseconds} ms");

        chrono.Restart();

        lignesModifiees = MiseAJourDureesAttenduesOleDb(
            connection,
            codeCourtModule,
            valeurCM,
            valeurTD,
            valeurTP);

        Debug.WriteLine($"[CHRONO OLEDB] UPDATE : {chrono.ElapsedMilliseconds} ms");
      }

      Debug.WriteLine($"[CHRONO OLEDB] Fermeture + TOTAL : {chronoTotal.ElapsedMilliseconds} ms");
      Debug.WriteLine($"[CHRONO OLEDB] Lignes modifiées : {lignesModifiees}");
    }

    private static int MiseAJourDureesAttenduesOleDb(
        OleDbConnection connection,
        string codeCourtModule,
        string valeurCM,
        string valeurTD,
        string valeurTP) {

      const string requete =
        "UPDATE [Global$] " +
        "SET [DUREE_ATTENDUE] = IIf([TYPE]='CM', ?, IIf([TYPE]='TD', ?, ?)) " +
        "WHERE [LIBELLE COURT] LIKE ? " +
        "AND [TYPE] IN ('CM','TD','TP')";

      using OleDbCommand commande = new(requete,connection);

      commande.Parameters.AddWithValue("@p1",valeurCM);
      commande.Parameters.AddWithValue("@p2",valeurTD);
      commande.Parameters.AddWithValue("@p3",valeurTP);
      commande.Parameters.AddWithValue("@p4",codeCourtModule + "-%");

      return commande.ExecuteNonQuery();
    }

    internal static void MiseAJourDureesAttenduesGlobal_Epplus(ClasseExcel fichierDeService,string codeCourtModule,string valeurCM,string valeurTD,string valeurTP) {
      if (fichierDeService == null || string.IsNullOrWhiteSpace(fichierDeService.CheminFichier) || !File.Exists(fichierDeService.CheminFichier))
        return;

      Stopwatch chrono = Stopwatch.StartNew();
      ClasseEpplus.ConfigureEpplusLicense();

      using ExcelPackage package = new(new FileInfo(fichierDeService.CheminFichier));

      Debug.WriteLine($"[CHRONO SERVICE] Ouverture package : {chrono.ElapsedMilliseconds} ms");

      chrono.Restart();

      ExcelWorksheet? feuille = package.Workbook.Worksheets[ExcelSchemaNames.Tables.NomTableGlobal];

      Debug.WriteLine($"[CHRONO SERVICE] Accès feuille Global : {chrono.ElapsedMilliseconds} ms");

      chrono.Restart();

      int colonneLibelleCourt = ClasseEpplus.FindColumnIndexByName_Epplus(feuille,"LIBELLE COURT");
      int colonneCours = ClasseEpplus.FindColumnIndexByName_Epplus(feuille,"TYPE");
      int colonneDureeAttendue = ClasseEpplus.FindColumnIndexByName_Epplus(feuille,"DUREE_ATTENDUE");

      Debug.WriteLine($"[CHRONO SERVICE] Recherche colonnes : {chrono.ElapsedMilliseconds} ms");

      if (colonneLibelleCourt <= 0 || colonneCours <= 0 || colonneDureeAttendue <= 0)
        return;

      chrono.Restart();

      int lignesParcourues = 0;
      int lignesModule = 0;
      int cellulesModifiees = 0;

      for (int row = 2;row <= feuille.Dimension.End.Row;row++) {
        lignesParcourues++;

        string libelleCourt = feuille.Cells[row,colonneLibelleCourt].Text.Trim();

        if (!libelleCourt.StartsWith(codeCourtModule + "-",StringComparison.OrdinalIgnoreCase))
          continue;

        lignesModule++;

        string cours = feuille.Cells[row,colonneCours].Text.Trim().ToUpperInvariant();

        string valeur = cours switch {
          "CM" => valeurCM,
          "TD" => valeurTD,
          "TP" => valeurTP,
          _ => string.Empty
        };

        if (!string.IsNullOrWhiteSpace(valeur)) {
          feuille.Cells[row,colonneDureeAttendue].Value = valeur;
          cellulesModifiees++;
        }
      }

      Debug.WriteLine($"[CHRONO SERVICE] Parcours + écriture : {chrono.ElapsedMilliseconds} ms");
      Debug.WriteLine($"[CHRONO SERVICE] Lignes parcourues : {lignesParcourues} | Lignes module : {lignesModule} | Cellules modifiées : {cellulesModifiees}");

      chrono.Restart();

      package.Save();

      Debug.WriteLine($"[CHRONO SERVICE] Save : {chrono.ElapsedMilliseconds} ms");
    }

    private static int TrouverDerniereColonneGroupes(DataTable grille,int rowTD,int rowTP,int colDebut) {
      if (grille == null || rowTD <= 0 || rowTP <= 0 || colDebut <= 0)
        return colDebut;

      int colMax = Math.Min(grille.Columns.Count,MaxColonnesRechercheGroupes);
      int derniereColonne = colDebut;
      int colonnesVidesConsecutives = 0;

      for (int col = colDebut;col <= colMax;col++) {
        string groupeTD = LireCelluleGrille(grille,rowTD,col);
        string groupeTP = LireCelluleGrille(grille,rowTP,col);

        if (!string.IsNullOrWhiteSpace(groupeTD) || !string.IsNullOrWhiteSpace(groupeTP)) {
          derniereColonne = col;
          colonnesVidesConsecutives = 0;
        }
        else {
          colonnesVidesConsecutives++;

          if (colonnesVidesConsecutives >= StopColonnesVidesConsecutives)
            break;
        }
      }

      return derniereColonne;
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

    private static int TrouverDerniereLigneUtileSemestre(DataTable grille) {
      if (grille == null || grille.Rows.Count == 0)
        return 1;

      int ligneLimite = Math.Min(grille.Rows.Count,MaxLignesRechercheSemestre);
      int derniereLigne = 1;

      for (int row = 1;row <= ligneLimite;row++) {
        string colonneA = LireCelluleGrille(grille,row,1);
        string colonneB = LireCelluleGrille(grille,row,2);

        if (!string.IsNullOrWhiteSpace(colonneA) || !string.IsNullOrWhiteSpace(colonneB))
          derniereLigne = row;
      }

      return derniereLigne;
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

    private static Dictionary<int,string> LireGroupesFeuilleSemestre(DataTable grille,int row,int colDebut,int colFin,bool estTp) {
      Dictionary<int,string> groupes = new();

      if (grille == null || row <= 0)
        return groupes;

      colFin = Math.Min(colFin,grille.Columns.Count);

      for (int col = colDebut;col <= colFin;col++) {
        string valeur = LireCelluleGrille(grille,row,col);

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
      if (!Regex.IsMatch(texte,@"^[A-Z]{2,}[A-Z0-9]*\d+(?:-\d+)?$",RegexOptions.IgnoreCase)) {
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
      if (string.IsNullOrWhiteSpace(moduleGlobal) || string.IsNullOrWhiteSpace(moduleFeuille))
        return false;

      string global = moduleGlobal.Trim();
      string feuille = moduleFeuille.Trim();

      if (string.Equals(global,feuille,StringComparison.OrdinalIgnoreCase))
        return true;

      string globalSansQualificatif = Regex.Replace(global,@"\s*\([^)]*\)\s*$",string.Empty).Trim();
      string feuilleSansQualificatif = Regex.Replace(feuille,@"\s*\([^)]*\)\s*$",string.Empty).Trim();

      if (string.Equals(globalSansQualificatif,feuilleSansQualificatif,StringComparison.OrdinalIgnoreCase))
        return true;

      string globalSuffixe = AjouterSuffixeModuleUnSiNecessaire(globalSansQualificatif);
      string feuilleSuffixe = AjouterSuffixeModuleUnSiNecessaire(feuilleSansQualificatif);

      if (string.Equals(globalSuffixe,feuilleSansQualificatif,StringComparison.OrdinalIgnoreCase))
        return true;

      if (string.Equals(globalSansQualificatif,feuilleSuffixe,StringComparison.OrdinalIgnoreCase))
        return true;

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

    private static int AjouterLigneNomTableGlobal_Epplus(ClasseExcel fichierDialogue,string nomBase,DataRow sourceRow,string id,string module,string coursAjoute = "",string nomAjoute = "") {
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
    internal static DataSet MiseAJourLigneModule_fichierDeServiceExcel(ClasseExcel fichierDialogue,DataSet dataSetExcel,string nomBase,string module,bool premierResultatSeulement = false) {
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
          MiseAJourHeuresModuleDepuisDataTable_Epplus(fichierDialogue,dataTable,nomBase,module,premierResultatSeulement);

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

    private static int MiseAJourHeuresModuleDepuisDataTable_Epplus(ClasseExcel fichierDialogue,DataTable dataTable,string nomBase,string module,bool premierResultatSeulement = false) {
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

        //int lignesModifiees =
        //    MiseAJourHeuresModuleParModuleEtUe_Epplus(
        //        fichierDialogue,
        //        nomBase,
        //        module,
        //        semestre!,
        //        intCm,
        //        intTd,
        //        intTp,
        //        premierResultatSeulement
        //    );

        int lignesModifiees =
            MiseAJourHeuresModuleParModuleEtUe_OleDb(
                fichierDialogue,
                nomBase,
                module,
                semestre!,
                intCm,
                intTd,
                intTp);

        return lignesModifiees;
      }

      return 0;
    }

    internal static void MiseAJourModuleSelection_DataTable(
    DataSet dataSet,
    string module,
    string valeurCM,
    string valeurTD,
    string valeurTP) {

      DataTable? table = dataSet.Tables[ExcelSchemaNames.Tables.Module];
      if (table == null)
        return;

      foreach (DataRow ligne in table.Rows) {
        if (!string.Equals(
            ligne[ExcelSchemaNames.Columns.Module]?.ToString()?.Trim(),
            module,
            StringComparison.OrdinalIgnoreCase))
          continue;

        ligne[ExcelSchemaNames.Columns.CM] = valeurCM;
        ligne[ExcelSchemaNames.Columns.TD] = valeurTD;
        ligne[ExcelSchemaNames.Columns.TP] = valeurTP;
        return;
      }
    }

    internal static int MiseAJourHeuresModuleParModuleEtUe_OleDb(ClasseExcel fichierDialogue,string nomBase,string module,string semestre,int cm,int td,int tp) {
      if (fichierDialogue == null || string.IsNullOrWhiteSpace(fichierDialogue.CheminFichier) || !File.Exists(fichierDialogue.CheminFichier))
        return 0;

      if (string.IsNullOrWhiteSpace(module) || string.IsNullOrWhiteSpace(semestre))
        return 0;

      Stopwatch chrono = Stopwatch.StartNew();

      try {
        string connexion =
          $"Provider=Microsoft.ACE.OLEDB.12.0;" +
          $"Data Source={fichierDialogue.CheminFichier};" +
          $"Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=0\";";

        int lignesModifiees;

        using (OleDbConnection connection = new(connexion)) {
          connection.Open();

          Debug.WriteLine($"[CHRONO OLEDB MODULE] Ouverture : {chrono.ElapsedMilliseconds} ms");
          chrono.Restart();

          string requete =
            $"UPDATE [{nomBase}$] " +
            $"SET [CM] = ?, [TD] = ?, [TP] = ? " +
            $"WHERE [MODULE] = ? AND [Semestre] = ?";

          using OleDbCommand commande = new(requete,connection);

          commande.Parameters.AddWithValue("@p1",cm);
          commande.Parameters.AddWithValue("@p2",td);
          commande.Parameters.AddWithValue("@p3",tp);
          commande.Parameters.AddWithValue("@p4",module);
          commande.Parameters.AddWithValue("@p5",semestre);

          lignesModifiees = commande.ExecuteNonQuery();

          Debug.WriteLine($"[CHRONO OLEDB MODULE] UPDATE : {chrono.ElapsedMilliseconds} ms");
        }

        Debug.WriteLine($"[CHRONO OLEDB MODULE] TOTAL : {chrono.ElapsedMilliseconds} ms");
        Debug.WriteLine($"[CHRONO OLEDB MODULE] Lignes modifiées : {lignesModifiees}");

        return lignesModifiees;
      }
      catch (Exception ex) {
        ShowOleDbError(
          ex,
          "MiseAJourHeuresModuleParModuleEtUe_OleDb",
          "Mise à jour OleDb Module : " + module + " / " + semestre);

        return 0;
      }
    }


    internal static int MiseAJourHeuresModuleParModuleEtUe_Epplus(ClasseExcel fichierDialogue,string nomBase,string module,string semestre,int cm,int td,int tp,bool premierResultatSeulement = false) {
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
            },
            premierResultatSeulement: premierResultatSeulement
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

      //package.Workbook.Calculate();

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
              @"^(?:(?:'(?<sheetq>[^']+)'|(?<sheet>[^'!]+))!)?(?<col>[A-Z]{1,3})(?<row1>[0-9]+)$",
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
          match.Groups["row1"].Value;

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
          //package.Workbook.Calculate();
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
                @"^(?:(?:'(?<sheetq>[^']+)'|(?<sheet>[^'!]+))!)?(?<col>[A-Z]{1,3})(?<row1>[0-9]+)$",
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
            match.Groups["row1"].Value;

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
    internal protected static DataTable ChangementBaseDeDonnéesFicheModule(DataSet dataSetExcelSélection,Label texteMatière,TextBox nouvelleValeurCM,
      TextBox nouvelleValeurTD,TextBox nouvelleValeurTP) {
      if (dataSetExcelSélection == null || !dataSetExcelSélection.Tables.Contains(ExcelSchemaNames.Tables.Module))
        return null!;

      DataTable tableDeDonnées = dataSetExcelSélection.Tables[ExcelSchemaNames.Tables.Module]!;

      string module = texteMatière.Content?.ToString()?.Trim() ?? string.Empty;

      if (string.IsNullOrWhiteSpace(module))
        return tableDeDonnées;

      tableDeDonnées.BeginLoadData();

      try {
        foreach (DataRow ligne in tableDeDonnées.Rows) {
          string moduleLigne = ligne[ExcelSchemaNames.Columns.Module]?.ToString()?.Trim() ?? string.Empty;

          if (!string.Equals(moduleLigne,module,StringComparison.OrdinalIgnoreCase))
            continue;

          ligne[ExcelSchemaNames.Columns.CM] = nouvelleValeurCM.Text.Trim();
          ligne[ExcelSchemaNames.Columns.TD] = nouvelleValeurTD.Text.Trim();
          ligne[ExcelSchemaNames.Columns.TP] = nouvelleValeurTP.Text.Trim();
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

      if (tableDeDonnée == ExcelSchemaNames.Tables.NomTableGlobal) {
        filtreListe = AjouterFiltreMetierGlobal(filtreListe,tableSource);
        InitialiserReferencesSources(tableSource);
      }

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

    internal static bool SupprimerAffectationSource(
    ClasseExcel fichierDeService,
    string sourceSheet,
    int sourceRow,
    int sourceColumn) {

      if (fichierDeService == null ||
          string.IsNullOrWhiteSpace(fichierDeService.CheminFichier) ||
          !File.Exists(fichierDeService.CheminFichier) ||
          string.IsNullOrWhiteSpace(sourceSheet) ||
          sourceRow <= 0 ||
          sourceColumn <= 0)
        return false;

      try {
        using ExcelPackage package = new(new FileInfo(fichierDeService.CheminFichier));

        ExcelWorksheet? feuille = package.Workbook.Worksheets[sourceSheet];

        if (feuille == null)
          return false;

        feuille.Cells[sourceRow,sourceColumn].Value = null;

        package.Save();

        return true;
      }
      catch (Exception ex) {
        MessageBox.Show(
            "Impossible de supprimer l'affectation dans le fichier Excel.\n\n" +
            ex.Message,
            "Suppression",
            MessageBoxButton.OK,
            MessageBoxImage.Error);

        return false;
      }
    }

    internal static void InitialiserReferencesSources(DataTable tableSource) {
      if (!tableSource.Columns.Contains(ExcelSchemaNames.Columns.SourceReferences))
        tableSource.Columns.Add(ExcelSchemaNames.Columns.SourceReferences,typeof(Dictionary<string,ExcelSourceReference>));

      foreach (DataRow row in tableSource.Rows) {
        string module = row[ExcelSchemaNames.Columns.Module]?.ToString()?.Trim() ?? string.Empty;
        string groupe = row[ExcelSchemaNames.Columns.Groupe]?.ToString()?.Trim() ?? string.Empty;
        string sourceSheet = row[ExcelSchemaNames.Columns.SourceSheet]?.ToString()?.Trim() ?? string.Empty;
        string sourceRowText = row[ExcelSchemaNames.Columns.SourceRow]?.ToString()?.Trim() ?? string.Empty;
        string sourceColumnText = row[ExcelSchemaNames.Columns.SourceColumn]?.ToString()?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(groupe) ||
            string.IsNullOrWhiteSpace(sourceSheet) ||
            !int.TryParse(sourceRowText,out int sourceRow) ||
            !int.TryParse(sourceColumnText,out int sourceColumn))
          continue;

        Dictionary<string,ExcelSourceReference> references = new() {
          [groupe] = new ExcelSourceReference {
            Module = module,
            SourceSheet = sourceSheet,
            SourceRow = sourceRow,
            SourceColumn = sourceColumn
          }
        };

        row[ExcelSchemaNames.Columns.SourceReferences] = references;
      }
    }


    internal static int AppliquerModificationsCellulesSourcesSemestre_Epplus(ClasseExcel fichierDialogue,IEnumerable<ExcelSourceCellChange> modifications) {

      ClasseEpplus.ConfigureEpplusLicense();

      if (fichierDialogue == null)
        return 0;

      if (string.IsNullOrWhiteSpace(fichierDialogue.CheminFichier)) {
        return 0;
      }



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

      int cellulesModifiees = 0;

      using ExcelPackage package = new(fichier);

      foreach (ExcelSourceCellChange modification in modificationsUniques.Values) {
        ExcelWorksheet? feuille =
            package.Workbook.Worksheets[modification.SourceSheet];

        if (feuille == null)
          continue;

        feuille.Cells[modification.SourceRow,modification.SourceColumn].Value = modification.NewValue;

        cellulesModifiees++;
      }

      if (cellulesModifiees > 0) {
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

    internal static bool AjouterIntervenantSourceSemestre_Epplus(
      ClasseExcel fichierDialogue,
      string sourceSheet,
      int sourceRow,
      string cours,
      string groupe,
      string nomIntervenant,
      out int nouvelleSourceRow,
      out int nouvelleSourceColumn) {

      nouvelleSourceRow = 0;
      nouvelleSourceColumn = 0;

      if (fichierDialogue == null || string.IsNullOrWhiteSpace(fichierDialogue.CheminFichier))
        return false;

      if (string.IsNullOrWhiteSpace(sourceSheet) || sourceRow <= 0 || string.IsNullOrWhiteSpace(cours) || string.IsNullOrWhiteSpace(groupe) || string.IsNullOrWhiteSpace(nomIntervenant))
        return false;

      FileInfo fichier = new(fichierDialogue.CheminFichier);

      if (!fichier.Exists)
        return false;

      using ExcelPackage package = new(fichier);

      ExcelWorksheet? feuille = package.Workbook.Worksheets[sourceSheet];

      if (feuille == null || feuille.Dimension == null)
        return false;

      int ligneDestination = TrouverLigneCoursMemeModule(feuille,sourceRow,cours);
      int colonneDestination = TrouverColonneGroupe(feuille,groupe,cours);

      if (ligneDestination <= 0) {
        MessageBox.Show(
          $"Ajout annulé : ligne {cours} introuvable dans la feuille {sourceSheet}.",
          "Ajout intervenant",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
        return false;
      }

      if (colonneDestination <= 0) {
        MessageBox.Show(
          $"Ajout annulé : groupe {groupe} introuvable dans la feuille {sourceSheet}.",
          "Ajout intervenant",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
        return false;
      }

      ExcelRange celluleDestination = feuille.Cells[ligneDestination,colonneDestination];

      if (!string.IsNullOrWhiteSpace(celluleDestination.Text)) {
        MessageBox.Show(
          $"Ajout annulé : la cellule destination {celluleDestination.Address} n'est pas vide.\n\nContenu : {celluleDestination.Text.Trim()}",
          "Ajout intervenant",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
        return false;
      }

      celluleDestination.Value = nomIntervenant;
      package.Save();

      nouvelleSourceRow = ligneDestination;
      nouvelleSourceColumn = colonneDestination;

      Debug.WriteLine($"AJOUT SOURCE : {sourceSheet}!{celluleDestination.Address} = {nomIntervenant}");

      return true;
    }

    internal static int ReecrireLignesSourcesGroupesSemestre_Epplus(ClasseExcel fichierDialogue,DataTable tableGlobal,IEnumerable<ExcelSourceRowRewrite> reecritures) {

      if (fichierDialogue == null || reecritures == null)
        return 0;

      if (string.IsNullOrWhiteSpace(fichierDialogue.CheminFichier))
        return 0;

      FileInfo fichier = new(fichierDialogue.CheminFichier);

      if (!fichier.Exists)
        return 0;

      List<ExcelSourceRowRewrite> reecrituresValides =
          reecritures
              .Where(r =>
                  r != null &&
                  !string.IsNullOrWhiteSpace(r.SourceSheet) &&
                  r.AncienneSourceRow > 0 &&
                  !string.IsNullOrWhiteSpace(r.AncienCours) &&
                  !string.IsNullOrWhiteSpace(r.NouveauCours) &&
                  !string.IsNullOrWhiteSpace(r.NouveauGroupe) &&
                  !string.IsNullOrWhiteSpace(r.NomIntervenant))
              .ToList();

      if (reecrituresValides.Count == 0)
        return 0;

      int lignesReecrites = 0;

      using ExcelPackage package = new(fichier);

      foreach (ExcelSourceRowRewrite reecriture in reecrituresValides) {

        ExcelWorksheet? feuille =
            package.Workbook.Worksheets[reecriture.SourceSheet];

        if (feuille == null || feuille.Dimension == null)
          continue;

        // Colonne physique du nouveau groupe dans S1 ou S2.
        int nouvelleColonne = TrouverColonneGroupe(feuille,reecriture.NouveauGroupe,reecriture.NouveauCours);
        int nouvelleLigne = TrouverLigneCoursMemeModule(feuille,reecriture.AncienneSourceRow,reecriture.NouveauCours);

        if (nouvelleLigne <= 0) {
          Debug.WriteLine(
              $"Ligne destination introuvable : " +
              $"{reecriture.SourceSheet} / {reecriture.NouveauCours}");

          continue;
        }

        if (nouvelleColonne <= 0) {
          Debug.WriteLine(
              $"Groupe destination introuvable : " +
              $"{reecriture.SourceSheet} / {reecriture.NouveauGroupe}");

          continue;
        }

        ExcelRange celluleSource = feuille.Cells[reecriture.AncienneSourceRow,reecriture.AncienneSourceColumn];
        ExcelRange celluleDestination = feuille.Cells[nouvelleLigne,nouvelleColonne];

        // Sécurité : la cellule source doit toujours contenir l'intervenant attendu.
        if (!string.Equals(
            celluleSource.Text.Trim(),
            reecriture.NomIntervenant,
            StringComparison.OrdinalIgnoreCase)) {

          MessageBox.Show(
              $"Déplacement annulé : la cellule source ne contient plus l'intervenant attendu." +
              $"\n\nCellule : {celluleSource.Address}" +
              $"\nAttendu : {reecriture.NomIntervenant}" +
              $"\nTrouvé : {celluleSource.Text.Trim()}",
              "Mise à jour source",
              MessageBoxButton.OK,
              MessageBoxImage.Warning);

          continue;
        }

        string intervenantDestination = celluleDestination.Text.Trim();
        bool echangeGroupes = !string.IsNullOrWhiteSpace(intervenantDestination);

        if (echangeGroupes) {
          DataRow? ligneIntervenantDestination = tableGlobal.Rows.Cast<DataRow>().FirstOrDefault(r =>
              string.Equals(r[ExcelSchemaNames.Columns.SourceSheet]?.ToString()?.Trim(),reecriture.SourceSheet,StringComparison.OrdinalIgnoreCase) &&
              string.Equals(r[ExcelSchemaNames.Columns.Noms]?.ToString()?.Trim(),intervenantDestination,StringComparison.OrdinalIgnoreCase) &&
              string.Equals(r[ExcelSchemaNames.Columns.SourceRow]?.ToString()?.Trim(),nouvelleLigne.ToString(),StringComparison.OrdinalIgnoreCase) &&
              string.Equals(r[ExcelSchemaNames.Columns.SourceColumn]?.ToString()?.Trim(),nouvelleColonne.ToString(),StringComparison.OrdinalIgnoreCase));

          if (ligneIntervenantDestination == null) {
            MessageBox.Show(
                $"Échange annulé : l'intervenant '{intervenantDestination}' a bien été trouvé dans Excel, mais sa ligne correspondante est introuvable dans le DataTable.",
                "Échange de groupes",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

            continue;
          }

          ligneIntervenantDestination[ExcelSchemaNames.Columns.Groupe] = reecriture.AncienGroupe;
          ligneIntervenantDestination[ExcelSchemaNames.Columns.SourceRow] = reecriture.AncienneSourceRow.ToString();
          ligneIntervenantDestination[ExcelSchemaNames.Columns.SourceColumn] = reecriture.AncienneSourceColumn.ToString();

          celluleSource.Value = intervenantDestination;
          celluleDestination.Value = reecriture.NomIntervenant;
        }
        else {
          celluleSource.Value = null;
          celluleDestination.Value = reecriture.NomIntervenant;
        }

        MessageBox.Show(
            $"Intervenant : {reecriture.NomIntervenant}" +
            $"\nAvant : {reecriture.AncienCours} / {reecriture.AncienGroupe}" +
            $"\nAprès : {reecriture.NouveauCours} / {reecriture.NouveauGroupe}" +
            $"\n\nSource : {celluleSource.Address} = {celluleSource.Text}" +
            $"\nDestination : {celluleDestination.Address}" +
            $"\nNouvelle ligne : {nouvelleLigne}" +
            $"\nNouvelle colonne : {nouvelleColonne}");

        reecriture.NouvelleSourceRow = nouvelleLigne;

        DataRow? ligneGlobal =
            tableGlobal.Rows
                .Cast<DataRow>()
                .FirstOrDefault(r =>
                    string.Equals(
                        r[ExcelSchemaNames.Columns.SourceSheet]?.ToString()?.Trim(),
                        reecriture.SourceSheet,
                        StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(
                        r[ExcelSchemaNames.Columns.Noms]?.ToString()?.Trim(),
                        reecriture.NomIntervenant,
                        StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(
                        r[ExcelSchemaNames.Columns.SourceRow]?.ToString()?.Trim(),
                        reecriture.AncienneSourceRow.ToString(),
                        StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(
                        r[ExcelSchemaNames.Columns.Groupe]?.ToString()?.Trim(),
                        reecriture.NouveauGroupe,
                        StringComparison.OrdinalIgnoreCase));

        if (ligneGlobal != null) {
          ligneGlobal[ExcelSchemaNames.Columns.SourceRow] = nouvelleLigne.ToString();
          ligneGlobal[ExcelSchemaNames.Columns.SourceColumn] = nouvelleColonne.ToString();
        }

        lignesReecrites++;
      }

      if (lignesReecrites > 0)
        package.Save();

      return lignesReecrites;
    }

    private static int TrouverLigneCoursMemeModule(ExcelWorksheet feuille,int ligneSource,string nouveauCours) {
      if (feuille.Dimension == null || ligneSource <= 0)
        return 0;

      string codeSource = feuille.Cells[ligneSource,2].Text.Trim();

      if (string.IsNullOrWhiteSpace(codeSource))
        return 0;

      int dernierTiret = codeSource.LastIndexOf('-');

      if (dernierTiret <= 0)
        return 0;

      string prefixeModule = codeSource[..dernierTiret];
      string codeRecherche = prefixeModule + "-" + nouveauCours.ToUpperInvariant();

      // On reste volontairement autour de la ligne source pour ne pas
      // tomber sur le même code OSE d'un autre module.
      int premiereLigne = Math.Max(feuille.Dimension.Start.Row,ligneSource - 5);
      int derniereLigne = Math.Min(feuille.Dimension.End.Row,ligneSource + 5);

      for (int ligne = premiereLigne;ligne <= derniereLigne;ligne++) {
        if (string.Equals(
            feuille.Cells[ligne,2].Text.Trim(),
            codeRecherche,
            StringComparison.OrdinalIgnoreCase))
          return ligne;
      }

      return 0;
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

    private static int TrouverColonneGroupe(ExcelWorksheet feuille,string groupe,string cours) {
      if (string.IsNullOrWhiteSpace(groupe) || feuille.Dimension == null)
        return 0;

      if (!TrouverCelluleTexteFeuilleSemestre(
          feuille,
          "SALLES",
          out int ligneSalles,
          out int colonneSalles))
        return 0;

      int ligneGroupes =
          string.Equals(cours,"TP",StringComparison.OrdinalIgnoreCase)
              ? ligneSalles + 1
              : ligneSalles;

      for (int colonne = colonneSalles + 1;colonne <= feuille.Dimension.End.Column;colonne++) {
        if (string.Equals(
            feuille.Cells[ligneGroupes,colonne].Text.Trim(),
            groupe,
            StringComparison.OrdinalIgnoreCase))
          return colonne;
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
      internal string SourceSheet { get; set; } = string.Empty;
      internal int SourceRow { get; set; }
      internal string Cours { get; set; } = string.Empty;
      public int AncienneSourceRow { get; set; }
      public string AncienCours { get; set; } = string.Empty;
      public string AncienGroupe { get; set; } = string.Empty;

      public int NouvelleSourceRow { get; set; }
      public string NouveauCours { get; set; } = string.Empty;
      public string NouveauGroupe { get; set; } = string.Empty;

      public string NomIntervenant { get; set; } = string.Empty;
      public int AncienneSourceColumn { get; set; }
    }


    internal sealed class ExcelSourceReference {
      internal string Module { get; init; } = string.Empty;
      internal string SourceSheet { get; init; } = string.Empty;
      internal int SourceRow { get; init; }
      internal int SourceColumn { get; init; }
    }
  }
}
