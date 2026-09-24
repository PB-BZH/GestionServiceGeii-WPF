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
║  Nom de fichier : FicheServicePdfHelper.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace GestionServiceGeii.Core.Services;

internal static class FicheServicePdfHelper {
  internal static void ExporterPdf(
      string cheminFichier,
      string titre,
      string intervenant,
      string statut,
      IEnumerable<LigneServicePrevisionnel> lignes,
      RegleStatutService regle) {

    List<LigneServicePrevisionnel> lignesTriees =
        lignes
            .OrderBy(ligne => FicheServiceHelper.ObtenirOrdreSemestre(ligne.Semestre))
            .ThenBy(ligne => ligne.Ose)
            .ThenBy(ligne => ligne.Module)
            .ToList();

    decimal totalCm =
        lignesTriees.Sum(ligne => ligne.Cm);

    decimal totalTd =
        lignesTriees.Sum(ligne => ligne.Td);

    decimal totalTp =
        lignesTriees.Sum(ligne => ligne.Tp);

    decimal totalRetenu =
        lignesTriees.Sum(ligne => ligne.TotalRetenu);

    decimal ecartService =
        totalRetenu - regle.ServiceReference;

    string messageHc = string.Empty;

    if (ecartService < 0)
      messageHc = "Sous service";
    else
      messageHc = "Heures complémentaires";

    Document
        .Create(container => {
          container.Page(page => {
            page.Size(PageSizes.A4);
            page.Margin(2,Unit.Centimetre);
            page.DefaultTextStyle(text => text.FontSize(10));

            page.Header()
                .Column(colonne => {
                  colonne.Item()
                      .Text(titre)
                      .SemiBold()
                      .FontSize(16);

                  colonne.Item()
                      .PaddingTop(12)
                      .Text($"Intervenant : {intervenant}");

                  colonne.Item()
                      .Text($"Statut      : {statut}");

                  colonne.Item()
                      .PaddingTop(6)
                      .Text($"Date édition : {DateTime.Now:dd/MM/yyyy HH:mm}");
                });

            page.Content()
                .PaddingTop(20)
                .Column(colonne => {
                  colonne.Item()
                      .Table(table => {
                        table.ColumnsDefinition(columns => {
                          columns.RelativeColumn(1.2f); // OSE
                          columns.RelativeColumn(3.8f); // Module
                          columns.RelativeColumn(1);
                          columns.RelativeColumn(1);
                          columns.RelativeColumn(1);
                          columns.RelativeColumn(1.4f);
                        });

                        AjouterCelluleEntete(table,"OSE");
                        AjouterCelluleEntete(table,"Module");
                        AjouterCelluleEnteteCentre(table,"CM");
                        AjouterCelluleEnteteCentre(table,"TD");
                        AjouterCelluleEnteteCentre(table,"TP");
                        AjouterCelluleEnteteCentre(table,"Retenu");

                        foreach (var groupeSemestre in lignesTriees.GroupBy(ligne => ligne.Semestre)) {
                          AjouterCelluleSemestre(
                              table,
                              FicheServiceHelper.ConstruireLibelleSemestre(groupeSemestre.Key));

                          foreach (LigneServicePrevisionnel ligne in groupeSemestre) {
                            AjouterCelluleTexte(table,ligne.Ose);
                            AjouterCelluleTexte(table,ligne.Module);
                            AjouterCelluleNombre(table,ligne.Cm);
                            AjouterCelluleNombre(table,ligne.Td);
                            AjouterCelluleNombre(table,ligne.Tp);
                            AjouterCelluleNombre(table,ligne.TotalRetenu);
                          }
                        }

                        AjouterCelluleTotalLibelle(table,"TOTAL");
                        AjouterCelluleTotalNombre(table,totalCm);
                        AjouterCelluleTotalNombre(table,totalTd);
                        AjouterCelluleTotalNombre(table,totalTp);
                        AjouterCelluleTotalNombre(table,totalRetenu);
                      });

                  colonne.Item()
                      .PaddingTop(20)
                      .Table(table => {
                        table.ColumnsDefinition(columns => {
                          columns.RelativeColumn(3);
                          columns.RelativeColumn(2);
                        });

                        AjouterLigneSynthese(
                            table,
                            "Service de référence",
                            regle.ServiceReference > 0m
                                ? $"{regle.ServiceReference:0.##} h"
                                : "-");

                        AjouterLigneSynthese(
                            table,
                            "Service retenu",
                            $"{totalRetenu:0.##} h");

                        if (regle.ServiceReference > 0m) {
                          AjouterLigneSynthese(
                              table,
                              messageHc,
                              $"{ecartService:0.##} h");
                        }
                      });
                });

            page.Footer()
                .AlignCenter()
                .Text(texte => {
                  texte.Span("Gestion Services GEII - ");
                  texte.Span("Page ");
                  texte.CurrentPageNumber();
                  texte.Span(" / ");
                  texte.TotalPages();
                });
          });
        })
        .GeneratePdf(cheminFichier);
  }

  private static void AjouterCelluleTotalLibelle(TableDescriptor table,string texte) {
    table.Cell()
        .ColumnSpan(2)
        .Background(Colors.Grey.Lighten4)
        .BorderTop(1)
        .BorderColor(Colors.Grey.Medium)
        .Padding(5)
        .Text(texte)
        .SemiBold();
  }

  private static void AjouterCelluleEntete(TableDescriptor table,string texte) {
    table.Cell()
        .Background(Colors.Grey.Lighten3)
        .BorderBottom(1)
        .BorderColor(Colors.Grey.Medium)
        .Padding(5)
        .Text(texte)
        .SemiBold();
  }

  private static void AjouterCelluleTexte(TableDescriptor table,string texte) {
    table.Cell()
        .BorderBottom(0.5f)
        .BorderColor(Colors.Grey.Lighten2)
        .Padding(4)
        .Text(texte);
  }

  private static void AjouterCelluleNombre(TableDescriptor table,decimal valeur) {
    table.Cell()
        .BorderBottom(0.5f)
        .BorderColor(Colors.Grey.Lighten2)
        .Padding(4)
        .AlignCenter()
        .Text($"{valeur:0.##}");
  }

  private static void AjouterCelluleTotal(TableDescriptor table,string texte) {
    table.Cell()
        .Background(Colors.Grey.Lighten4)
        .BorderTop(1)
        .BorderColor(Colors.Grey.Medium)
        .Padding(5)
        .Text(texte)
        .SemiBold();
  }

  private static void AjouterCelluleTotalNombre(TableDescriptor table,decimal valeur) {
    table.Cell()
        .Background(Colors.Grey.Lighten4)
        .BorderTop(1)
        .BorderColor(Colors.Grey.Medium)
        .Padding(5)
        .AlignCenter()
        .Text($"{valeur:0.##}")
        .SemiBold();
  }

  private static void AjouterCelluleEnteteCentre(TableDescriptor table,string texte) {
    table.Cell()
        .Background(Colors.Grey.Lighten3)
        .BorderBottom(1)
        .BorderColor(Colors.Grey.Medium)
        .Padding(5)
        .AlignCenter()
        .Text(texte)
        .SemiBold();
  }

  private static void AjouterCelluleSemestre(TableDescriptor table,string texte) {
    table.Cell()
        .ColumnSpan(6)
        .Background(Colors.Grey.Lighten4)
        .BorderTop(1)
        .BorderColor(Colors.Grey.Lighten2)
        .PaddingVertical(6)
        .PaddingHorizontal(5)
        .Text(texte)
        .SemiBold();
  }

  private static void AjouterLigneSynthese(
      TableDescriptor table,
      string libelle,
      string valeur) {

    table.Cell()
        .PaddingVertical(3)
        .Text(libelle)
        .SemiBold();

    table.Cell()
        .PaddingVertical(3)
        .AlignRight()
        .Text(valeur);
  }
}