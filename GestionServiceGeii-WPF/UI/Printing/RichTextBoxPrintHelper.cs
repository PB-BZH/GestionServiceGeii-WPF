using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace GestionServiceGeii.UI.Printing;

internal sealed class RichTextBoxPrintHelper {
  private readonly RichTextBox _richTextBox;

  internal RichTextBoxPrintHelper(RichTextBox richTextBox) {
    _richTextBox = richTextBox ?? throw new ArgumentNullException(nameof(richTextBox));
  }

  internal void ShowPreview(Window owner,string documentName) {
    FlowDocument document = CloneDocument(_richTextBox.Document);

    DocumentViewer viewer = new() {
      Document = document
    };

    Window previewWindow = new() {
      Owner = owner,
      Title = $"Aperçu avant impression - {documentName}",
      Content = viewer,
      Width = 1100,
      Height = 800,
      WindowStartupLocation = WindowStartupLocation.CenterOwner
    };

    previewWindow.ShowDialog();
  }

  internal void Print(Window owner,string documentName) {
    PrintDialog printDialog = new();

    if (printDialog.ShowDialog() != true)
      return;

    FlowDocument document = CloneDocument(_richTextBox.Document);
    ConfigureDocument(document,printDialog);

    IDocumentPaginatorSource paginatorSource = document;
    printDialog.PrintDocument(paginatorSource.DocumentPaginator,documentName);
  }

  internal void ShowPreviewWithPrinterDialog(Window owner,string documentName) {
    PrintDialog printDialog = new();

    if (printDialog.ShowDialog() != true)
      return;

    FlowDocument document = CloneDocument(_richTextBox.Document);
    ConfigureDocument(document,printDialog);

    DocumentViewer viewer = new() {
      Document = document
    };

    Window previewWindow = new() {
      Owner = owner,
      Title = $"Aperçu avant impression - {documentName}",
      Content = viewer,
      Width = 1100,
      Height = 800,
      WindowStartupLocation = WindowStartupLocation.CenterOwner
    };

    previewWindow.ShowDialog();
  }

  private static void ConfigureDocument(FlowDocument document,PrintDialog printDialog) {
    document.PageWidth = printDialog.PrintableAreaWidth;
    document.PageHeight = printDialog.PrintableAreaHeight;
    document.PagePadding = new Thickness(50);
    document.ColumnGap = 0;
    document.ColumnWidth = double.PositiveInfinity;
  }

  private static FlowDocument CloneDocument(FlowDocument source) {
    FlowDocument document = new();

    TextRange sourceRange = new(source.ContentStart,source.ContentEnd);
    TextRange destinationRange = new(document.ContentStart,document.ContentEnd);

    using MemoryStream stream = new();
    sourceRange.Save(stream,DataFormats.XamlPackage);
    stream.Position = 0;
    destinationRange.Load(stream,DataFormats.XamlPackage);

    return document;
  }
}