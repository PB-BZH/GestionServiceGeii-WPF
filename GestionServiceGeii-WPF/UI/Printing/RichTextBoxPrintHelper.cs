using System.IO;
using System.IO.Packaging;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;

namespace GestionServiceGeii.UI.Printing;

internal sealed class RichTextBoxPrintHelper {
  private readonly RichTextBox _richTextBox;

  internal RichTextBoxPrintHelper(RichTextBox richTextBox) {
    _richTextBox = richTextBox ?? throw new ArgumentNullException(nameof(richTextBox));
  }

  internal void ShowPreview(Window owner,string documentName) {
    FlowDocument document = CloneDocument(_richTextBox.Document);

    PrintDialog printDialog = new();
    ConfigureDocument(document,printDialog);

    AfficherApercu(owner,document,documentName);
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

    AfficherApercu(owner,document,documentName);
  }

  private static void AfficherApercu(Window owner,FlowDocument document,string documentName) {
    MemoryStream xpsStream = new();
    Package package = Package.Open(xpsStream,FileMode.Create,FileAccess.ReadWrite);

    string packageUriString = $"memorystream://{Guid.NewGuid():N}.xps";
    Uri packageUri = new(packageUriString);
    PackageStore.AddPackage(packageUri,package);

    XpsDocument xpsDocument = new(package,CompressionOption.Fast,packageUriString);

    IDocumentPaginatorSource paginatorSource = document;
    XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(xpsDocument);
    writer.Write(paginatorSource.DocumentPaginator);

    DocumentViewer viewer = new() {
      Document = xpsDocument.GetFixedDocumentSequence()
    };

    Window previewWindow = new() {
      Owner = owner,
      Title = $"Aperçu avant impression - {documentName}",
      Content = viewer,
      Width = 1100,
      Height = 800,
      WindowStartupLocation = WindowStartupLocation.CenterOwner
    };

    previewWindow.Closed += (_,_) => {
      viewer.Document = null;
      xpsDocument.Close();
      PackageStore.RemovePackage(packageUri);
      package.Close();
      xpsStream.Dispose();
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