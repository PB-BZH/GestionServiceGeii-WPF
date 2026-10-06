namespace GestionServiceGeii.Core.Enums;

public enum SaveWindowStateMode {
  Normal,
  Minimized,
  Maximized
}

public sealed class SaveWindowState {
  public double Left { get; set; }
  public double Top { get; set; }
  public double Width { get; set; }
  public double Height { get; set; }

  public SaveWindowStateMode State { get; set; } = SaveWindowStateMode.Normal;
}