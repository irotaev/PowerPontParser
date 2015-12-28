using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PowerPointPresentation.Control;
using PowerPointPresentation.PresentationControl;

namespace PowerPointPresentation.Views
{
  /// <summary>
  /// Interaction logic for PresentationControl.xaml
  /// </summary>
  public partial class PresentationControl : UserControl
  {
    public PresentationControl(MainWindow window, Dictionary<Categortie, string> categories)
    {
      _window = window;
      ControlState = PresentationControlState.WaitingExecution;

      Categories = categories;

      InitializeComponent();

      PresentationFileName.Text = "Файл не выбран";
    }

    private readonly MainWindow _window;
    private string _presentationFullPath;
    public Dictionary<Categortie, string> Categories { get; set; }

    internal PresentationControlState ControlState { get; set; }

    private void ButtonRemove_OnClick(object sender, RoutedEventArgs e)
    {
      ((Panel)this.Parent).Children.Remove(this);

      _window.RemoveControl(this);
    }

    public PresentationData GetData()
    {
      return new PresentationData
      {
        Category = CategorieComboBox.SelectedItem,
        PresentationFullPath = _presentationFullPath,
        PresentationName = PresentationName.Text
      };
    }

    private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
    {
      Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();

      string allFileFormatsString = String.Join(";", PPTFiles.SupportedPowerPointFileFormats.Concat(PPTFiles.SupportedArchiveFormats).Select(el => { return "*" + el; }));

      dialog.DefaultExt = ".ppt";
      dialog.Filter = String.Format("Презентация или архив power point|{0}", allFileFormatsString);

      bool? result = dialog.ShowDialog();

      if (result == true)
      {
        PresentationFileName.Text = System.IO.Path.GetFileName(dialog.FileName);
        _presentationFullPath = dialog.FileName;
      }
    }
  }
}
