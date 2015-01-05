using MahApps.Metro.Controls;
using MySql.Data.MySqlClient;
using PowerPointPresentation.Views;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;

namespace PowerPointPresentation
{
  public partial class MainWindow : MetroWindow
  {
    
    private Dictionary<Categortie, string> _Categories = new Dictionary<Categortie, string>();
    public Dictionary<Categortie, string> Categories { get { return _Categories; } }

    private string _PresentationFullPath;

    protected override void OnInitialized(EventArgs e)
    {
      base.OnInitialized(e);

      _Categories.Add(Categortie.NA, EnumConverter.Categorie(Categortie.NA));
      List<Categortie> allCAtegories = Enum.GetValues(typeof(Categortie)).Cast<Categortie>().ToList();
      allCAtegories.Remove(Categortie.NA);

      foreach (Categortie category in allCAtegories)
      {
        _Categories.Add(category, EnumConverter.Categorie(category));
      }

      PresentationFileName.Text = "Файл не выбран";
    }

    public MainWindow()
    {
      InitializeComponent();

      #region Проверка лицензии
      //try
      //{
      //  using (var licenseVerifier = new PowerPointPresentation.Lib.LicenseVerifier())
      //  {
      //    if (!licenseVerifier.CheckLicense())
      //    {
      //      MessageBox.Show(String.Format("Ваша лицензия не активна\nВозможно Вам необходимо продлить лицензию"));
      //      Application.Current.Shutdown();
      //    }
      //  }
      //}
      //catch (Exception ex)
      //{
      //  MessageBox.Show(ex.Message);
      //  Application.Current.Shutdown();
      //}
      #endregion
    }
      
    private void Button_Click_1(object sender, RoutedEventArgs e)
    {
      Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();

      string allFileFormatsString = String.Join(";", PPTFiles.SupportedPowerPointFileFormats.Concat(PPTFiles.SupportedArchiveFormats).Select(el => { return "*" + el; }));

      dialog.DefaultExt = ".ppt";
      dialog.Filter = String.Format("Презентация или архив power point|{0}", allFileFormatsString);

      bool? result = dialog.ShowDialog();

      if (result == true)
      {
        PresentationFileName.Text = System.IO.Path.GetFileName(dialog.FileName);
        _PresentationFullPath = dialog.FileName;
      }
    }

    private void Button_Click_2(object sender, RoutedEventArgs e)
    {
      #region Проверка на валидность
      if (PresentationFileName.Text == "Файл не выбран"
        || String.IsNullOrEmpty(PresentationName.Text)
        || String.IsNullOrEmpty(PresentationTitle.Text))
        //|| CategorieComboBox.SelectedItem == null
        //|| ((KeyValuePair<Categortie, string>)CategorieComboBox.SelectedItem).Key == Categortie.NA)
        
      {
        MessageBox.Show("Вы неправильно заполнили поля");
        return;
      }
      #endregion

      MainGrid.Opacity = 0.2;
      AppWindow.IsEnabled = false;
      
      BackgroundWorker worker = new BackgroundWorker();
      worker.DoWork += worker_DoWork;
      worker.RunWorkerCompleted += worker_RunWorkerCompleted;
      worker.WorkerReportsProgress = true;
      worker.ProgressChanged += worker_ProgressChanged;

      WorkerArgument workerArgument = new WorkerArgument
      {
        PresentationName = PresentationName.Text,
        PresentationTitle = PresentationTitle.Text,
        SelectedItem = CategorieComboBox.SelectedItem,
        UrlNews = UrlNews.Text
      };

      worker.RunWorkerAsync(workerArgument);
    }

    void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
      if (ProgressInfo.Visibility != System.Windows.Visibility.Visible)
        ProgressInfo.Visibility = System.Windows.Visibility.Visible;
      
      if (ProgressBar.Visibility != System.Windows.Visibility.Visible)
        ProgressBar.Visibility = System.Windows.Visibility.Visible;

      ProgressInfo.Text = String.Format("{0}", e.UserState);

      ProgressBar.Value = e.ProgressPercentage;
    }

    private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      MainGrid.Opacity = 1;
      ProgressInfo.Visibility = System.Windows.Visibility.Collapsed;
      ProgressBar.Visibility = System.Windows.Visibility.Collapsed;
      AppWindow.IsEnabled = true;

      if (e.Error != null)
      {
        MessageBox.Show(String.Format("Во время обработки презентации {0} произошла ошибка \n\n Ошибка:\n {1}", PresentationFileName.Text, e.Error.Message));
      }
      else
      {
        // Всплывающее собщение, что парсинг прошел успешно
        MessagePopUp.Text = "Парсинг прошел успешно";
        //MessagePopUp.Visibility = System.Windows.Visibility.Visible;
        Storyboard messagePopUp = (Storyboard)TryFindResource("StoryboardMessagePopUp");
        messagePopUp.Begin();
      }
    }

    private void worker_DoWork(object sender, DoWorkEventArgs e)
    {
      WorkerArgument argument = (WorkerArgument)e.Argument;

      #region Парсинг презентации
      PresentationInfo presInfo = null;
      IAbstractDBTable abstractpresTable = null;
      using (PPTFiles pptFiles = new PPTFiles())
      {
        #region Получаюданные настройки соединения с БД
        string dbRemoteHost = null,
               dbName = null,
               dbUser = null,
               dbPassword = null;

        try
        {
          XDocument xmlDBDoc = XDocument.Load("Lib\\FCashProfile.tss");

          var XdbRemoteHost = xmlDBDoc.Root.Element(XName.Get("ExportDBInfo")).Element(XName.Get("DBRemoteHost"));
          dbRemoteHost = XdbRemoteHost.Value;

          var XdbName = xmlDBDoc.Root.Element(XName.Get("ExportDBInfo")).Element(XName.Get("DBName"));
          dbName = XdbName.Value;

          var XdbUser = xmlDBDoc.Root.Element(XName.Get("ExportDBInfo")).Element(XName.Get("DBUser"));
          dbUser = XdbUser.Value;

          var XdbPassword = xmlDBDoc.Root.Element(XName.Get("ExportDBInfo")).Element(XName.Get("DBPassword"));
          dbPassword = XdbPassword.Value;
        }
        catch (Exception ex)
        {
          throw new Exception(String.Format("Не получилось получить конфигурационные данные из файла конфигурации: {0}", ex.Message));
        }

        if (!String.IsNullOrEmpty(argument.UrlNews))
          presInfo.UrlNews = argument.UrlNews;


        if (String.IsNullOrEmpty(dbRemoteHost) || String.IsNullOrEmpty(dbName) || String.IsNullOrEmpty(dbUser))
          throw new Exception("У вас не заполнена конфигурация соединения с базой данных для экспорта\nПожалуйста заполните ее через настройки");

        MySQLPresentationTable presTable = new MySQLPresentationTable(dbRemoteHost, dbName, dbUser, dbPassword);        
        abstractpresTable = presTable;
        #endregion

        pptFiles.ParseSlideCompleteCallback += (object pptFile, SlideCompleteParsingInfo slideParsingInfo) =>
        {
          ((BackgroundWorker)sender).ReportProgress((int)((decimal)slideParsingInfo.SlideCurrentNumber / (decimal)slideParsingInfo.SlideTotalNumber * 100), "Обработка слайдов");
        };

        presInfo = pptFiles.ExtractInfo(_PresentationFullPath, presTable);
        presInfo.Name = argument.PresentationName;
        presInfo.Title = argument.PresentationTitle;
        presInfo.Categorie = ((KeyValuePair<Categortie, string>)argument.SelectedItem).Key;
      }
      #endregion

      #region Заливка информации по презентации в БД
      {
        abstractpresTable.PutDataOnServer(presInfo);
      }
      #endregion

      #region Отправка на FTP
      try
      {
        XDocument xmlFtpDoc = XDocument.Load("Lib\\FCashProfile.tss");

        var ftpHost = xmlFtpDoc.Root.Element(XName.Get("ExportFtpInfo")).Element(XName.Get("Host"));
        var ftpUserName = xmlFtpDoc.Root.Element(XName.Get("ExportFtpInfo")).Element(XName.Get("UserName"));
        var ftpUserPassword = xmlFtpDoc.Root.Element(XName.Get("ExportFtpInfo")).Element(XName.Get("UserPassword"));
        var ftpImagesDir = xmlFtpDoc.Root.Element(XName.Get("ExportFtpInfo")).Element(XName.Get("ImagesDir"));

        FTP ftp = new FTP(ftpHost.Value, ftpUserName.Value, ftpUserPassword.Value, ftpImagesDir.Value);
        ftp.UploadImageCompleteCallback += (object ftpSender, UploadImageCompliteInfo completeInfo) =>
          {
            ((BackgroundWorker)sender).ReportProgress((int)((decimal)completeInfo.CurrentImageNumber / (decimal)completeInfo.TotalImagesCount * 100), "Загрузка изображений на FTP");
          };

        List<string> imageNames = new List<string>();

        foreach (var slideInfo in presInfo.SlidersInfo)
        {
          if (!String.IsNullOrEmpty(slideInfo.ImageNameClientSmall))
            imageNames.Add(slideInfo.ImageNameClientSmall);

          if (!String.IsNullOrEmpty(slideInfo.ImageNameClientAverage))
            imageNames.Add(slideInfo.ImageNameClientAverage);

          if (!String.IsNullOrEmpty(slideInfo.ImageNameClientBig))
            imageNames.Add(slideInfo.ImageNameClientBig);
        }

        ftp.UploadImages(presInfo);
      }
      catch (Exception ex)
      {
        throw new Exception(String.Format("Во время отправки изображений на FTP возникла ошибка: {0}", ex.Message));
      }
      #endregion
    }

    private class WorkerArgument
    {
      public string PresentationName { get; set; }
      public string PresentationTitle { get; set; }
      public object SelectedItem { get; set; }
      public string UrlNews { get; set; }
    }

    private void Button_Click(object sender, RoutedEventArgs e)
    {
      SettingsWindow settingsWindow = new SettingsWindow();
      settingsWindow.ShowDialog();
    }
  }
}
