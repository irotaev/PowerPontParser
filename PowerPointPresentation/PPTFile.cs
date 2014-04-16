using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.IO;
using System.Text.RegularExpressions;
using SevenZip;
using BinaryAnalysis.UnidecodeSharp;

namespace PowerPointPresentation
{
  /// <summary>
  /// Работа с файлами PowerPoint
  /// </summary>
  public class PPTFiles : IDisposable
  {
    public const string _PresentationDir = "Presentations";
    public const string _PresentationImageDir = "Presentations/Images";
    public const string _ExtractRelativeDir = "Temp";
    /// <summary>
    /// Поддерживаемые форматы архиватора
    /// </summary>
    public static string[] SupportedArchiveFormats { get { return _SupportedArchiveFormats.ToArray(); } }
    /// <summary>
    /// Поддерживаемые форматы файла презентации power point
    /// </summary>
    public static string[] SupportedPowerPointFileFormats { get { return _SupportedPowerPointFileFormats.ToArray(); } }

    private static readonly string[] _SupportedArchiveFormats = new string[] { ".rar", ".zip", ".7z", ".gzip", ".gz", ".tgz", ".bz2", ".bzip2", ".tbz2", ".tbz", ".tar", ".rpm", ".iso", ".deb", ".cab" };
    private static readonly string[] _SupportedPowerPointFileFormats = new string[] { ".ppt", ".pptx", ".pps", ".ppsx" };
    private const string _External7ZipLib = "Lib/7z.dll";
    private static readonly string _ExtractDir;

    private bool _IsErrorExists;

    /// <summary>
    /// Каллбэк при окончании парсинга слайда презентации
    /// </summary>
    public EventHandler<SlideCompleteParsingInfo> ParseSlideCompleteCallback { get; set; }

    static PPTFiles()
    {
      _ExtractDir = Path.Combine(Directory.GetCurrentDirectory(), _PresentationDir, _ExtractRelativeDir);

      if (!Directory.Exists(_ExtractDir))
        Directory.CreateDirectory(_ExtractDir);

      SevenZipExtractor.SetLibraryPath(Path.Combine(Directory.GetCurrentDirectory(), _External7ZipLib));
    }

    public PPTFiles()
    {
      ClearTemDir();
    }

    /// <summary>
    /// Изъять информацию из презентации
    /// </summary>
    /// <param name="ppFilePath">Полный путь к презентации</param>
    /// <returns>Информация о презентации</returns>
    public PresentationInfo ExtractInfo(string ppFilePath)
    {
      PresentationInfo presInfo = null;

      if (SupportedArchiveFormats.Contains(Path.GetExtension(ppFilePath)))
      {
        ExtractArchive(ppFilePath);

        List<string> allDirectories = Directory.GetDirectories(_ExtractDir).ToList();
        allDirectories.Add(_ExtractDir);

        bool isFileFound = false;

        foreach (string dir in allDirectories)
        {
          foreach (string filePath in Directory.GetFiles(dir))
          {
            if (SupportedPowerPointFileFormats.Contains(Path.GetExtension(filePath)))
            {
              PPTFile pptFile = new PPTFile(filePath, ppFilePath);

              if (ParseSlideCompleteCallback != null)
                pptFile.ParseSlideComplite += ParseSlideCompleteCallback;

              presInfo = pptFile.ParsePreesentation();
              isFileFound = true;

              FileInfo archiveInfo = new FileInfo(ppFilePath);
              presInfo.FileSize = archiveInfo.Length;
              break;
            }
          }
        }

        if (!isFileFound)
        {
          _IsErrorExists = true;
          throw new Exception(String.Format("Файл презентации в архиве не найден\nАрхив был распакован в папку: {0}\nПри следующем запуске программы архив будет удален из временного хранилища",
            _ExtractDir));
        }
      }
      else
      {
        PPTFile pptFile = new PPTFile(ppFilePath);

        if (ParseSlideCompleteCallback != null)
          pptFile.ParseSlideComplite += ParseSlideCompleteCallback;

        presInfo = pptFile.ParsePreesentation();

        FileInfo fileInfo = new FileInfo(ppFilePath);
        presInfo.FileSize = fileInfo.Length;
      }

      return presInfo;
    }

    private void ExtractArchive(string archivePath)
    {
      using (SevenZipExtractor extractor = new SevenZipExtractor(archivePath))
      {
        extractor.ExtractArchive(Path.Combine(_ExtractDir));
      }
    }

    /// <summary>
    /// Очистить временную директорию хранения файлов архива
    /// </summary>
    private void ClearTemDir()
    {
      var directory = new System.IO.DirectoryInfo(_ExtractDir);

      foreach (var file in directory.GetFiles())
      {
        file.Delete();
      }

      foreach (var dir in directory.GetDirectories())
      {
        dir.Delete(true);
      }
    }

    public void Dispose()
    {
      if (!_IsErrorExists)
      {
        ClearTemDir();
      }
    }
  }

  /// <summary>
  /// Файл power point
  /// </summary>
  class PPTFile
  {
    private readonly _Presentation _Presentation;
    private readonly PresentationInfo _PresentationInfo;

    public event EventHandler<SlideCompleteParsingInfo> ParseSlideComplite;

    /// <summary>
    /// Создать файл презентации
    /// </summary>
    /// <param name="ppFilePath">Полный путь к файлу с презентацией</param>
    /// <param name="archivePath">Полный путь к архиву, содержащему презентацию ()</param>
    public PPTFile(string ppFilePath, string archivePath = null)
    {
      Microsoft.Office.Interop.PowerPoint._Application powerPointApp = new Microsoft.Office.Interop.PowerPoint.Application();
      Microsoft.Office.Interop.PowerPoint.Presentations ppPresentations = powerPointApp.Presentations;
      _Presentation = ppPresentations.Open(ppFilePath, MsoTriState.msoCTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

      if (archivePath == null)
        _PresentationInfo = new PresentationInfo(ppFilePath) { SlidersInfo = new List<SlideInfo>() };
      else
        _PresentationInfo = new PresentationInfo(archivePath) { SlidersInfo = new List<SlideInfo>() };
    }

    /// <summary>
    /// Получить текст презентации
    /// </summary>
    /// <returns>Текст</returns>
    public PresentationInfo ParsePreesentation()
    {
      for (int i = 0; i < _Presentation.Slides.Count; i++)
      {
        SlideInfo slideInfo = new SlideInfo(_PresentationInfo) { SlideNumber = i + 1 };
        _PresentationInfo.SlidersInfo.Add(slideInfo);

        foreach (var shape in _Presentation.Slides[i + 1].Shapes.Cast<Microsoft.Office.Interop.PowerPoint.Shape>())
        {
          if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
            slideInfo.Text += shape.TextFrame.TextRange.Text + "<br/>";

          if (shape.HasTable == MsoTriState.msoTrue)
          {
            for (int rowNum = 1; rowNum <= shape.Table.Rows.Count; rowNum++)
            {
              bool isNeedBrBetweenRows = false;

              for (int cellNum = 1; cellNum <= shape.Table.Columns.Count; cellNum++)
              {
                var cell = shape.Table.Cell(rowNum, cellNum);

                if (cell.Shape.HasTextFrame == MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                  slideInfo.Text += cell.Shape.TextFrame.TextRange.Text + " ";
                  isNeedBrBetweenRows = true;
                }
              }

              if (isNeedBrBetweenRows)
                slideInfo.Text += "<br/>";
            }
          }
        }

        string slidePath = Path.Combine(Directory.GetCurrentDirectory(), PPTFiles._PresentationImageDir);

        if (!Directory.Exists(slidePath))
          Directory.CreateDirectory(slidePath);

        if (i == 0)
        {
          string slideImageNameSmall = Guid.NewGuid().ToString() + "_195x146" + ".jpg";
          _Presentation.Slides[i + 1].Export(Path.Combine(slidePath, slideImageNameSmall), "JPG", 195, 146);
          slideInfo.ImageNameClientSmall = slideImageNameSmall;
        }

        string slideImageNameAverage = Guid.NewGuid().ToString() + "_225x167" + ".jpg";
        _Presentation.Slides[i + 1].Export(Path.Combine(slidePath, slideImageNameAverage), "JPG", 225, 167);

        string slideImageNameBig = Guid.NewGuid().ToString() + "_500x374" + ".jpg";
        _Presentation.Slides[i + 1].Export(Path.Combine(slidePath, slideImageNameBig), "JPG", 500, 374);

        slideInfo.ImageNameClientAverage = slideImageNameAverage;
        slideInfo.ImageNameClientBig = slideImageNameBig;

        ParseSlideComplite(this, new SlideCompleteParsingInfo { SlideCurrentNumber = slideInfo.SlideNumber, SlideTotalNumber = _Presentation.Slides.Count });
      }

      return _PresentationInfo;
    }

    /// <summary>
    /// Форматирование текста
    /// </summary>
    /// <param name="text">Текст</param>
    private string FormatText(string text)
    {
      text = Regex.Replace(text, @"\.([а-яА-Яa-zA-Z])", new MatchEvaluator((Match match) =>
        {
          if (match.Groups[1] != null && String.IsNullOrEmpty(match.Groups[1].Value))
          {
            return String.Format(@"\.{0}", match.Groups[1].Value.ToUpper());
          }
          else
          {
            return match.Groups[0].Value;
          }
        }));

      return text;
    }
  }

  /// <summary>
  /// Информация о презентации
  /// </summary>
  public class PresentationInfo
  {
    private readonly string _UniqueId;
    private string _Name;

    public PresentationInfo(string clientFilePath)
    {
      if (String.IsNullOrEmpty(clientFilePath))
        throw new ArgumentNullException("clientFilePath");

      _ClientFilePath = clientFilePath;
      _UniqueId = Guid.NewGuid().ToString();
    }

    private string _UrlNews;
    /// <summary>
    /// Url_news поле в базе данных
    /// </summary>
    public string UrlNews
    {
      get
      {
        if (String.IsNullOrEmpty(_UrlNews))
        {
          return NameAsTranslit;
        }
        else
        {
          return _UrlNews;
        }
      }
      set
      {
        _UrlNews = value;
      }
    }

    /// <summary>
    /// Уникальный индентификатор презентации
    /// </summary>
    public string UniqueId { get { return _UniqueId; } }
    /// <summary>
    /// Информация о слайдах
    /// </summary>
    public List<SlideInfo> SlidersInfo { get; set; }
    /// <summary>
    /// Имя презентации
    /// </summary>
    public string Name
    {
      get
      {
        return _Name;
      }
      set
      {
        _Name = value;

        string nameFormat1 = Regex.Replace(_Name.Unidecode(), @"\s+", "_");
        string nameFormat2 = Regex.Replace(nameFormat1, @"[^[a-zA-Z0-9_]", "");
        NameAsTranslit = nameFormat2.ToLower();
      }
    }
    /// <summary>
    /// Заголовок презентации
    /// </summary>
    public string Title { get; set; }
    /// <summary>
    /// Категория презентации
    /// </summary>
    public Categortie Categorie { get; set; }
    /// <summary>
    /// Уникальный Id презентации в базе данных
    /// </summary>
    public long DbId { get; set; }
    /// <summary>
    /// Имя презентации транслитом
    /// </summary>
    public string NameAsTranslit { get; set; }
    /// <summary>
    /// Размер файла с презентацией
    /// </summary>
    public float? FileSize { get; set; }
    /// <summary>
    /// Последний индекс маленькой картинки в папке с маленькими каринками на ftp
    /// </summary>
    public int LastImageSmallIndex { get; set; }
    /// <summary>
    /// Имя презентации на сервере
    /// </summary>
    public string ServerFileName
    {
      get
      {
        return String.Format(@"volna_org_{0}{1}", UrlNews, Path.GetExtension(ClientFilePath));
      }
    }
    /// <summary>
    /// Url для скачивания презентации
    /// </summary>
    public string UrlDownload
    {
      get
      {
        return String.Format("{0}/{1}", FTP.FilesServerDir, ServerFileName);
      }
    }

    private readonly string _ClientFilePath;
    /// <summary>
    /// Путь к файлу на стороне клиента, откуда взята презентация
    /// </summary>
    public string ClientFilePath { get { return _ClientFilePath; } }
  }

  /// <summary>
  /// Информация о слайде
  /// </summary>
  public class SlideInfo
  {
    public SlideInfo(PresentationInfo presInfo)
    {
      if (presInfo == null)
        throw new ArgumentNullException("presInfo");

      _PresentationInfo = presInfo;
    }

    private readonly PresentationInfo _PresentationInfo;

    /// <summary>
    /// Номер слайда
    /// </summary>
    public int SlideNumber { get; set; }
    /// <summary>
    /// Текст слайда
    /// </summary>
    public string Text { get; set; }
    /// <summary>
    /// Имя картинки слайда на клиенте (маленький)
    /// </summary>
    public string ImageNameClientSmall { get; set; }
    /// <summary>
    /// Имя картинки слайда на сервере (маленький)
    /// </summary>
    public string ImageNameServerSmall { get; set; }
    /// <summary>
    /// Имя картинки слайда на клиенте (средний)
    /// </summary>
    public string ImageNameClientAverage { get; set; }
    /// <summary>
    /// Имя картинки слайда на сервере (средний)
    /// </summary>
    public string ImageNameServerAverage { get; set; }
    /// <summary>
    /// Имя картинки слайда на клиенте (большой)
    /// </summary>
    public string ImageNameClientBig { get; set; }
    /// <summary>
    /// Имя картинки слайда на сервере (большой)
    /// </summary>
    public string ImageNameServerBig { get; set; }
  }

  /// <summary>
  /// Категории
  /// </summary>
  public enum Categortie
  {
    NA = -1,
    anglijskij_jazyk = 0,
    astronomija = 1,
    algebra = 2,
    biologija = 4,
    geografija = 5,
    geometrija = 6,
    informatika = 7,
    istorija = 8,
    literatura = 9,
    matematika = 10,
    medicina = 11,
    mhk_i_izo = 12,
    muzyka = 13,
    obzh = 14,
    obshhestvoznanie = 15,
    okruzhajushhij_mir = 16,
    pedagogika = 17,
    russkij_jazyk = 18,
    tehnologija = 19,
    ukrainskij_jazyk = 20,
    fizika = 21,
    fizkultura = 22,
    filosofija = 23,
    himija = 24,
    jekologija = 25,
    jekonomika = 26,
    detskie_prezentacii = 27
  }

  /// <summary>
  /// Конвертирует перечеслитель
  /// </summary>
  public class EnumConverter
  {
    public static string Categorie(Categortie categorie)
    {
      switch (categorie)
      {
        case Categortie.anglijskij_jazyk:
          return "Английский язык";
        case Categortie.astronomija:
          return "Астрономия";
        case Categortie.algebra:
          return "Алгебра";
        case Categortie.biologija:
          return "Биология";
        case Categortie.geografija:
          return "География";
        case Categortie.geometrija:
          return "Геометрия";
        case Categortie.informatika:
          return "Информатика";
        case Categortie.istorija:
          return "История";
        case Categortie.literatura:
          return "Литература";
        case Categortie.matematika:
          return "Математика";
        case Categortie.medicina:
          return "Медицина";
        case Categortie.mhk_i_izo:
          return "МХК и ИЗО";
        case Categortie.muzyka:
          return "Музыка";
        case Categortie.obzh:
          return "ОБЖ";
        case Categortie.obshhestvoznanie:
          return "Обществознание";
        case Categortie.okruzhajushhij_mir:
          return "Окружающий мир";
        case Categortie.pedagogika:
          return "Педагогика";
        case Categortie.russkij_jazyk:
          return "Русский язык";
        case Categortie.tehnologija:
          return "Технология";
        case Categortie.ukrainskij_jazyk:
          return "Украинский язык";
        case Categortie.fizika:
          return "Физика";
        case Categortie.fizkultura:
          return "Физкультура";
        case Categortie.filosofija:
          return "Философия";
        case Categortie.himija:
          return "Химия";
        case Categortie.jekologija:
          return "Экология";
        case Categortie.jekonomika:
          return "Экономика";
        case Categortie.detskie_prezentacii:
          return "Детские презентации";
        case Categortie.NA:
          return "Не выбрана";
        default:
          return "Не задан конвертер";
      }
    }
  }

  /// <summary>
  /// Информации при заверщении парсинга слайда
  /// </summary>
  public class SlideCompleteParsingInfo : EventArgs
  {
    /// <summary>
    /// Текущий номер слайда
    /// </summary>
    public int SlideCurrentNumber { get; set; }
    /// <summary>
    /// Полное число слайдов
    /// </summary>
    public int SlideTotalNumber { get; set; }
  }
}
