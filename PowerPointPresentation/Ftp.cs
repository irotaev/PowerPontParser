﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace PowerPointPresentation
{
  /// <summary>
  /// Работа с FTP
  /// </summary>
  public class FTP
  {
    public const string AverageAndBigImageServerDir = "images";
    public const string SmallImageServerDir = "img_main";
    public const string FilesServerDir = "files";

    private readonly string _FTPHost;
    private readonly string _UserName;
    private readonly string _UserPassword;
    private readonly string _UploadImagesBaseDir;

    private List<string> _CreatedFtpDirs = new List<string>();

    /// <summary>
    /// Коллбэк на завершение загрузки картинки по FTP
    /// </summary>
    public event EventHandler<UploadImageCompliteInfo> UploadImageCompleteCallback;

    /// <summary>
    /// 
    /// </summary>
    /// <param name="ftpHost">Имя FTP хоста</param>
    /// <param name="userName">Имя пользователя</param>
    /// <param name="userPassword">Пароль пользователя</param>
    /// <param name="uploadImagesBaseDirectory">Директория для загрузки фотографий</param>
    public FTP(string ftpHost, string userName, string userPassword, string uploadImagesBaseDirectory)
    {
      if (String.IsNullOrEmpty(ftpHost) || String.IsNullOrEmpty(userName))
        throw new ArgumentNullException(String.Format("неправильный формат имени хоста \"{0}\" или пользователя \"{1}\" FTP для экспорта", userName, userPassword));

      _FTPHost = ftpHost;
      _UserName = userName;
      _UserPassword = userPassword;
      _UploadImagesBaseDir = uploadImagesBaseDirectory;

      #region Содаю необходимые директории
      if (!String.IsNullOrEmpty(_UploadImagesBaseDir))
      {
        CreateFtpFolder(_UploadImagesBaseDir);
        CreateFtpFolder(Path.Combine(_UploadImagesBaseDir, FilesServerDir));
      }
      else
      {
        CreateFtpFolder(FilesServerDir);
      }

      CreateFtpFolder(Path.Combine(_UploadImagesBaseDir, SmallImageServerDir));
      CreateFtpFolder(Path.Combine(_UploadImagesBaseDir, AverageAndBigImageServerDir));
      #endregion
    }

    /// <summary>
    /// Выложить картинки на FTP
    /// </summary>
    /// <param name="presInfo">Информация о презентации</param>
    public void UploadImages(PresentationInfo presInfo)
    {
      if (presInfo == null)
        throw new ArgumentNullException("presentation info is not set");

      for (int index = 0; index < presInfo.SlidersInfo.Count; index++)
      {
        SlideInfo slideInfo = presInfo.SlidersInfo[index];

        if (!String.IsNullOrEmpty(slideInfo.ImageNameServerSmall))
        {
          UploadImage(slideInfo.ImageNameClientSmall, String.Format("{0}/{1}", SmallImageServerDir, slideInfo.ImageNameServerSmall));
        }

        if (!String.IsNullOrEmpty(slideInfo.ImageNameServerAverage))
        {
          CreateFtpFolder(Path.Combine(_UploadImagesBaseDir, AverageAndBigImageServerDir, presInfo.DbId.ToString()));
          CreateFtpFolder(Path.Combine(_UploadImagesBaseDir, AverageAndBigImageServerDir, presInfo.DbId.ToString(), "225"));
          UploadImage(slideInfo.ImageNameClientAverage, String.Format("{0}/{1}/225/{2}", AverageAndBigImageServerDir, presInfo.DbId, slideInfo.ImageNameServerAverage));
        }

        if (!String.IsNullOrEmpty(slideInfo.ImageNameServerBig))
        {
          CreateFtpFolder(Path.Combine(_UploadImagesBaseDir, AverageAndBigImageServerDir, presInfo.DbId.ToString()));
          CreateFtpFolder(Path.Combine(_UploadImagesBaseDir, AverageAndBigImageServerDir, presInfo.DbId.ToString(), "500"));
          UploadImage(slideInfo.ImageNameClientBig, String.Format("{0}/{1}/500/{2}", AverageAndBigImageServerDir, presInfo.DbId, slideInfo.ImageNameServerBig));
        }

        if (UploadImageCompleteCallback != null)
          UploadImageCompleteCallback(this, new UploadImageCompliteInfo { TotalImagesCount = presInfo.SlidersInfo.Count, CurrentImageNumber = index + 1 });
      }

      try
      {
        FtpWebRequest request = (FtpWebRequest)WebRequest.Create(String.Format("ftp://{0}/{1}/{2}", _FTPHost, Path.Combine(_UploadImagesBaseDir, FilesServerDir), presInfo.ServerFileName));

        request.UseBinary = true;
        request.Method = WebRequestMethods.Ftp.UploadFile;
        request.Credentials = new NetworkCredential(_UserName, _UserPassword);

        Stream requestStream = request.GetRequestStream();

        byte[] fileData = File.ReadAllBytes(Path.Combine(Directory.GetCurrentDirectory(), presInfo.ClientFilePath));
        requestStream.Write(fileData, 0, fileData.Length);

        requestStream.Close();
      }
      catch (Exception ex)
      {
        throw new Exception(String.Format("Во время загрузки файла с презентацией {0} на сервер, возникла ошибка: {1}", presInfo.ServerFileName, ex.Message));
      }
    }

    /// <summary>
    /// Загружает изображение на сервер
    /// </summary>
    /// <param name="uploadImagePath">Путь загрузки картинки на сервер с именем файла</param>
    /// <param name="clientImagePath">Путь к клиентсокой картинке с именем файла</param>
    private void UploadImage(string clientImagePath, string uploadImagePath)
    {
      if (String.IsNullOrEmpty(uploadImagePath) || String.IsNullOrEmpty(clientImagePath))
        throw new ArgumentNullException("upload image path or client image path is null");

      try
      {
        FtpWebRequest request = (FtpWebRequest)WebRequest.Create(String.Format("ftp://{0}/{1}/{2}", _FTPHost, _UploadImagesBaseDir, uploadImagePath));

        request.UseBinary = true;
        request.Method = WebRequestMethods.Ftp.UploadFile;
        request.Credentials = new NetworkCredential(_UserName, _UserPassword);

        Stream requestStream = request.GetRequestStream();

        byte[] fileData = File.ReadAllBytes(Path.Combine(Directory.GetCurrentDirectory(), PPTFiles._PresentationImageDir, clientImagePath));
        requestStream.Write(fileData, 0, fileData.Length);

        requestStream.Close();

        // Удаляю картинку с клиента в случаи успешной отправки на сервер
        File.Delete(Path.Combine(Directory.GetCurrentDirectory(), PPTFiles._PresentationImageDir, clientImagePath));
      }
      catch (Exception ex)
      {
        throw new Exception(String.Format("Во время загрузки картинки {0} на Ftp {1} возникла ошибка: {2}", clientImagePath, uploadImagePath, ex.Message));
      }
    }

    private void CreateFtpFolder(string folderPath)
    {
      if (String.IsNullOrEmpty(folderPath))
        throw new ArgumentNullException("При создании папки на FTP путь к папке не может быть пустым");

      if (_CreatedFtpDirs.Contains(folderPath))
        return;

      _CreatedFtpDirs.Add(folderPath);

      try
      {
        WebRequest request = WebRequest.Create(String.Format("ftp://{0}/{1}", _FTPHost, folderPath.Replace("\\", "/")));
        request.Method = WebRequestMethods.Ftp.MakeDirectory;
        request.Credentials = new NetworkCredential(_UserName, _UserPassword);
        request.GetResponse();
      }
      catch (WebException ex)
      {
        FtpWebResponse response = (FtpWebResponse)ex.Response;
        if (response.StatusCode != FtpStatusCode.ActionNotTakenFileUnavailable)
          throw new Exception(String.Format("При создании папки {0} на FTP произошла ошибка: {1}", folderPath, ex.Message));
      }
    }
  }

  /// <summary>
  /// Информация о завершении загрузки картинки по FTP
  /// </summary>
  public class UploadImageCompliteInfo : EventArgs
  {
    /// <summary>
    /// Текущий порядковый номер картинки
    /// </summary>
    public int CurrentImageNumber { get; set; }
    /// <summary>
    /// Общее количество картинок
    /// </summary>
    public int TotalImagesCount { get; set; }
  }
}
