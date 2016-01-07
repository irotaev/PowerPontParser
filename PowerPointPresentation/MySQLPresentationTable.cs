using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;

namespace PowerPointPresentation
{
  public class MySQLPresentationTable : AbstractDBTable
  {
    public static volatile bool IsTableCreated;

    /// <summary>
    /// Создать объект таблицы презентации
    /// </summary>
    /// <param name="hostName">Имя хоста</param>
    /// <param name="dbName">Имя базы данных</param>
    /// <param name="userName">Имя пользователя</param>
    /// <param name="userPassword">Пароль пользователя</param>
    public MySQLPresentationTable(string hostName, string dbName, string userName, string userPassword) : base(hostName, dbName, userName, userPassword, "present") { }

    #region Methods
    public override long GetCurrentPresentationIndex()
    {
      try
      {
        MySqlCommand command = _MySqlConnection.CreateCommand();
        command.CommandText = String.Format(new System.Globalization.CultureInfo("en-GB"), @"
          INSERT INTO `{0}` (`naz`, `title`, `size`, `slides`, `content`, `login`, `url`, `cat`)
           VALUES ('', '', '', '', '', '', '', '')
        ", SecurityElement.Escape(_TableName));

        command.ExecuteNonQuery();

        return command.LastInsertedId;
      }
      catch (Exception ex)
      {
        throw new Exception(String.Format("Во время заполнения таблицы 'main' презентацией '{0}' произошла ошибка: {1}", ex.Message));
      }
    }

    private string FormContentDbColumn(PresentationInfo presInfo)
    {
      string result = null;

      for (int index = 0; index < presInfo.SlidersInfo.Count; index++)
      {
        result += String.Format("<div class='slide-block'>" +
                                   "<h3>Слайд {0}</h3><!-- slide-title -->" +

                                    "<div class='all-sl-img'>" +
                                      "<img src='/files/{1}/268/{2}' alt='' />" +
                                    "</div><!-- all-sl-img -->" +

                                    "<div class='all-sl-txt'>{3}</div><!-- all-sl-txt -->" +
                                 "</div><!-- slide-block -->",
                                 (index + 1), presInfo.DbId, presInfo.SlidersInfo[index].ImageNameClientAverage, presInfo.SlidersInfo[index].Text);
      }

      return result;
    }

    public void CreateTable()
    {
      if (IsTableCreated) return;

      #region Создание таблицы, если это необходимо
      try
      {
        {
          MySqlCommand command = _MySqlConnection.CreateCommand();
          command.CommandText = String.Format(@"CREATE TABLE IF NOT EXISTS `{0}`
                              (
                              `id` MEDIUMINT NOT NULL AUTO_INCREMENT,
                              `naz` VARCHAR(255) NOT NULL,
                              `title` VARCHAR(255) NULL,
                              `size` FLOAT NOT NULL,
                              `slides` SMALLINT NOT NULL,
                              `content` TEXT(1000) NOT NULL,  
                              `login` VARCHAR(100) NULL,  
                              `url` VARCHAR(255) NOT NULL,
                              `like` MEDIUMINT NULL,
                              `count` BIGINT NULL,                           
                              `cat` VARCHAR(255) NOT NULL,
                              UNIQUE KEY (id)
                              ) CHARSET={1}", SecurityElement.Escape(_TableName), TABLE_CHARSET);
          command.ExecuteNonQuery();
        }
      }
      catch (Exception ex)
      {
        throw new Exception(String.Format("Во время создания таблицы {0} произошла ошибка: {1}", _TableName, ex.Message));
      }
      #endregion

      IsTableCreated = true;
    }

    public override void PutDataOnServer(PresentationInfo presInfo)
    {      
      #region Заполнение таблицы
      try
      {
        MySqlCommand command = _MySqlConnection.CreateCommand();
//        command.CommandText = String.Format(new System.Globalization.CultureInfo("en-GB"), @"
//          INSERT INTO `{6}` (`naz`, `title`, `size`, `slides`, `content`, `login`)
//           VALUES ('{0}', '{1}', '{2:0.00}', '{3}', '{4}', '{5}')
//        ",
//         SecurityElement.Escape(presInfo.Name),
//         SecurityElement.Escape(presInfo.Title),
//         Convert.ToSingle(presInfo.FileSize / 1024 / 1024, System.Globalization.CultureInfo.InvariantCulture),
//         SecurityElement.Escape(presInfo.SlidersInfo.Count.ToString()),
//         SecurityElement.Escape(FormContentDbColumn(presInfo)),
//         SecurityElement.Escape(presInfo.Login),
//         SecurityElement.Escape(_TableName));
        command.CommandText = String.Format(new System.Globalization.CultureInfo("en-GB"), @"
          UPDATE `{6}` SET `naz`='{0}', `title`='{1}', `size`='{2:0.00}', `slides`='{3}', `content`='{4}', `login`='{5}', `url`='{8}', `cat`='{9}'
           WHERE `id`='{7}'
        ",
         SecurityElement.Escape(presInfo.Name),
         SecurityElement.Escape(presInfo.Title),
         Convert.ToSingle(presInfo.FileSize / 1024 / 1024, System.Globalization.CultureInfo.InvariantCulture),
         SecurityElement.Escape(presInfo.SlidersInfo.Count.ToString()),
         SecurityElement.Escape(FormContentDbColumn(presInfo)),
         SecurityElement.Escape(presInfo.Login),
         SecurityElement.Escape(_TableName),
         SecurityElement.Escape(presInfo.DbId.ToString()),
         SecurityElement.Escape(Regex.Replace(presInfo.Name, @"[^\da-zA-Zа-яА-Я]", "_")),
         SecurityElement.Escape(presInfo.Categorie.Key));
         //SecurityElement.Escape(new JoZhTranslit.Transliterator(JoZhTranslit.TransliterationMaps.EnRu.MapJson).Transliterate(presInfo.Name)));

        command.ExecuteNonQuery();
      }
      catch (Exception ex)
      {
        throw new Exception(String.Format("Во время заполнения таблицы 'main' презентацией '{0}' произошла ошибка: {1}", presInfo.Name, ex.Message));
      }
      #endregion
    }
    #endregion
  }
}
