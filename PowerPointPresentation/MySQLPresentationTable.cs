using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;

namespace PowerPointPresentation
{
  public class MySQLPresentationTable : AbstractDBTable
  {
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
      long id = 0;

      MySqlCommand command = _MySqlConnection.CreateCommand();
      command.CommandText = String.Format("SELECT `auto_increment` FROM INFORMATION_SCHEMA.TABLES WHERE table_name = '{0}'", _TableName);
      //command.CommandText = String.Format("SELECT id FROM `{0}` ORDER BY `id` DESC", _TableName);

      try
      {
        using (var reader = command.ExecuteReader())
        {
          if (reader.HasRows)
          {
            while (reader.Read() && id == 0)
            {
              id = Int64.Parse(reader["auto_increment"].ToString());
            }
          }
        }
      }
      catch
      {
        id = 0;
      }

      return id;
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

    public override void PutDataOnServer(PresentationInfo presInfo)
    {
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

      #region Заполнение таблицы
      try
      {
        MySqlCommand command = _MySqlConnection.CreateCommand();
        command.CommandText = String.Format(new System.Globalization.CultureInfo("en-GB"), @"
          INSERT INTO `{6}` (`naz`, `title`, `size`, `slides`, `content`, `login`)
           VALUES ('{0}', '{1}', '{2:0.00}', '{3}', '{4}', '{5}')
        ",
         SecurityElement.Escape(presInfo.Name),
         SecurityElement.Escape(presInfo.Title),
         Convert.ToSingle(presInfo.FileSize / 1024 / 1024, System.Globalization.CultureInfo.InvariantCulture),
         SecurityElement.Escape(presInfo.SlidersInfo.Count.ToString()),
         SecurityElement.Escape(FormContentDbColumn(presInfo)),
         SecurityElement.Escape(presInfo.Login),
         SecurityElement.Escape(_TableName));

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
