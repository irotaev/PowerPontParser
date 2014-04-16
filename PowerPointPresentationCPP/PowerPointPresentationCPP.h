// PowerPointPresentationCPP.h

#pragma managed

using namespace System;
using namespace System::Collections::Generic;
using namespace System::Data;
using namespace MySql::Data::MySqlClient;

namespace PowerPointPresentationCPP 
{
  public ref class SlideInfo
  {
  public:
    int SlideNumber;
    String^ Text;
    String^ ImageNameSmall;
    String^ ImageNameAverage;
    String^ ImageNameBig;
  };

  // Информация о презентации
  public ref class PresentationInfo
  {
  private:
    String^ _UniqueId;

  public:
    PresentationInfo() : _UniqueId(Guid::NewGuid().ToString()) { }

    String^ UniqueId;
    List<SlideInfo^>^ SlidersInfo;
    String^ Name;
    String^ Title;
    String^ CategorieName;
    String^ CategorieCode;
  };

  // Отвечает за операции с MySql хостом
  public ref class MySqlPresentationHost
  {
  private:
    MySqlConnection^ _MySqlConnection;

    void PutInMainTable(PresentationInfo^ presInfo)
    {
      // Создаю таблицу в базе, если ее еще нет
      MySqlCommand^ command = _MySqlConnection->CreateCommand();
      command->CommandText = "CREATE TABLE IF NOT EXISTS `main`"
                             "("
                             "`Id` MEDIUMINT NOT NULL AUTO_INCREMENT"
                             "`url_news` CHAR(200) NOT NULL"
                             "`title` CHAR(200) NOT NULL"
                             ") CHARSET=cp1251";

      command->ExecuteNonQuery();
    }

  public:
    MySqlPresentationHost(String^ hostName, String^ dbName, String^ userName, String^ userPassword)
    {
      try
      {
        _MySqlConnection = gcnew MySqlConnection(String::Format("SERVER={0};DATABASE={1};UID={2};PASSWORD={3};CharSet=utf8", hostName, dbName, userName, userPassword));
        _MySqlConnection->Open();
      }
      catch(Exception^ ex)
      {
        throw gcnew Exception(String::Format("Ошибка соединения с базой данных: {0}", ex->Message));
      }
    }

    // Выложить данные презентации на сервер
    // @param presInfo Информация о презентации
    void PutDataOnServer(PresentationInfo^ presInfo)
    {
    }

    ~MySqlPresentationHost()
    {
      _MySqlConnection->Close();
    }
  };
}
