using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointPresentation;

namespace PowerPointPresentationTests
{
  [TestClass]
  public class MySQLPresentationTableTest
  {
    [TestMethod]
    public void ConnectionTest()
    {
      PPTFiles pptFiles = new PPTFiles();

      using (MySQLPresentationTable presTable = new MySQLPresentationTable("mysql-srv35659.ht-systems.ru", "srv35659_test", "srv35659_tester", "re5ts2b"))
      {
        var presInfo = pptFiles.ExtractInfo(@"e:\Projects\PowerPointPresentation\PowerPointPresentationTests\bin\Debug\Presentations\test.ppt");
        presTable.PutDataOnServer(presInfo);
      }
    }
  }
}
