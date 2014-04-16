using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointPresentation;

namespace PowerPointPresentationTests
{
  [TestClass]
  public class PPTFileTest
  {
    [TestMethod]
    public void GetText()
    {
      PPTFiles pptFile = new PPTFiles();
      var result = pptFile.ExtractInfo("test.zip");
    }
  }
}
