using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WhosGotTheMasterSchedule
{
  internal class ExcelOwner
  {
    string _fileName;
    public const string NOT_BEING_EDITED = "No one";
    private ExcelOwner () {}

    public string Workbook 
    {
      private set { }
      get { return _fileName; }
    }
    private ExcelOwner (string fileName)
    {
      _fileName = fileName;
    }
    public static ExcelOwner Factory (String fileName)
    {
      return String.IsNullOrEmpty(fileName) ? null : new ExcelOwner(fileName);
    }

    public bool Exists
    {
      private set { }
      get { return File.Exists(_fileName); }
    }

    public bool Locked
    {
      private set { }
      get { return Owner != NOT_BEING_EDITED; }
    }
    public string Owner
    {
      private set { }
      get
      {
        if (!Exists)
          return NOT_BEING_EDITED;

        string returnValue = NOT_BEING_EDITED;
        FileInfo info = new FileInfo(GetXlTempFullFileName());
        try
        {
          returnValue = info.GetAccessControl().GetOwner(typeof (System.Security.Principal.NTAccount)).ToString();
        }
        catch
        {
          
        }
        return returnValue;
      }
    }

    private  string GetXlTempFullFileName()
    {
      string fileName = Path.GetFileName(_fileName);
      string directory = Path.GetDirectoryName(_fileName);
      string tempFileName = "~$" + fileName;
      return Path.Combine(directory, tempFileName);
    }

  }
}
