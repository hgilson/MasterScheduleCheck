using System;
using System.IO;

namespace WhosGotTheMasterSchedule
{
  /// <summary>
  /// Used to view status of a Excel file on a share drive.
  /// -- Done by looking @ owner of the "lock file"
  /// </summary>
  internal class ExcelOwner
  {
    private readonly string _fileName;
    public const string NOT_BEING_EDITED = "No one";

    private ExcelOwner()
    {
    }

    /// <summary>
    /// Gets or sets the workbook name.
    /// </summary>
    /// <value>
    /// The workbook.
    /// </value>
    public string Workbook
    {
      private set { }
      get { return _fileName; }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelOwner"/> class.
    /// </summary>
    /// <param name="fileName">Name of the file.</param>
    private ExcelOwner(string fileName)
    {
      _fileName = fileName;
    }

    /// <summary>
    /// Only method to construct an instance.
    /// </summary>
    /// <param name="fileName">Name of the Excel file to monitor.</param>
    /// <returns></returns>
    public static ExcelOwner Factory(String fileName)
    {
      return String.IsNullOrEmpty(fileName) ? null : new ExcelOwner(fileName);
    }

    /// <summary>
    /// Gets or sets a value indicating whether the file used to initialize <see cref="ExcelOwner"/> exists.
    /// </summary>
    /// <value>
    ///   <c>true</c> if exists; otherwise, <c>false</c>.
    /// </value>
    public bool Exists
    {
      private set { }
      get { return File.Exists(_fileName); }
    }

    /// <summary>
    /// Gets or sets a value indicating whether this <see cref="ExcelOwner"/> is locked.
    /// </summary>
    /// <value>
    ///   <c>true</c> if locked; otherwise, <c>false</c>.
    /// </value>
    public bool Locked
    {
      private set { }
      get { return Owner != NOT_BEING_EDITED; }
    }

    /// <summary>
    /// Gets the owner of the lock file.
    /// </summary>
    /// <value>
    /// The owner.
    /// </value>
    public string Owner
    {
      private set { }
      get
      {
        if (!Exists)
          return NOT_BEING_EDITED;

        string returnValue = NOT_BEING_EDITED;
        FileInfo info = new FileInfo(GetXlsTempFullFileName());
        try
        {
          returnValue = info.GetAccessControl().GetOwner(typeof(System.Security.Principal.NTAccount)).ToString();
        }
        catch
        {
        }
        return returnValue;
      }
    }

    /// <summary>
    /// Gets the name of the xls temporary full file based on the file to monitor.
    /// </summary>
    /// <returns></returns>
    private string GetXlsTempFullFileName()
    {
      string fileName = Path.GetFileName(_fileName);
      string directory = Path.GetDirectoryName(_fileName);
      string tempFileName = "~$" + fileName;
      return Path.Combine(directory, tempFileName);
    }
  }
}