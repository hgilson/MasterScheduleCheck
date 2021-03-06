﻿using System;
using System.Drawing;
using System.Windows.Forms;
using System.Configuration;
using WhoHasTheMasterSchedule.Properties;
using WhosGotTheMasterSchedule;


namespace WhoHasTheMasterScheduleApp
{
  /// <summary>
  /// TODO: Need to add unit tests
  /// This is the primary class used for implementing a simple tray tool to indicate what user has locked a
  /// file.
  /// </summary>
  public partial class SysTrayApp : Form
  {
    private NotifyIcon _trayIcon;
    private ContextMenu _trayMenu;
    private string _oldOwner = ExcelOwner.NOT_BEING_EDITED;
    readonly Timer _timer = new Timer();
    private readonly ExcelOwner _masterSheetOwner;


    /// <summary>
    /// Initializes a new instance of the <see cref="SysTrayApp"/> class.
    /// </summary>
    public SysTrayApp()
    {
      _masterSheetOwner = ExcelOwner.Factory(ConfigurationManager.AppSettings["MasterSchedulePath"]);

      // Create a simple tray menu with only one item.
      _trayMenu = new ContextMenu();
      _trayMenu.MenuItems.Add("Exit", OnExit);
      _trayMenu.MenuItems.Add("Owner", ShowOwner);

      // Create a tray icon. In this example we use a
      // standard system icon for simplicity, but you
      // can of course use your own custom icon too.
      _trayIcon = new NotifyIcon {Text = "Who's Got The PTP Schedule"};
      SetIcon();

      // Add menu to tray icon and show it.
      _trayIcon.ContextMenu = _trayMenu;
      _trayIcon.Visible = true;

      // Create a _timer with a two second interval. 
      _timer.Tick += new EventHandler(OnTimer);
      _timer.Interval = 3000;
      _timer.Enabled = true;
    }

   /// <summary>
   /// Sets the balloon tip.
   /// Note: Do not call this unless the reference to _masterSheetOwner is valid.
   /// </summary>
   /// <param name="Owner">The owner.</param>
    private void SetBalloonTip()
    {
      _trayIcon.BalloonTipTitle = "PTP Master Schedule";
      _trayIcon.BalloonTipText = !_masterSheetOwner.Locked ? _masterSheetOwner.Workbook + "is available." : _masterSheetOwner.Owner + " has it open.";
      _trayIcon.BalloonTipIcon = ToolTipIcon.Info;
      _trayIcon.Visible = true;
      _trayIcon.ShowBalloonTip(30000);
    }


    /// <summary>
    /// Shows the owner of the file.
    /// </summary>
    /// <param name="myObject">My object.</param>
    /// <param name="myEventArgs">The <see cref="EventArgs"/> instance containing the event data.</param>
    private void ShowOwner(Object myObject, EventArgs myEventArgs)
     {
       MessageBox.Show(!_masterSheetOwner.Exists ? "File not found - no owner" : _masterSheetOwner.Owner + " has " + _masterSheetOwner.Workbook + " locked.");
     }

     /// <summary>
     /// Called when [timer].
     /// </summary>
     /// <param name="myObject">My object.</param>
     /// <param name="myEventArgs">The <see cref="EventArgs"/> instance containing the event data.</param>
    private void OnTimer(Object myObject, EventArgs myEventArgs) 
    {
      SetIcon();
       _timer.Enabled = true;

      if (_masterSheetOwner != null && 
          _masterSheetOwner.Exists && 
          _masterSheetOwner.Owner != _oldOwner)
      {
        SetBalloonTip();
        _oldOwner = _masterSheetOwner.Owner;
      }
    }

    private void SetIcon()
    {
      Icon iconForThisState = (_masterSheetOwner == null || _masterSheetOwner.Locked)
        ? Resources.RedLight : Resources.GreenLight;
      _trayIcon.Icon = new Icon(iconForThisState, 40, 40);
    }

    /// <summary>
    /// Raises the <see cref="E:System.Windows.Forms.Form.Load" /> event.
    /// </summary>
    /// <param name="e">An <see cref="T:System.EventArgs" /> that contains the event data.</param>
    protected override void OnLoad(EventArgs e)
    {
      Visible = false; 
      ShowInTaskbar = false; 

      base.OnLoad(e);
    }

    private void OnExit(object sender, EventArgs e)
    {
      Application.Exit();
    }
  }
}

