// VisualVisioSVGLight.cs
// Librairie VisualVisioSVGLight
// Copyright © ShareVisual Michel LAPLANE
// All rights reserved.

//-------------------------------------------------------------------------//
//					TABLEAU DE BORD DES MISES A JOUR
//-------------------------------------------------------------------------//
//Modifié: V1.0  |   ML		| 00/00/2011 15:52:49  |
//-------------------------------------------------------------------------//

using System;
using System.Collections;
using System.Data.Common;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;
using Visio = Microsoft.Office.Interop.Visio;
using System.Configuration;

#if STRINGASM
using StringAsm;
#endif
#if UTILASM
using UtilAsm;
using UtlMeth = UtilAsm.ULMethods;
using OLMeth = OfficeAsm.OLMethods;
#endif
#if VISIOASM
using VisioAsm;
using VisMeth = VisioAsm.VLMethods;
using VisCst = VisioAsm.VLConstants;
#endif
#if DATAASM
using DataAsm;
#endif
using System.Security.Principal;

namespace VisualVisioSVGLight
  {
  /// <summary>
  /// Description résumée de VisualVisioSVGLight.
  /// </summary>
  public class VisualVisioSVGLight
    {
    #region constantes
    static internal int OL_SHORTTEXT = 25;
    #endregion
    public static string strLoggedUserName = "";
    public static string strSystemPath, strStencilPath, strTemplatePath, strProjectPath, strStencilName, strExcelTemplatelName, strStencilList;
    public Microsoft.Office.Interop.Visio.Application visApplication;
    public static string strCulture;
    internal FrmSVG frmSVG = null;


#if VISIOASM
    private AddAdvise documentAdvise = null;
#endif
    private Microsoft.Office.Interop.Visio.Document visDocument = null;
    private Microsoft.Office.Interop.Visio.Document visStencil = null;
    private bool bNewFile = false;
    internal static DbConnection dbMainConnection;
    internal static int iMainProvider;
    public string strSPServerName = "", strSPPortNum = "", strSPSiteName = "", strSPLibraryName = "";
    public string strPublishMode = "", strFolderName = "";
    public string strUserMode;
    private string strDataBasePassword = "";
    private bool bDatabaseConnectionOK = false;
    private string strApplicationProvider, strApplicationMode, strApplicationDatabasePath, strApplicationDataSource, strApplicationInitialCatalog;
    private string strApplicationUser, strApplicationPassword, strDataBaseVersion;

    private string strOptions, strCompanyName, strApplicationName;

    static internal ArrayList projectList = new ArrayList();
    public object curMarkerEventApplicationDoc = null;

    internal static bool bNewProjectInProgress;
    internal static int iLastNewProjectNum;
    public static bool bOpenProjectInProgress;
    public static bool bPropInRefresh;

    private bool bMustSetTitle;

    // Reporting
    static internal Visio.Window visWindowReport = null;
    public string strTitle;

    public Microsoft.Office.Interop.Visio.Application VisApplication
      {
      get { return visApplication; }
      set { visApplication = value; }
      }

    public VisualVisioSVGLight(Microsoft.Office.Interop.Visio.Application theApplication, string strParamCulture)
      {
      visApplication = theApplication;
      strCulture = strParamCulture;
      CultureInfo cultInfo = new CultureInfo(visApplication.Language, false);
      // lecture des paths de l'application
      strStencilName = "VisualVisioSVGLight.vssx";
      strStencilPath = ConfigurationManager.AppSettings["StencilsPath"];
      strTemplatePath = ConfigurationManager.AppSettings["TemplatesPath"];
      strProjectPath = ConfigurationManager.AppSettings["ProjectsPath"];

      }

    public void InitializeMember(Visio.Document visDocument)
      {
      if (this.visDocument != visDocument)
        {
        this.visDocument = visDocument;
        }
      }

    public void ReleaseEvents()
      {
#if VISIOASM
      // suppression des évènements
      if (documentAdvise != null)
        {
        documentAdvise.Dispose(true);
        documentAdvise = null;
        }
#endif
      }

    /// <summary>
    /// Creation of a new document VisualVisioSVGLight
    /// </summary>
    public void NewFile()
      {
      string strFullTemplateFilename, strFullStencilName;

      try
        {
        Cursor.Current = Cursors.WaitCursor;
        strFullTemplateFilename = Path.Combine(strTemplatePath, "VisualVisioSVGLight.vstx");
        strFullStencilName = Path.Combine(strStencilPath, "VisualVisioSVGLight.vssx");
        visDocument = visApplication.Documents.OpenEx(strFullTemplateFilename, (short)Visio.VisOpenSaveArgs.visOpenCopy);
        visStencil = visApplication.Documents.OpenEx(strFullStencilName,
          (short)Visio.VisOpenSaveArgs.visOpenRO
          + (short)Visio.VisOpenSaveArgs.visOpenMinimized
          + (short)Visio.VisOpenSaveArgs.visOpenDocked
          + (short)Visio.VisOpenSaveArgs.visOpenNoWorkspace);
        }
      catch (Exception excep)
        {
        }
      finally
        {
        Cursor.Current = Cursors.Default;
        }
      }

    internal void DisplaySVGWindow()
      {
      frmSVG = new FrmSVG(visApplication);
      // Make it child of the Visio Window
      //int windowHandle = frmMermaid.Handle.ToInt32();
      //NativeMethods.SetParent(windowHandle, this.visApplication.WindowHandle32);
      frmSVG.Show();
      }

    public void About()
      {
        DlgAbout dlgAbout;
        string strTemplateVersion = "?";
        string strStencilVersion = "?";
        ArrayList arModuleName = new ArrayList();

        dlgAbout = new DlgAbout(visApplication);
        arModuleName.Add("VisualVisioMermaid".ToUpper());
        arModuleName.Add("ControlAsm".ToUpper());
        arModuleName.Add("DataAsm".ToUpper());
        arModuleName.Add("GeoAsm".ToUpper());
        arModuleName.Add("OfficeAsm".ToUpper());
        arModuleName.Add("StringAsm".ToUpper());
        arModuleName.Add("VisioAsm".ToUpper());
        dlgAbout.InitializeModuleName(arModuleName, "VisualVisioMermaidDocType", "VisualVisioMermaidDocVersion", strDataBaseVersion);
        dlgAbout.strTemplateVersion = strTemplateVersion;
        dlgAbout.strStencilVersion = strStencilVersion;
        dlgAbout.ShowDialog();
      }


    }
  }
