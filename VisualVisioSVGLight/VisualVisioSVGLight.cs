// VisualVisioSVGLight.cs
// Librairie VisualVisioSVGLight
// Copyright © Michel LAPLANE
// All rights reserved.

//-------------------------------------------------------------------------//
//					TABLEAU DE BORD DES MISES A JOUR
//-------------------------------------------------------------------------//
//Modifié: V1.0  |   ML		| 00/00/2011 15:52:49  |
//-------------------------------------------------------------------------//

using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;
using Visio = Microsoft.Office.Interop.Visio;
using System.Configuration;


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
    public static string strStencilPath, strTemplatePath, strProjectPath, strStencilName, strExcelTemplatelName, strStencilList;
    private string strApplicationName = "VisualVisioSVGLight";
    private string strPathKey = "Path";
    private string strStencilPathKey = "Stencils";
    private string strStencilDefaultPath = @"C:\Users\Documents\VisualVisioSVGLight\Stencils";
    private string strTemplatePathKey = "Templates";
    private string strTemplatelDefaultPath = @"C:\Users\Documents\VisualVisioSVGLight\Templates";
    private string strProjectPathKey = "Projects";
    private string strProjectDefaultPath = @"C:\Users\Documents\VisualVisioSVGLight\Projects";


    public Microsoft.Office.Interop.Visio.Application visApplication;
    public static string strCulture;
    internal FrmSVG frmSVG = null;

    private Microsoft.Office.Interop.Visio.Document visDocument = null;
    private Microsoft.Office.Interop.Visio.Document visStencil = null;
    private bool bNewFile = false;

    private string strOptions;

    static internal ArrayList projectList = new ArrayList();
    public object curMarkerEventApplicationDoc = null;

    internal static bool bNewProjectInProgress;
    internal static int iLastNewProjectNum;
    public static bool bOpenProjectInProgress;
    public static bool bPropInRefresh;

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
      ReadRegistryPathSection(strApplicationName, strPathKey,
                              strStencilPathKey, strStencilDefaultPath, out strStencilPath,
                              strTemplatePathKey, strTemplatelDefaultPath, out strTemplatePath,
                              strProjectPathKey, strProjectDefaultPath, out strProjectPath);
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

    public void OpenFile()
      {
      string strFullFilename;

      try
        {
        Cursor.Current = Cursors.WaitCursor;
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Title = "Open a diagram";
        openFileDialog.Filter = "Drawing(*.vsdx; *.vsd; *.vdx)| *.vsdx; *.vsd; *.vdx";
        openFileDialog.FilterIndex = 1;  // 1 based index
        if (openFileDialog.ShowDialog() == DialogResult.OK)
          {
          Cursor.Current = Cursors.WaitCursor;

          strFullFilename = openFileDialog.FileName;
          Cursor.Current = Cursors.WaitCursor;
          visDocument = visApplication.Documents.Open(strFullFilename);
          }
        }
      catch
        {
        }
      finally
        {
        Cursor.Current = Cursors.Default;
        }
      }

    public void SaveFile()
      {
      if (visDocument.Path == "")
        {
        // Not already saved
        SaveAsFile();
        }
      else
        {
        try
          {
          visDocument.Save();
          }
        catch
          {
          }
        }
      }

    public void SaveAsFile()
      {
      SaveFileDialog saveFileDialog = new SaveFileDialog();
      saveFileDialog.Title = "Save diagram";
      saveFileDialog.Filter = "Drawing(*.vsdx; *.vsd; *.vdx)| *.vsdx; *.vsd; *.vdx";
      saveFileDialog.FilterIndex = 1;  // 1 based index
      // Affiche la boite de sélection du logigramme
      if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
        string strFileName;

        strFileName = saveFileDialog.FileName;
        try
          {
          visDocument.SaveAs(strFileName);
          }
        catch
          {
          }
        finally
          {
          }
        }
      }


    public void CloseFile()
      {
      visDocument.Close();
      }

    public static void ReadRegistryPathSection(string strApplication, string strPathKey,
                                       string strStencilPathKey, string strStencilDefaultValue, out string strStencilPath,
                                       string strTemplatePathKey, string strTemplateDefaultValue, out string strTemplatePath,
                                       string strProjectPathKey, string strProjectDefaultValue, out string strProjectPath)
      {
      RegistryKey regKey, regApplicationKey, regFieldKey;

      strStencilPath = strStencilDefaultValue;
      strTemplatePath = strTemplateDefaultValue;
      strProjectPath = strProjectDefaultValue;
      try
        {
        // Création d'une clé pour accéder à la clé HKEY_CURRENT_USER
        // de la base de registe
        regKey = Registry.CurrentUser;
        if ((regApplicationKey = regKey.CreateSubKey("Software\\" + strApplication)) != null)
          {
          // Clé path
          if ((regFieldKey = regApplicationKey.CreateSubKey(strPathKey)) != null)
            {
            // Stencils path	
            if ((strStencilPath = (String)regFieldKey.GetValue(strStencilPathKey)) == null)
              {
              regFieldKey.SetValue(strStencilPathKey, strStencilDefaultValue);
              strStencilPath = strStencilDefaultValue;
              }
            // Template path	
            if ((strTemplatePath = (String)regFieldKey.GetValue(strTemplatePathKey)) == null)
              {
              regFieldKey.SetValue(strTemplatePathKey, strTemplateDefaultValue);
              strTemplatePath = strTemplateDefaultValue;
              }
            // Projects path	
            if ((strProjectPath = (String)regFieldKey.GetValue(strProjectPathKey)) == null)
              {
              regFieldKey.SetValue(strProjectPathKey, strProjectDefaultValue);
              strProjectPath = strProjectDefaultValue;
              }
            }
          }
        }
      catch
        {
        }
      }

    public static void UpdateRegistryInfos(string strApplication, string strPathKey,
                       string strStencilPathKey, string strStencilValue,
                       string strTemplatePathKey, string strTemplateValue,
                       string strProjectPathKey, string strProjectValue)
      {
      RegistryKey regKey, regCompanyKey, regApplicationKey, regFieldKey;

      // Création d'une clé pour accéder à la clé HKEY_CURRENT_USER
      // de la base de registe
      regKey = Registry.CurrentUser;
      if ((regApplicationKey = regKey.CreateSubKey("Software\\" + strApplication)) != null)
        {
        // Clé path
        if ((regFieldKey = regApplicationKey.CreateSubKey(strPathKey)) != null)
          {
          // Stencils path
          regFieldKey.SetValue(strStencilPathKey, strStencilValue);
          // Template path
          regFieldKey.SetValue(strTemplatePathKey, strTemplateValue);
          // Project path
          regFieldKey.SetValue(strProjectPathKey, strProjectValue);
          }
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

    public void Options()
      {
      DlgOptions dlgOptions;

      dlgOptions = new DlgOptions();
      if (dlgOptions.ShowDialog() == DialogResult.OK)
        {
        strStencilPath = dlgOptions.StencilPath;
        strTemplatePath = dlgOptions.TemplatePath;
        strProjectPath = dlgOptions.ProjectPath;
        Configuration configurationFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        configurationFile.AppSettings.Settings["StencilsPath"].Value = dlgOptions.StencilPath;
        configurationFile.AppSettings.Settings["TemplatesPath"].Value = dlgOptions.TemplatePath;
        configurationFile.AppSettings.Settings["ProjectsPath"].Value = dlgOptions.ProjectPath;
        UpdateRegistryInfos(strApplicationName, strPathKey,
                           strStencilPathKey, dlgOptions.StencilPath,
                           strTemplatePathKey, dlgOptions.TemplatePath,
                           strProjectPathKey, dlgOptions.ProjectPath);
        }
      }

    public void About()
      {
      DlgAbout dlgAbout;
      string strTemplateVersion = "?";
      string strStencilVersion = "?";
      ArrayList arModuleName = new ArrayList();

      dlgAbout = new DlgAbout(visApplication);
      arModuleName.Add("VisualVisioSVGLight".ToUpper());
      dlgAbout.InitializeModuleName(arModuleName, "VisualVisioSVGLightDocType", "VisualVisioSVGLightVersion");
      dlgAbout.strTemplateVersion = strTemplateVersion;
      dlgAbout.strStencilVersion = strStencilVersion;
      dlgAbout.ShowDialog();
      }
    }
  }
