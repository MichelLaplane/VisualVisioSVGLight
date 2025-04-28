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


#if VISIOASM
    private AddAdvise documentAdvise = null;
#endif
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

    //public void InitializeModuleName(ArrayList arModuleName, string strParamDocType, string strParamDocVersion,
    //                                 string strDataBaseVersion)
    //  {
    //  ArrayList arItems;
    //  Assembly[] arAssembly;
    //  AppDomain currentDomain;

    //  lvVersionInfo.Items.Clear();
    //  arItems = new ArrayList();
    //  try
    //    {
    //    currentDomain = AppDomain.CurrentDomain;
    //    arAssembly = currentDomain.GetAssemblies();
    //    arItems = new ArrayList();
    //    foreach (Assembly curAssembly in arAssembly)
    //      {
    //      AssemblyName curAssembyName;

    //      curAssembyName = curAssembly.GetName();
    //      ListViewItem item = new ListViewItem();
    //      item.Text = curAssembyName.Name;
    //      string versionStr = String.Format("{0}.{1}.{2}.{3}",
    //                                      curAssembyName.Version.Major.ToString(),
    //                                      curAssembyName.Version.Minor.ToString(),
    //                                        curAssembyName.Version.MajorRevision.ToString(),
    //                                        curAssembyName.Version.MinorRevision.ToString());
    //      item.SubItems.Add(versionStr);
    //      // Récupération information de date
    //      DateTime lastWriteDate = File.GetLastWriteTime(curAssembly.Location);
    //      string dateStr = lastWriteDate.ToString("g");
    //      item.SubItems.Add(dateStr);
    //      lvVersionInfo.Items.Add(item);
    //      // Rajout des modules
    //      if (arModuleName.Contains(curAssembyName.Name.ToUpper()))
    //        {
    //        arItems.Add(item);
    //        }
    //      }
    //    }
    //  catch
    //    {
    //    }
    //  lvDocVersionInfo.Items.Clear();
    //  try
    //    {
    //    //if (strDataBaseVersion != "")
    //    //  this.labDatabaseVersion.Text = String.Format("Base de données version {0}", strDataBaseVersion);
    //    //else
    //    this.labDatabaseVersion.Visible = false;
    //    // Récupération des modules
    //    //        ArrayList arItems = new ArrayList();
    //    Process tempProcess = Process.GetCurrentProcess();
    //    foreach (ProcessModule module in Process.GetCurrentProcess().Modules)
    //      {
    //      ListViewItem item = new ListViewItem();
    //      item.Text = module.ModuleName;
    //      // Récupération information de version
    //      try
    //        {
    //        FileVersionInfo verInfo = module.FileVersionInfo;
    //        string versionStr = String.Format("{0}.{1}.{2}.{3}",
    //                                        verInfo.FileMajorPart,
    //                                          verInfo.FileMinorPart,
    //                                          verInfo.FileBuildPart,
    //                                          verInfo.FilePrivatePart);
    //        item.SubItems.Add(versionStr);
    //        // Récupération information de date
    //        DateTime lastWriteDate = File.GetLastWriteTime(module.FileName);
    //        string dateStr = lastWriteDate.ToString("g");
    //        item.SubItems.Add(dateStr);
    //        }
    //      catch
    //        {
    //        }
    //      lvVersionInfo.Items.Add(item);
    //      //// Rajout des modules
    //      //if (arModuleName.Contains(module.ModuleName.ToUpper()))
    //      //  {
    //      //  arItems.Add(item);
    //      //  }

    //      }
    //    // Remise en forme de la liste pour mettre les modules de l'application
    //    // en tête de liste
    //    for (int i = arItems.Count; i > 0; i--)
    //      {
    //      ListViewItem item = (ListViewItem)arItems[i - 1];
    //      lvVersionInfo.Items.Remove(item);
    //      lvVersionInfo.Items.Insert(0, item);
    //      }
    //    // Récupération version documents ouverts
    //    foreach (Visio.Document visCurDocument in visApp.Documents)
    //      {
    //      Visio.VisDocumentTypes visDocType;
    //      string strStencilVersion, strTemplateVersion;

    //      visDocType = (Visio.VisDocumentTypes)visCurDocument.Type;
    //      switch (visDocType)
    //        {
    //        case Visio.VisDocumentTypes.visTypeStencil:
    //          // Lecture version du gabarit
    //          if (VisMeth.IsDocumentUserCellExist(visCurDocument, strParamDocType, "StandardStencil") ||
    //              VisMeth.IsDocumentUserCellExist(visCurDocument, "safeprojectnameDocType", "StandardStencil"))
    //            {
    //            ListViewItem item = new ListViewItem();
    //            item.Text = visCurDocument.Name;
    //            item.SubItems.Add("Stencil");
    //            try
    //              {
    //              VisMeth.GetStringCellUser(visCurDocument, "User." + strParamDocVersion, out strStencilVersion);
    //              if (strStencilVersion == "")
    //                {
    //                VisMeth.GetStringCellUser(visCurDocument, "User." + "safeprojectnameVersion", out strStencilVersion);
    //                }
    //              item.SubItems.Add(strStencilVersion);
    //              lvDocVersionInfo.Items.Add(item);
    //              }
    //            catch
    //              {
    //              }
    //            }
    //          break;
    //        case Visio.VisDocumentTypes.visTypeDrawing:
    //          // Lecture version du modèle
    //          if (VisMeth.IsDocumentUserCellExist(visCurDocument, strParamDocType, "StandardDoc") ||
    //              VisMeth.IsDocumentUserCellExist(visCurDocument, "safeprojectnameDocType", "StandardDoc"))
    //            {
    //            ListViewItem item = new ListViewItem();
    //            item.Text = visCurDocument.Name;
    //            item.SubItems.Add("Document");
    //            try
    //              {
    //              VisMeth.GetStringCellUser(visCurDocument, "User." + strParamDocVersion, out strTemplateVersion);
    //              if (strTemplateVersion == "")
    //                {
    //                VisMeth.GetStringCellUser(visCurDocument, "User." + "safeprojectnameVersion", out strTemplateVersion);
    //                }
    //              item.SubItems.Add(strTemplateVersion);
    //              lvDocVersionInfo.Items.Add(item);
    //              }
    //            catch
    //              {
    //              }
    //            }
    //          break;
    //        default:
    //          break;
    //        }
    //      }
    //    }
    //  catch (Exception ex)
    //    {
    //    MessageBox.Show(this, ex.ToString(), "Erreur WUtilAsm", MessageBoxButtons.OK, MessageBoxIcon.Error);
    //    }
    //  }


    }
  }
