// About.cs
// Librairie UtilAsm
// Copyright © ShareVisual Michel LAPLANE
// All rights reserved.

//-------------------------------------------------------------------------//
//					TABLEAU DE BORD DES MISES A JOUR
//-------------------------------------------------------------------------//
//Modifié: V1.0  |   ML		| 00/00/2011 15:52:49  |
//-------------------------------------------------------------------------//
using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisualVisioSVGLight
  {
  /// <summary>
  /// Description résumée de About.
  /// </summary>
  public class DlgAbout : System.Windows.Forms.Form
    {
    /// <summary>
    /// 
    /// </summary>
    public string strTemplateVersion;
    /// <summary>
    /// 
    /// </summary>
    public string strStencilVersion;
    private System.Windows.Forms.ListView lvVersionInfo;
    private System.Windows.Forms.ColumnHeader columnHeaderAssembly;
    private System.Windows.Forms.ColumnHeader columnHeaderVersion;
    private System.Windows.Forms.ColumnHeader columnHeaderDate;
    private System.Windows.Forms.Button closeButton;
    private System.Windows.Forms.Label labelProductName;
    private System.Windows.Forms.Label labelVersion;
    Assembly applicationAssembly;
    private Label labelCopyright;
    private Label labelCompanyName;
    private Label textBoxDescription;
    private Visio.Application visApp;
    private PictureBox pictureBox1;

    /// <summary>
    /// Variable nécessaire au concepteur.
    /// </summary>
    private System.ComponentModel.Container components = null;

    /// <summary>
    /// 
    /// </summary>
    public DlgAbout(Visio.Application visParamApp)
      {
      string[] arSplitVersion;

      //
      // Requis pour la prise en charge du Concepteur Windows Forms
      //
      visApp = visParamApp;
      InitializeComponent();
      applicationAssembly = Assembly.GetCallingAssembly();
      this.Text = String.Format("À propos de {0}", AssemblyTitle);
      this.labelProductName.Text = AssemblyProduct;
      arSplitVersion = AssemblyVersion.Split('.');
      this.labelVersion.Text = String.Format("Version {0}", arSplitVersion[0] + "." +
                                             arSplitVersion[1] + "." +
                                             arSplitVersion[2] + "." +
                                             arSplitVersion[3]);
      this.labelCopyright.Text = AssemblyCopyright;
      this.labelCompanyName.Text = AssemblyCompany;
      this.textBoxDescription.Text = AssemblyDescription;
      }

    /// <summary>
    /// 
    /// </summary>
    public void InitializeModuleName(ArrayList arModuleName, string strParamDocType, string strParamDocVersion,
                                     string strDataBaseVersion)
      {
      ArrayList arItems;
      Assembly[] arAssembly;
      AppDomain currentDomain;

      lvVersionInfo.Items.Clear();
      arItems = new ArrayList();
      try
        {
        currentDomain = AppDomain.CurrentDomain;
        arAssembly = currentDomain.GetAssemblies();
        arItems = new ArrayList();
        foreach (Assembly curAssembly in arAssembly)
          {
          AssemblyName curAssembyName;

          curAssembyName = curAssembly.GetName();
          ListViewItem item = new ListViewItem();
          item.Text = curAssembyName.Name;
          string versionStr = String.Format("{0}.{1}.{2}.{3}",
                                          curAssembyName.Version.Major.ToString(),
                                          curAssembyName.Version.Minor.ToString(),
                                            curAssembyName.Version.MajorRevision.ToString(),
                                            curAssembyName.Version.MinorRevision.ToString());
          item.SubItems.Add(versionStr);
          // Récupération information de date
          DateTime lastWriteDate = File.GetLastWriteTime(curAssembly.Location);
          string dateStr = lastWriteDate.ToString("g");
          item.SubItems.Add(dateStr);
          lvVersionInfo.Items.Add(item);
          // Rajout des modules
          if (arModuleName.Contains(curAssembyName.Name.ToUpper()))
            {
            arItems.Add(item);
            }
          }
        }
      catch
        {
        }
      try
        {
        // Récupération des modules
        //        ArrayList arItems = new ArrayList();
        Process tempProcess = Process.GetCurrentProcess();
        foreach (ProcessModule module in Process.GetCurrentProcess().Modules)
          {
          ListViewItem item = new ListViewItem();
          item.Text = module.ModuleName;
          // Récupération information de version
          try
            {
            FileVersionInfo verInfo = module.FileVersionInfo;
            string versionStr = String.Format("{0}.{1}.{2}.{3}",
                                            verInfo.FileMajorPart,
                                              verInfo.FileMinorPart,
                                              verInfo.FileBuildPart,
                                              verInfo.FilePrivatePart);
            item.SubItems.Add(versionStr);
            // Récupération information de date
            DateTime lastWriteDate = File.GetLastWriteTime(module.FileName);
            string dateStr = lastWriteDate.ToString("g");
            item.SubItems.Add(dateStr);
            }
          catch
            {
            }
          lvVersionInfo.Items.Add(item);
          //// Rajout des modules
          //if (arModuleName.Contains(module.ModuleName.ToUpper()))
          //  {
          //  arItems.Add(item);
          //  }

          }
        // Remise en forme de la liste pour mettre les modules de l'application
        // en tête de liste
        for (int i = arItems.Count; i > 0; i--)
          {
          ListViewItem item = (ListViewItem)arItems[i - 1];
          lvVersionInfo.Items.Remove(item);
          lvVersionInfo.Items.Insert(0, item);
          }
        }
      catch (Exception ex)
        {
        MessageBox.Show(this, ex.ToString(), "Erreur WUtilAsm", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
      }

    /// <summary>
    /// Nettoyage des ressources utilisées.
    /// </summary>
    protected override void Dispose(bool disposing)
      {
      if (disposing)
        {
        if (components != null)
          {
          components.Dispose();
          }
        }
      base.Dispose(disposing);
      }

    #region Code généré par le Concepteur Windows Form
    /// <summary>
    /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
    /// le contenu de cette méthode avec l'éditeur de code.
    /// </summary>
    private void InitializeComponent()
      {
      System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DlgAbout));
      this.closeButton = new System.Windows.Forms.Button();
      this.lvVersionInfo = new System.Windows.Forms.ListView();
      this.columnHeaderAssembly = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
      this.columnHeaderVersion = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
      this.columnHeaderDate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
      this.labelProductName = new System.Windows.Forms.Label();
      this.labelVersion = new System.Windows.Forms.Label();
      this.labelCopyright = new System.Windows.Forms.Label();
      this.labelCompanyName = new System.Windows.Forms.Label();
      this.textBoxDescription = new System.Windows.Forms.Label();
      this.SuspendLayout();
      // 
      // closeButton
      // 
      this.closeButton.DialogResult = System.Windows.Forms.DialogResult.OK;
      resources.ApplyResources(this.closeButton, "closeButton");
      this.closeButton.Name = "closeButton";
      this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
      // 
      // lvVersionInfo
      // 
      this.lvVersionInfo.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderAssembly,
            this.columnHeaderVersion,
            this.columnHeaderDate});
      this.lvVersionInfo.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
      this.lvVersionInfo.HideSelection = false;
      resources.ApplyResources(this.lvVersionInfo, "lvVersionInfo");
      this.lvVersionInfo.Name = "lvVersionInfo";
      this.lvVersionInfo.UseCompatibleStateImageBehavior = false;
      this.lvVersionInfo.View = System.Windows.Forms.View.Details;
      // 
      // columnHeaderAssembly
      // 
      resources.ApplyResources(this.columnHeaderAssembly, "columnHeaderAssembly");
      // 
      // columnHeaderVersion
      // 
      resources.ApplyResources(this.columnHeaderVersion, "columnHeaderVersion");
      // 
      // columnHeaderDate
      // 
      resources.ApplyResources(this.columnHeaderDate, "columnHeaderDate");
      // 
      // labelProductName
      // 
      resources.ApplyResources(this.labelProductName, "labelProductName");
      this.labelProductName.Name = "labelProductName";
      // 
      // labelVersion
      // 
      resources.ApplyResources(this.labelVersion, "labelVersion");
      this.labelVersion.Name = "labelVersion";
      // 
      // labelCopyright
      // 
      resources.ApplyResources(this.labelCopyright, "labelCopyright");
      this.labelCopyright.Name = "labelCopyright";
      // 
      // labelCompanyName
      // 
      resources.ApplyResources(this.labelCompanyName, "labelCompanyName");
      this.labelCompanyName.Name = "labelCompanyName";
      // 
      // textBoxDescription
      // 
      resources.ApplyResources(this.textBoxDescription, "textBoxDescription");
      this.textBoxDescription.Name = "textBoxDescription";
      // 
      // DlgAbout
      // 
      this.AcceptButton = this.closeButton;
      resources.ApplyResources(this, "$this");
      this.CancelButton = this.closeButton;
      this.Controls.Add(this.labelProductName);
      this.Controls.Add(this.lvVersionInfo);
      this.Controls.Add(this.closeButton);
      this.Controls.Add(this.textBoxDescription);
      this.Controls.Add(this.labelCompanyName);
      this.Controls.Add(this.labelCopyright);
      this.Controls.Add(this.labelVersion);
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
      this.Name = "DlgAbout";
      this.ShowInTaskbar = false;
      this.ResumeLayout(false);

      }
    #endregion

    private void closeButton_Click(object sender, System.EventArgs e)
      {
      this.Close();
      }

    private void DlgAbout_VisibleChanged(object sender, EventArgs e)
      {
      //// Mise à jour des versions
      //labStencilVersionValue.Text = strStencilVersion;
      //labTemplateVersionValue.Text = strTemplateVersion;
      }

    #region Accesseurs d'attribut de l'assembly

    /// <summary>
    /// 
    /// </summary>
    public string AssemblyTitle
      {
      get
        {
        // Obtenir tous les attributs Title de cet assembly
        object[] attributes = applicationAssembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
        // Si au moins un attribut Title existe
        if (attributes.Length > 0)
          {
          // Sélectionnez le premier
          AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
          // Si ce n'est pas une chaîne vide, le retourner
          if (titleAttribute.Title != "")
            return titleAttribute.Title;
          }
        // Si aucun attribut Title n'existe ou si l'attribut Title était la chaîne vide, retourner le nom .exe
        return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
        }
      }

    /// <summary>
    /// 
    /// </summary>
    public string AssemblyVersion
      {
      get
        {
        return applicationAssembly.GetName().Version.ToString();
        }
      }

    /// <summary>
    /// 
    /// </summary>
    public string AssemblyDescription
      {
      get
        {
        // Obtenir tous les attributs Description de cet assembly
        object[] attributes = applicationAssembly.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
        // Si aucun attribut Description n'existe, retourner une chaîne vide
        if (attributes.Length == 0)
          return "";
        // Si un attribut Description existe, retourner sa valeur
        return ((AssemblyDescriptionAttribute)attributes[0]).Description;
        }
      }

    /// <summary>
    /// 
    /// </summary>
    public string AssemblyProduct
      {
      get
        {
        // Obtenir tous les attributs Product de cet assembly
        object[] attributes = applicationAssembly.GetCustomAttributes(typeof(AssemblyProductAttribute), false);
        // Si aucun attribut Product n'existe, retourner une chaîne vide
        if (attributes.Length == 0)
          return "";
        // Si un attribut Product existe, retourner sa valeur
        return ((AssemblyProductAttribute)attributes[0]).Product;
        }
      }

    /// <summary>
    /// 
    /// </summary>
    public string AssemblyCopyright
      {
      get
        {
        // Obtenir tous les attributs Copyright de cet assembly
        object[] attributes = applicationAssembly.GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
        // Si aucun attribut Copyright n'existe, retourner une chaîne vide
        if (attributes.Length == 0)
          return "";
        // Si un attribut Copyright existe, retourner sa valeur
        return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
        }
      }

    /// <summary>
    /// 
    /// </summary>
    public string AssemblyCompany
      {
      get
        {
        // Obtenir tous les attributs Company de cet assembly
        object[] attributes = applicationAssembly.GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
        // Si aucun attribut Company n'existe, retourner une chaîne vide
        if (attributes.Length == 0)
          return "";
        // Si un attribut Company existe, retourner sa valeur
        return ((AssemblyCompanyAttribute)attributes[0]).Company;
        }
      }
    #endregion

    }
  }
