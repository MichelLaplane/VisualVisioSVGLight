// DlgOptions.cs
// Librairie VisualVisioSVGLight
// Copyright © Michel LAPLANE
// All rights reserved.

//-------------------------------------------------------------------------//
//					TABLEAU DE BORD DES MISES A JOUR
//-------------------------------------------------------------------------//
//Modifié: V1.0  |   ML		| 00/00/2011 15:52:49  |
//-------------------------------------------------------------------------//
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

#if STRINGASM
using StringAsm;
#endif

namespace VisualVisioSVGLight
  {
  public partial class DlgOptions : Form
    {

    private string strTemplatePath, strStencilPath, strProjectPath, strSystemPath;
    private string strStencilList;
    public string TemplatePath
      {
      get
        {
        return strTemplatePath;
        }
      set
        {
        strTemplatePath = value;
        }
      }

    public string StencilPath
      {
      get
        {
        return strStencilPath;
        }
      set
        {
        strStencilPath = value;
        }
      }

    public string ProjectPath
      {
      get
        {
        return strProjectPath;
        }
      set
        {
        strProjectPath = value;
        }
      }

    public string SystemPath
      {
      get
        {
        return strSystemPath;
        }
      set
        {
        strSystemPath = value;
        }
      }

    public string StencilList
      {
      get { return strStencilList; }
      set { strStencilList = value; }
      }

    public DlgOptions()
      {
      InitializeComponent();
      InitializeControl();
      }

    private void InitializeControl()
      {
      StencilPath = VisualVisioSVGLight.strStencilPath;
      TemplatePath = VisualVisioSVGLight.strTemplatePath;
      ProjectPath = VisualVisioSVGLight.strProjectPath;
      // Initialisation valeur par défaut
      edTemplatePath.Text = TemplatePath;
      edProjectPath.Text = ProjectPath;
      edStencilPath.Text = StencilPath;
      }

    private bool Explore(out string strSelectedPath)
      {
      FolderBrowserDialog dlgExplore;
      bool bSelected = false;

      strSelectedPath = "";
      // Affichage de la boîte de choix d'un répertoire
      dlgExplore = new FolderBrowserDialog(); 
      dlgExplore.SelectedPath = VisualVisioSVGLight.strProjectPath;
      if (dlgExplore.ShowDialog() == DialogResult.OK)
        {
        bSelected = true;
        strSelectedPath = dlgExplore.SelectedPath;
        }
      return bSelected;
      }

    private void btnTemplateExplore_Click(object sender, System.EventArgs e)
      {
      string strSelectedPath;

      if (Explore(out strSelectedPath))
        edTemplatePath.Text = strSelectedPath;
      }

    private void btnStencilExplore_Click(object sender, System.EventArgs e)
      {
      string strSelectedPath;

      if (Explore(out strSelectedPath))
        edStencilPath.Text = strSelectedPath;
      }

    private void btnProjectExplore_Click(object sender, System.EventArgs e)
      {
      string strSelectedPath;

      if (Explore(out strSelectedPath))
        edProjectPath.Text = strSelectedPath;
      }

    private void btnCancel_Click(object sender, System.EventArgs e)
      {
      Close();
      }

    private void btnOk_Click(object sender, System.EventArgs e)
      {
      TemplatePath = edTemplatePath.Text;
      ProjectPath = edProjectPath.Text;
      StencilPath = edStencilPath.Text;
      Close();
      }


    }
  }