// ThisAddin.cs
// Librairie VisualVisioSVGLight
// Copyright © ShareVisual Michel LAPLANE
// All rights reserved.

//-------------------------------------------------------------------------//
//					TABLEAU DE BORD DES MISES A JOUR
//-------------------------------------------------------------------------//
//Modifié: V1.0  |   ML		| 00/00/2011 15:52:49  |
//-------------------------------------------------------------------------//
using Visio = Microsoft.Office.Interop.Visio;

namespace VisualVisioSVGLight
  {
  public partial class ThisAddIn
    {
    private Visio.Application visApplication;
    static public VisualVisioSVGLight addinApplication;
    static internal RibbonVisualVisioSVGLight ribbonApplication;
    string strCulture;

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
      {
      if (visApplication == null)
        {
        visApplication = (Microsoft.Office.Interop.Visio.Application)this.Application;
        }
      addinApplication = new VisualVisioSVGLight(visApplication, strCulture);
      }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
      {
      try
        {
        // Libération des évènements
        addinApplication.ReleaseEvents();
        visApplication = null;
        }
      catch
        {
        }
      }

    #region vsto ribbon support

    /// <summary>
    /// Fourni l'objet Ribbon de l'application au chargement de Visio
    /// </summary>
    /// <returns></returns>
    protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
      {
      ThisAddIn.ribbonApplication = new RibbonVisualVisioSVGLight(strCulture);
      return ThisAddIn.ribbonApplication;
      }


    #endregion

    #region Code généré par VSTO

    /// <summary>
    /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
    /// le contenu de cette méthode avec l'éditeur de code.
    /// </summary>
    private void InternalStartup()
      {
      this.Startup += new System.EventHandler(ThisAddIn_Startup);
      this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
      }

    #endregion

    }
  }
