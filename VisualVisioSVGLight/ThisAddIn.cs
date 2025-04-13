// ThisAddin.cs
// Librairie VisualVisioSVGLight
// Copyright © ShareVisual Michel LAPLANE
// All rights reserved.

//-------------------------------------------------------------------------//
//					TABLEAU DE BORD DES MISES A JOUR
//-------------------------------------------------------------------------//
//Modifié: V1.0  |   ML		| 00/00/2011 15:52:49  |
//-------------------------------------------------------------------------//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using System.Globalization;
using System.Threading;
#if OFFICEASM
using OfficeAsm;
#endif
#if STRINGASM
using StringAsm;
#endif
#if VISIOASM
using VisioAsm;
using VisMeth = VisioAsm.VLMethods;
using VisCst = VisioAsm.VLConstants;
#endif
#if UTILASM
using UtilAsm;
using UtlMeth = UtilAsm.ULMethods;
using OLMeth = OfficeAsm.OLMethods;
#endif
using System.Collections;

namespace VisualVisioSVGLight
  {
#if VISIOASM
  public partial class ThisAddIn : IEventCallbackAppOnly
    {
    private AddAdviseAppOnly applicationAdvise = null;
#else
  public partial class ThisAddIn
    {
#endif
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
      #region déclaration évènements
#if VISIOASM
      applicationAdvise = new AddAdviseAppOnly(visApplication);
      applicationAdvise.SetAddAdviseAppOnly(this,StringEx.GetString("idsApplicationName"));
#endif
      #endregion
      }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
      {
      try
        {
        // Libération des évènements
        addinApplication.ReleaseEvents();
#if VISIOASM
        applicationAdvise.Dispose(true);
#endif
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
#if UTILASM
                UtlMeth.ReadRegistrySection(StringEx.GetString("idsCompanyName"), StringEx.GetString("idsApplicationName"),
                                            StringEx.GetString("idsOptionsKey"),
                                            StringEx.GetString("idsOptionsKeyCulture"),
                                            StringEx.GetString("idsOptionsKeyCultureDefaultValue"),
                                            out strCulture);
#endif
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

#if VISIOASM
    #region fonctions évènementielles

      /// <summary>
      /// Au idle de Visio
      /// </summary>
      /// <param name="visApplication"></param>
      public void OnVisioIdle(Visio.Application visApplication)
        {
        }

      /// <summary>
      /// A la demande de fermeture de l'application
      /// </summary>
      /// <param name="visApplication"></param>
      public void OnVisioQueryApplicationQuit(Visio.Application visApplication)
        {
        }

      /// <summary>
      /// Après qu'une instance visio soit devenu active
      /// </summary>
      /// <param name="visApplication"></param>
      public void OnVisioAppObjActivated(Visio.Application visApplication)
        {
        }

      /// <summary>
      /// Avant la fermeture de l'application
      /// </summary>
      /// <param name="visApplication"></param>
      public void OnVisioBeforeApplicationQuit(Visio.Application visApplication)
        {
        }
    
      /// <summary>
      /// Réception d'un message keystroke (touche enfoncée) de  Windows 
      /// destiné a une fenêtre Add-on (fenêtre dans l'espace Visio) ou une fenêtre enfant d'une
      /// fenêtre Add-on  
      /// </summary>
      /// <param name="visPage"></param>
      public bool OnKeystrokeMessageForAddon(Visio.MSGWrap msgWrap)
        {
        return false;
        }
    
      /// <summary>
      /// Avant la fermeture du document
      /// </summary>
      /// <param name="visPage"></param>
      public void OnBeforeDocumentClose(Visio.Document visDocument)
        {
        try
          {
          if (VisMeth.IsDocumentUserCellExist(visDocument, "VisualVisioSVGLightDocType", "StandardDoc")||
            VisMeth.IsDocumentUserCellExist(visDocument, "safeprojectnameDocType", "StandardDoc"))
            {
            // c'est un document VisioAddin
            addinApplication.InitializeMember(visDocument);
            }
          }
        catch
          {
          }
        }

      /// <summary>
      /// Avant la fermeture du document
      /// </summary>
      /// <param name="visPage"></param>
      public void OnDocumentChanged(Visio.Document visDocument)
        {
        }
    
      /// <summary>
      /// A la création d'un document
      /// </summary>
      /// <param name="visDocument"></param>
      public void OnDocumentCreated(Visio.Document visDocument)
        {
        }

      /// <summary>
      /// Ouverture d'un document
      /// </summary>
      /// <param name="visPage"></param>
      public void OnDocumentOpened(Visio.Document visDocument)
        {
        try
          {
          if (VisMeth.IsDocumentUserCellExist(visDocument, "VisualVisioSVGLightDocType", "StandardDoc")||
            VisMeth.IsDocumentUserCellExist(visDocument, "safeprojectnameDocType", "StandardDoc"))
            {
            // c'est un document VisioAddin
            addinApplication.InitializeMember(visDocument);
            }
          }
        catch
          {
          }
        }

      /// <summary>
      /// Avant la fermeture d'une fenêtre
      /// </summary>
      /// <param name="visWindow"></param>
      public void OnBeforeWindowClosed(Visio.Window visWindow)
        {
        string strCaption = visWindow.Caption;
        if (strCaption.Contains(StringEx.ParamString("idsReportTitle")))
          {
          addinApplication.DisplayReportClose();
          }
        }

      /// <summary>
      /// Activation d'une fenêtre
      /// </summary>
      /// <param name="visWindow"></param>
      public void OnWindowActivated(Visio.Window visWindow)
        {
        try
          {
          Visio.Document visDocument;

          visDocument = visWindow.Document;
          if (VisMeth.IsDocumentUserCellExist(visDocument, "VisualVisioSVGLightDocType", "StandardDoc"))
            {
            // c'est un document VisioAddin
            addinApplication.InitializeMember(visDocument);
            }
          }
        catch
          {
          }
        }

      /// <summary>
      /// Ajout d'une page
      /// </summary>
      /// <param name="visPage"></param>
      //[CLSCompliant(false)]
      public void OnPageAdded(Visio.Page visPage)
        {
        try
          {
          }
        catch
          {
          }
        }

      /// <summary>
      /// Suppression de page
      /// </summary>
      /// <param name="visPage"></param>
      //[CLSCompliant(false)]
      public void OnPageDeleted(Visio.Page visPage)
        {
        try
          {
          }
        catch
          {
          }
        }

      /// <summary>
      /// Changement de la page active
      /// </summary>
      /// <param name="visWindow"></param>
      //[CLSCompliant(false)]
      public void OnPageTurn(Visio.Window visWindow)
        {
        try
          {
          }
        catch
          {
          }
        }

      /// <summary>
      /// La page active a changé
      /// </summary>
      /// <param name="visWindow"></param>
      //[CLSCompliant(false)]
      public void OnPageTurned(Visio.Window visWindow)
        {

        try
          {
          }
        catch
          {
          }
        }

      /// <summary>
      /// Ajout d'une forme
      /// </summary>
      /// <param name="visShape"></param>
      //[CLSCompliant(false)]
      public void OnShapeAdded(Visio.Shape visShape)
        {
        try
          {
          }
        catch
          {
          StringEx.MessageBox("idsErrorAdd");
          }
        }

      /// <summary>
      /// Avant l'effacement d'une forme
      /// </summary>
      /// <param name="visShape"></param>
      //[CLSCompliant(false)]
      public void OnBeforeShapeDeleted(Visio.Shape visShape)
        {
        try
          {
          }
        catch
          {
          }
        }

      /// <summary>
      /// Aprés l'effacement d'une forme
      /// </summary>
      /// <param name="visSelection"></param>
      /// <param name="strMoreInfo"></param>
      //[CLSCompliant(false)]
      public void OnShapeDeleted(Visio.Selection visSelection, string strMoreInfo)
        {
        try
          {
          }
        catch (Exception exp)
          {
          throw exp;
          }
        }


      /// <summary>
      /// Demande d'autorisation d'effacement d'une forme
      /// </summary>
      /// <param name="visSelection"></param>
      /// <returns></returns>
      //[CLSCompliant(false)]
      public bool OnQueryCancelSelectionDelete(Visio.Selection visSelection)
        {
        bool bCancelDelete = false;

        try
          {
          }
        catch
          {
          }
        return bCancelDelete;
        }

      /// <summary>
      /// Traitement à faire en tâche de fond
      /// </summary>
      public void OnHandleNonePending()
        {
        }

      /// <summary>
      /// Aprés édition du texte d'une forme
      /// </summary>
      /// <param name="visioShape"></param>
      //[CLSCompliant(false)]
      public void OnShapeExitTextEdit(Visio.Shape visioShape)
        {
        try
          {
          }
        catch
          {
          }
        }

      /// <summary>
      /// Au changement de parent d'une shape
      /// </summary>
      /// <param name="visioShape"></param>
      //[CLSCompliant(false)]
      public void OnShapeParentChange(Visio.Shape visioShape)
        {

        try
          {
          }
        catch
          {
          }
        }

      /// <summary>
      /// Connexion d'un lien à une shape
      /// </summary>
      /// <param name="visConnect"></param>
      //[CLSCompliant(false)]
      public void OnConnectAdded(Visio.Connect visConnect)
        {

        try
          {
          }
        catch (Exception except)
          {
          object[] tabParams = {"OnConnectAdded :",
															 except.Message};

          StringEx.ParamMessageBox("strIdsErreurApp2", tabParams);
          }
        }

      /// <summary>
      /// Déconnexion d'un lien d'une shape
      /// </summary>
      /// <param name="visConnect"></param>
      //[CLSCompliant(false)]
      public void OnConnectDeleted(Visio.Connect visConnect)
        {

        try
          {
          }
        catch (Exception except)
          {
          object[] tabParams = { "OnConnectDeleted :", except.Message };

          StringEx.ParamMessageBox("strIdsErreurApp2", tabParams);
          }
        }

      /// <summary>
      /// Au changement du contenu d'une cellule
      /// </summary>
      /// <param name="visCell"></param>
      //[CLSCompliant(false)]
      public void OnCellChanged(Visio.Cell visCell)
        {
        try
          {
          }
        catch
          {
          }
        }

      /// <summary>
      /// Au changement de la formule d'une cellule
      /// </summary>
      /// <param name="visCell"></param>
      //[CLSCompliant(false)]
      public void OnFormulaChanged(Visio.Cell visCell)
        {

        try
          {
          }
        catch
          {
          }
        }

      //[CLSCompliant(false)]
      public void OnSelectionChanged(Visio.Window visWindow)
        {

        try
          {
          }
        catch
          {
          }
        }


      /// <summary>
      /// 
      /// </summary>
      /// <param name="x"></param>
      /// <param name="y"></param>
      public void OnMouseDown(double x, double y)
        {
        }



      /// <summary>
      /// Gestion du déplacement de la souris
      /// </summary>
      /// <param name="x"></param>
      /// <param name="y"></param>
      /// <param name="lKeyButtonstate"></param>
      public void OnMouseMove(double x, double y, long lKeyButtonstate)
        {

        }


      public void OnControlKeyPress(int iKeyButtonState, int iKeyAscii)
        {

        }

      //[CLSCompliant(false)]
      public void OnControlKeyDown(Visio.KeyboardEvent visKeyboardEvent)
        {

        }

    #region marker

      /// <summary>
      /// 
      /// </summary>
      /// <param name="visApp"></param>
      /// <param name="iEvent"></param>
      /// <param name="iSubEvent"></param>
      //[CLSCompliant(false)]
      public void OnMarkerConnectorPopupEvent(Visio.Application visApp,
                                              int iEvent, int iSubEvent, int iShapeID)
        {

        }

      /// <summary>
      /// Gestion de l'événement double clic sur un objet
      /// </summary>
      /// <param name="visApp"></param>
      /// <param name="iEvent"></param>
      /// <param name="iSubEvent"></param>
      //[CLSCompliant(false)]
      public void OnMarkerShapePopupEvent(Visio.Application visApp,
                                     int iEvent, int iSubEvent)
        {

        }

    #endregion
    #endregion

    #region menu

      public void Click(string strMenuToCall)
        {

        try
          {
          if (strMenuToCall == StringEx.GetString("idsMenuItemNewTag"))
            {
            addinApplication.NewFile();
            }
          else if (strMenuToCall == StringEx.GetString("idsMenuItemOptionsTag"))
            {
            addinApplication.Options();
            }
          else if (strMenuToCall == StringEx.GetString("idsMenuItemHelpTag"))
            {
            addinApplication.Help();
            }
          else if (strMenuToCall == StringEx.GetString("idsMenuItemAboutTag"))
            {
            addinApplication.About();
            }
          }
        catch
          {
          }
        }

    #endregion

#endif

    }
  }
