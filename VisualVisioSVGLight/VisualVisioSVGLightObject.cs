// VisualVisioSVGLightObject.cs
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
using System.Collections.Generic;
using System.Text;
using Visio = Microsoft.Office.Interop.Visio;

#if VISIOASM
using VisioAsm;
using VisMeth = VisioAsm.VLMethods;
using VisCst = VisioAsm.VLConstants;
#endif
#if STRINGASM
using StringAsm;
#endif

namespace VisualVisioSVGLight
  {
  class VisualVisioSVGLightObject
    {
    #region variables
    private string strName;
    private Visio.Shape visShape;

    #endregion
    #region propriétés

    public string Name
      {
      get
        {
        return strName;
        }
      set
        {
        strName = value;
        }
      }

    #endregion
    /// <summary>
    /// Classes contenant les méthodes de gestion des objets WVisioAddinBidon
    /// </summary>
    public VisualVisioSVGLightObject()
      {
      //
      // TODO : ajoutez ici la logique du constructeur
      //
      }

    public void Fill(string strNom)
      {

      Name = strNom;
      }

    public string GetMasterName()
      {
      string strMasterName = "";

      switch (strName)
        {
        case "CDISC":
          strMasterName = "CDISC";
          break;
        default:
          strMasterName = "";
          break;
        }
      return strMasterName;
      }

    public bool DropObject(Visio.Documents visDocuments, Visio.Document visStencil, string strStencilPath,
                           string strStencilName, Visio.Page visPage)
      {
      string strMasterName = "";

      try
        {
        strMasterName = GetMasterName();
        if (strMasterName != "")
          {
#if VISIOASM
          visShape = VisMeth.AddShapeToPage(visDocuments, visStencil, visPage, strStencilPath,
                                            strStencilName, strMasterName, false);
#endif
          UpdateAllCharacteristics();
          }
        }
      catch
        {
#if STRINGASM
        StringEx.ParamMessageBox("idsInvalidShape", strMasterName);
#endif
        }

      return true;
      }

    public bool UpdateAllCharacteristics()
      {
      object[,] arProp;
      ArrayList arFieldToIgnore = new ArrayList();

      arProp = new object[9, 4];
      // Mise à jour propriété personnalisée Nom
      // nom, valeur, type (Non utilisé dans ce cas), format de la propriété personnalisée (Non utilisé dans ce cas)
      arProp[0, 0] = "Prop." + "PNOM";
      arProp[0, 1] = Name;
      arProp[0, 2] = (short)Visio.VisCellVals.visPropTypeString;
      arProp[0, 3] = null;
#if VISIOASM
      VisMeth.SetPropFromArray(visShape, arProp, false);
#endif
      return true;
      }

    static public bool IsValidObjecType(string strObjName)
      {
      bool bValid = false;

      switch (strObjName)
        {
        case "CDISC":
          bValid = true;
          break;
        default:
          break;
        }
      return bValid;
      }

    }
  }
