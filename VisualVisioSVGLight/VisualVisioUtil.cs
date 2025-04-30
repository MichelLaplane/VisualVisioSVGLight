using System;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisualVisioSVGLight
  {
  internal class VisualVisioUtil
    {

      /// <summary>
    /// Récupération de la valeur double de la cellule SRC d'une Shape
    /// </summary>
    /// <param name="visShape"></param>
    /// <param name="srcValue"></param>
    /// <param name="dblValue"></param>
    /// <returns></returns>
    public static bool GetDoubleCellVal(Visio.Shape visShape, int srcValue, out double dblValue)
      {
      Visio.Cell visCell;

      // Section = visioSRCValue[srcValue,0]
      // Ligne = visioSRCValue[srcValue,1]
      // Cellule = visioSRCValue[srcValue,2]
      visCell = visShape.get_CellsSRC((short)VLConstants.visioSRCValue[srcValue, 0],
                                      (short)VLConstants.visioSRCValue[srcValue, 1],
                                      (short)VLConstants.visioSRCValue[srcValue, 2]);
      return GetDoubleCellVal(visCell, out dblValue);
      }

    /// <summary>
    /// Récupération de la valeur double d'une cellule
    /// </summary>
    /// <param name="visCell"></param>
    /// <param name="dblValue"></param>
    /// <returns></returns>
    public static bool GetDoubleCellVal(Visio.Cell visCell, out double dblValue)
      {

      dblValue = visCell.get_Result(Visio.VisUnitCodes.visNumber);
      return true;
      }

    /// <summary>
    /// Affectation d'une valeur double dans la cellule SRC d'une Shape
    /// sans unités
    /// </summary>
    /// <param name="visShape"></param>
    /// <param name="srcValue"></param>
    /// <param name="dblValue"></param>
    /// <returns></returns>
    public static bool SetDoubleCellVal(Visio.Shape visShape, int srcValue, double dblValue)
      {
      Visio.Cell visCell;

      try
        {
        // Section = visioSRCValue[srcValue,0]
        // Ligne = visioSRCValue[srcValue,1]
        // Cellule = visioSRCValue[srcValue,2]
        visCell = visShape.get_CellsSRC((short)VLConstants.visioSRCValue[srcValue, 0],
                                        (short)VLConstants.visioSRCValue[srcValue, 1],
                                        (short)VLConstants.visioSRCValue[srcValue, 2]);
        return SetDoubleCellVal(visCell, dblValue);
        }
      catch
        {
        return false;
        }
      }

    /// <summary>
    /// Affectation d'une valeur double en unité visUnits à une cellule
    /// </summary>
    /// <param name="visCell"></param>
    /// <param name="visUnits"></param>
    /// <param name="dblValue"></param>
    /// <returns></returns>
    public static bool SetDoubleCellVal(Visio.Cell visCell, int visUnits, double dblValue)
      {

      try
        {
        visCell.set_ResultForce(visUnits, dblValue);
        return true;
        }
      catch
        {
        return false;
        }
      }

    /// <summary>
    /// Affectation d'une formule string à une cellule
    /// </summary>
    /// <param name="visCell"></param>
    /// <param name="strFormula"></param>
    /// <returns></returns>
    public static bool SetFormulaCell(Visio.Cell visCell, string strFormula)
      {
      try
        {
        if (strFormula != null)
          {
          visCell.FormulaForceU = strFormula;
          return true;
          }
        }
      catch
        {
        return false;
        }
      return false;
      }

    /// <summary>
    /// Affectation d'une valeur double sans unité à une cellule
    /// </summary>
    /// <param name="visCell"></param>
    /// <param name="dblValue"></param>
    /// <returns></returns>
    public static bool SetDoubleCellVal(Visio.Cell visCell, double dblValue)
      {

      try
        {
        visCell.set_ResultForce((int)Visio.VisUnitCodes.visNumber, dblValue);
        return true;
        }
      catch
        {
        return false;
        }
      }

    /// <summary>
    /// Affectation d'une formule dans la cellule SRC d'une Shape
    /// </summary>
    /// <param name="visShape"></param>
    /// <param name="sectionIndex"></param>
    /// <param name="rowIndex"></param>
    /// <param name="cellIndex"></param>
    /// <param name="strFormula"></param>
    /// <returns></returns>
    public static bool SetFormulaCell(Visio.Shape visShape, int sectionIndex,
                                      int rowIndex, int cellIndex, string strFormula)
      {
      Visio.Cell visCell;

      try
        {
        // Section = visioSRCValue[srcValue,0]
        // Ligne = visioSRCValue[srcValue,1]
        // Cellule = visioSRCValue[srcValue,2]
        visCell = visShape.get_CellsSRC((short)sectionIndex,
                                        (short)rowIndex,
                                        (short)cellIndex);
        return SetFormulaCell(visCell, strFormula);
        }
      catch
        {
        }
      return false;
      }

    /// <summary>
    /// Affectation d'une valeur int sans unités à une cellule
    /// </summary>
    /// <param name="visCell"></param>
    /// <param name="iValue"></param>
    /// <returns></returns>
    public static bool SetIntCellVal(Visio.Cell visCell, int iValue)
      {

      try
        {
        visCell.set_ResultFromIntForce((int)Visio.VisUnitCodes.visNumber, iValue);
        }
      catch (System.Exception except)
        {
        string strMessage = except.Message;
        return false;
        }
      return true;
      }

    /// <summary>
    /// Affectation d'une valeur int dans la cellule SRC d'une Shape
    /// sans unité
    /// </summary>
    /// <param name="visShape"></param>
    /// <param name="srcValue"></param>
    /// <param name="iValue"></param>
    /// <returns></returns>
    public static bool SetIntCellVal(Visio.Shape visShape, int srcValue, int iValue)
      {
      Visio.Cell visCell;

      try
        {
        // Section = visioSRCValue[srcValue,0]
        // Ligne = visioSRCValue[srcValue,1]
        // Cellule = visioSRCValue[srcValue,2]
        visCell = visShape.get_CellsSRC((short)VLConstants.visioSRCValue[srcValue, 0],
                                        (short)VLConstants.visioSRCValue[srcValue, 1],
                                        (short)VLConstants.visioSRCValue[srcValue, 2]);
        return SetIntCellVal(visCell, iValue);
        }
      catch
        {
        return false;
        }
      }

    /// <summary>
    /// Affectation d'une valeur double dans la cellule Section, Row, Column d'une Shape
    /// sans unités
    /// </summary>
    /// <param name="visShape"></param>
    /// <param name="srcValue"></param>
    /// <param name="dblValue"></param>
    /// <returns></returns>
    public static bool SetDoubleCellVal(Visio.Shape visShape, int iSection, int iRow, int iColumn,
                                        double dblValue)
      {
      Visio.Cell visCell;

      try
        {
        visCell = visShape.get_CellsSRC((short)iSection, (short)iRow, (short)iColumn);
        return SetDoubleCellVal(visCell, dblValue);
        }
      catch
        {
        return false;
        }
      }


    /// <summary>
    /// Affectation d'une formule dans la cellule SRC d'une Shape
    /// </summary>
    /// <param name="visShape"></param>
    /// <param name="srcValue"></param>
    /// <param name="strFormula"></param>
    /// <returns></returns>
    public static bool SetFormulaCell(Visio.Shape visShape, int srcValue, string strFormula)
      {
      Visio.Cell visCell;

      try
        {
        // Section = visioSRCValue[srcValue,0]
        // Ligne = visioSRCValue[srcValue,1]
        // Cellule = visioSRCValue[srcValue,2]
        visCell = visShape.get_CellsSRC((short)VLConstants.visioSRCValue[srcValue, 0],
                                        (short)VLConstants.visioSRCValue[srcValue, 1],
                                        (short)VLConstants.visioSRCValue[srcValue, 2]);
        return SetFormulaCell(visCell, strFormula);
        }
      catch
        {
        }
      return false;
      }

    /// <summary>
    /// Affectation d'une valeur double dans la cellule SRC d'une Shape
    /// en unité visUnits
    /// </summary>
    /// <param name="visShape"></param>
    /// <param name="srcValue"></param>
    /// <param name="visUnits"></param>
    /// <param name="dblValue"></param>
    /// <returns></returns>
    public static bool SetDoubleCellVal(Visio.Shape visShape, int srcValue, int visUnits, double dblValue)
      {
      Visio.Cell visCell;

      try
        {
        // Section = visioSRCValue[srcValue,0]
        // Ligne = visioSRCValue[srcValue,1]
        // Cellule = visioSRCValue[srcValue,2]
        visCell = visShape.get_CellsSRC((short)VLConstants.visioSRCValue[srcValue, 0],
                                        (short)VLConstants.visioSRCValue[srcValue, 1],
                                        (short)VLConstants.visioSRCValue[srcValue, 2]);
        return SetDoubleCellVal(visCell, visUnits, dblValue);
        }
      catch
        {
        return false;
        }
      }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="visShape"></param>
    /// <param name="bFill"></param>
    /// <param name="bLine"></param>
    public static void SetGeometryVisibility(Visio.Shape visShape, bool bFill, bool bLine)
      {
      int nbGeometry;

      nbGeometry = visShape.GeometryCount;
      for (int i = 0; i < nbGeometry; i++)
        {
        SetFormulaCell(visShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + i),
                      (int)Visio.VisRowIndices.visRowFirst, (int)Visio.VisCellIndices.visCompNoFill, (!bFill).ToString());
        SetFormulaCell(visShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + i),
                      (int)Visio.VisRowIndices.visRowFirst, (int)Visio.VisCellIndices.visCompNoLine, (!bLine).ToString());
        }
      }

    /// <summary>
    /// Récupération de la formule d'une cellule
    /// </summary>
    /// <param name="visCell"></param>
    /// <param name="strFormula"></param>
    /// <returns></returns>
    public static bool GetFormulaUCell(Visio.Cell visCell, out string strFormula)
      {

      strFormula = visCell.FormulaU;
      return true;
      }

    /// <summary>
    /// Récupération de la formule de la cellule SRC d'une Shape
    /// </summary>
    /// <param name="visShape"></param>
    /// <param name="srcValue"></param>
    /// <param name="strFormula"></param>
    /// <returns></returns>
    public static bool GetFormulaUCell(Visio.Shape visShape, int srcValue, out string strFormula)
      {
      Visio.Cell visCell;

      // Section = visioSRCValue[srcValue,0]
      // Ligne = visioSRCValue[srcValue,1]
      // Cellule = visioSRCValue[srcValue,2]
      visCell = visShape.get_CellsSRC((short)VLConstants.visioSRCValue[srcValue, 0],
                                      (short)VLConstants.visioSRCValue[srcValue, 1],
                                      (short)VLConstants.visioSRCValue[srcValue, 2]);
      return GetFormulaUCell(visCell, out strFormula);
      }

    /// <summary>
    /// Récupération d'une shape identifiée par son nom
    /// la page visPage
    /// </summary>
    /// <param name="visPage"></param>
    /// <param name="strName"></param>
    /// <param name="visShape"></param>
    /// <returns></returns>
    public static bool GetVisShape(Visio.Page visPage, string strName, out Visio.Shape visShape)
      {
      Visio.Shapes visShapes;
      bool bFounded = false;

      visShapes = visPage.Shapes;
      visShape = null;
      try
        {
        foreach (Visio.Shape visCurShape in visShapes)
          {
          if (visCurShape.Name == strName)
            {
            bFounded = true;
            visShape = visCurShape;
            break;
            }
          }
        }
      catch (Exception except)
        {
        visShape = null;
        return bFounded;
        }
      return bFounded;
      }

    }




  /// <summary>
  /// Classe des constante de la librairie Visio
  /// </summary>
  public class VLConstants
    {
    /// <summary>
    /// 
    /// </summary>
    public const UInt32 VISCMD_DRCONNECTORTOOL = (UInt32)Visio.VisUICmds.visCmdDRConnectorTool;

    #region valeurs SRC

    /// <summary>
    /// Valeur permettant d'accéder aux éléments du tableau
    /// des valeurs SRC d'une cellule
    /// </summary>
    public enum SRCValue
      {
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_FILLFOREGNB = 0,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_FILLBACKGND = 1,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LINEWEIGHT = 2,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LINECOLOR = 3,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_WIDTH = 4,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_HEIGHT = 5,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PINX = 6,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PINY = 7,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LOCPINX = 8,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LOCPINY = 9,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PAGESHDWOFFSETX = 10,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PAGESHDWOFFSETY = 11,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CONNECTIONPOINTS = 12,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_1DBEGINX = 13,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_1DENDX = 14,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_1DBEGINY = 15,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_1DENDY = 16,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PAGEWIDTH = 17,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PAGEHEIGHT = 18,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PAGELINEADJUSTFROM = 19,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PAGELINEADJUSTTO = 20,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_FILLFOREGNBTRANS = 21,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_FILLBACKGNBTRANS = 22,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_FILLPATTERN = 23,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_SHDWFOREGND = 24,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_SHDWFOREGNDTRANS = 25,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_SHDWBACKGND = 26,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_SHDWBACKGNDTRANS = 27,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_SHDWPATTERN = 28,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_SHDWOFFSETX = 29,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_SHDWOFFSETY = 30,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_SHDWTYPE = 31,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_SHDWOBLANGLE = 32,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_SHDWSCALEFACT = 33,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LINEPATTERN = 34,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LINECAP = 35,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LINEBEGINARROW = 36,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LINEENDARROW = 37,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LINECOLORTRANS = 38,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LINEBEGINARROWSIZE = 39,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LINEENDARROWSIZE = 40,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_LINEROUNDING = 41,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_XGRIDSPACING = 42,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_YGRIDSPACING = 43,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_XGRIDDENSITY = 44,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_YGRIDDENSITY = 45,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_ENABLEGRID = 46,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_EVENTDROP = 47,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_IMAGEPROPERTYTRANSPARENCY = 48,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_HYPERLINKADDRESS = 49,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_HYPERLINKSUBADDRESS = 50,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PAGESCALE = 51,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PAGEDRAWINGSCALE = 52,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PAGEXRULERORIGIN = 53,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PAGEYRULERORIGIN = 54,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_ANGLE = 55,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_TXTWIDTH = 56,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_TXTHEIGHT = 57,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_TXTANGLE = 58,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_TXTPINX = 59,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_TXTPINY = 60,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_TXTLOCPINX = 61,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_TXTLOCPINY = 62,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CHARFONT = 63,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CHARSIZE = 64,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CHARCOLOR = 65,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CHARTRANSPARENCY = 66,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_TEXTFIELD = 67,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_HIDETEXT = 68,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_USERVALUE = 69,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CONTROLX = 70,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CONTROLY = 71,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CONTROLDINX = 72,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CONTROLDINY = 73,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CONTROLBEHAVX = 74,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CONTROLBEHAVY = 75,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CONTROLGLUE = 76,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_CONTROLTIP = 77,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTPAGELEFTMARGIN = 78,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTPAGERIGHTMARGIN = 79,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTPAGETOPMARGIN = 80,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTPAGEBOTTOMMARGIN = 81,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTSCALEX = 82,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTSCALEY = 83,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTPAGESX = 84,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTPAGESY = 85,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTCENTERX = 86,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTCENTERY = 87,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTONPAGE = 88,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTGRID = 89,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTORIENTATION = 90,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTPAPERKIND = 91,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_PRINTPAPERSOURCE = 92,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_GROUPSELECTMODE = 93,
      /// <summary>
      /// 
      /// </summary>
      ID_SRC_BIDON = 94
      }
    /// <summary>
    /// Tableau contenant les valeurs SRC des cellules
    /// </summary>
    public static int[,] visioSRCValue = {
													  // ID_SRC_FILLFOREGNB
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillForegnd},
													  // ID_SRC_FILLBACKGND
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillBkgnd},
													  // ID_SRC_LINEWEIGHT
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowLine,
                            (int)Visio.VisCellIndices.visLineWeight},
													  // ID_SRC_LINECOLOR
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowLine,
                            (int)Visio.VisCellIndices.visLineColor},
													  // ID_SRC_WIDTH
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowXFormOut,
                            (int)Visio.VisCellIndices.visXFormWidth},
													  // ID_SRC_HEIGHT
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowXFormOut,
                            (int)Visio.VisCellIndices.visXFormHeight},
													  // ID_SRC_PINX
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowXFormOut,
                            (int)Visio.VisCellIndices.visXFormPinX},
													  // ID_SRC_PINY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowXFormOut,
                            (int)Visio.VisCellIndices.visXFormPinY},
													  // ID_SRC_LOCPINX
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowXFormOut,
                            (int)Visio.VisCellIndices.visXFormLocPinX},
													  // ID_SRC_LOCPINY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowXFormOut,
                            (int)Visio.VisCellIndices.visXFormLocPinY},
													  // ID_SRC_PAGESHDWOFFSETX
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowPage,
                            (int)Visio.VisCellIndices.visPageShdwOffsetX},
													  // ID_SRC_PAGESHDWOFFSETY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowPage,
                            (int)Visio.VisCellIndices.visPageShdwOffsetY},
													  // ID_SRC_CONNECTIONPOINTS
													  {(int)Visio.VisSectionIndices.visSectionConnectionPts,
                            0,
                            0},
													  // ID_SRC_1DBEGINX
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowXForm1D,
                            (int)Visio.VisCellIndices.vis1DBeginX},
													  // ID_SRC_1DENDX
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowXForm1D,
                            (int)Visio.VisCellIndices.vis1DEndX},
													  // ID_SRC_1DBEGINY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowXForm1D,
                            (int)Visio.VisCellIndices.vis1DBeginY},
													  // ID_SRC_1DENDY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowXForm1D,
                            (int)Visio.VisCellIndices.vis1DEndY},
													  // ID_SRC_PAGEWIDTH
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowPage,
                            (int)Visio.VisCellIndices.visPageWidth},
													  // ID_SRC_PAGEHEIGHT
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowPage,
                            (int)Visio.VisCellIndices.visPageHeight},
													  // ID_SRC_PAGELINEADJUSTFROM
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowPageLayout,
                            (int)Visio.VisCellIndices.visPLOLineAdjustFrom},
													  // ID_SRC_PAGELINEADJUSTTO
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowPageLayout,
                            (int)Visio.VisCellIndices.visPLOLineAdjustTo},
													  // ID_SRC_FILLFOREGNBTRANS
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillForegndTrans},
													  // ID_SRC_FILLBACKGNBTRANS
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillBkgndTrans},
													  // ID_SRC_FILLPATTERN
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillPattern},
													  // ID_SRC_SHDWFOREGND
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillShdwForegnd},
													  // ID_SRC_SHDWFOREGNDTRANS
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillShdwForegndTrans},
													  // ID_SRC_SHDWBACKGND
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillShdwBkgnd},
													  // ID_SRC_SHDWBACKGNDTRANS
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillShdwBkgndTrans},
													  // ID_SRC_SHDWPATTERN
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillShdwPattern},
													  // ID_SRC_SHDWOFFSETX
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillShdwOffsetX},
													  // ID_SRC_SHDWOFFSETY
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillShdwOffsetY},
													  // ID_SRC_SHDWTYPE
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillShdwType},
													  // ID_SRC_SHDWOBLANGLE
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillShdwObliqueAngle},
													  // ID_SRC_SHDWSCALEFACT
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowFill,
                            (int)Visio.VisCellIndices.visFillShdwScaleFactor},
													  // ID_SRC_LINEPATTERN
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowLine,
                            (int)Visio.VisCellIndices.visLinePattern},
													  // ID_SRC_LINECAP
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowLine,
                            (int)Visio.VisCellIndices.visLineEndCap},
													  // ID_SRC_LINEBEGINARROW
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowLine,
                            (int)Visio.VisCellIndices.visLineBeginArrow},
													  // ID_SRC_LINEENDARROW
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowLine,
                            (int)Visio.VisCellIndices.visLineEndArrow},
													  // ID_SRC_LINECOLORTRANS
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowLine,
                            (int)Visio.VisCellIndices.visLineColorTrans},
													  // ID_SRC_LINEBEGINARROWSIZE
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowLine,
                            (int)Visio.VisCellIndices.visLineBeginArrowSize},
													  // ID_SRC_LINEENDARROWSIZE
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowLine,
                            (int)Visio.VisCellIndices.visLineEndArrowSize},
													  // ID_SRC_LINEROUNDING
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowLine,
                            (int)Visio.VisCellIndices.visLineRounding},
													  // ID_SRC_XGRIDSPACING
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowRulerGrid,
                            (int)Visio.VisCellIndices.visXGridSpacing},
													  // ID_SRC_YGRIDSPACING
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowRulerGrid,
                            (int)Visio.VisCellIndices.visYGridSpacing},
													  // ID_SRC_XGRIDDENSITY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowRulerGrid,
                            (int)Visio.VisCellIndices.visXGridDensity},
													  // ID_SRC_YGRIDDENSITY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowRulerGrid,
                            (int)Visio.VisCellIndices.visYGridDensity},
													  // ID_SRC_ENABLEGRID
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowPageLayout,
                            (int)Visio.VisCellIndices.visPLOEnableGrid},
													  // ID_SRC_EVENTDROP
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowEvent,
                            (int)Visio.VisCellIndices.visEvtCellDrop},
													  // ID_SRC_IMAGEPROPERTYTRANSPARENCY
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowImage,
                            (int)Visio.VisCellIndices.visImageTransparency},
													  // ID_SRC_HYPERLINKADDRESS
													  {(int)Visio.VisSectionIndices.visSectionHyperlink,
                             0,
                             (int)Visio.VisCellIndices.visHLinkAddress},
													  // ID_SRC_HYPERLINKSUBADDRESS
													  {(int)Visio.VisSectionIndices.visSectionHyperlink,
                             0,
                             (int)Visio.VisCellIndices.visHLinkSubAddress},
													  // ID_SRC_PAGESCALE
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowPage,
                            (int)Visio.VisCellIndices.visPageScale},
													  // ID_SRC_PAGEDRAWINGSCALE
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowPage,
                            (int)Visio.VisCellIndices.visPageDrawingScale},
													  // ID_SRC_PAGEXRULERORIGIN
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowRulerGrid,
                            (int)Visio.VisCellIndices.visXRulerOrigin},
													  // ID_SRC_PAGEYRULERORIGIN
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowRulerGrid,
                            (int)Visio.VisCellIndices.visYRulerOrigin},
													  // ID_SRC_ANGLE
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowXFormOut,
                            (int)Visio.VisCellIndices.visXFormAngle},
													  // ID_SRC_TXTWIDTH
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowTextXForm,
                            (int)Visio.VisCellIndices.visXFormWidth},
													  // ID_SRC_TXTHEIGHT
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowTextXForm,
                            (int)Visio.VisCellIndices.visXFormHeight},
													  // ID_SRC_TXTANGLE
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowTextXForm,
                            (int)Visio.VisCellIndices.visXFormAngle},
													  // ID_SRC_TXTPINX
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowTextXForm,
                            (int)Visio.VisCellIndices.visXFormPinX},
													  // ID_SRC_TXTPINY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowTextXForm,
                            (int)Visio.VisCellIndices.visXFormPinY},
													  // ID_SRC_TXTLOCPINX
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowTextXForm,
                            (int)Visio.VisCellIndices.visXFormLocPinX},
													  // ID_SRC_TXTLOCPINY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowTextXForm,
                            (int)Visio.VisCellIndices.visXFormLocPinX},
													  // ID_SRC_CHARFONT
													  {(int)Visio.VisSectionIndices.visSectionCharacter,
                             0,
                             (int)Visio.VisCellIndices.visCharacterFont},
													  // ID_SRC_CHARSIZE
													  {(int)Visio.VisSectionIndices.visSectionCharacter,
                             0,
                             (int)Visio.VisCellIndices.visCharacterSize},
													  // ID_SRC_CHARCOLOR
													  {(int)Visio.VisSectionIndices.visSectionCharacter,
                             0,
                             (int)Visio.VisCellIndices.visCharacterColor},
													  // ID_SRC_CHARTRANSPARENCY
													  {(int)Visio.VisSectionIndices.visSectionCharacter,
                             0,
                             (int)Visio.VisCellIndices.visCharacterColorTrans},
													  // ID_SRC_TEXTFIELD
														{(int)Visio.VisSectionIndices.visSectionTextField,
                            0,
                            (int)Visio.VisCellIndices.visFieldCell},
													  // ID_SRC_HIDETEXT
														{(int)Visio.VisSectionIndices.visSectionObject,
                            (int)Visio.VisRowIndices.visRowMisc,
                            (int)Visio.VisCellIndices.visHideText},
													  // ID_SRC_USERVALUE
														{(int)Visio.VisSectionIndices.visSectionUser,
                            0,
                            (int)Visio.VisCellIndices.visUserValue},
													  // ID_SRC_CONTROLX
														{(int)Visio.VisSectionIndices.visSectionControls,
                            0,
                            (int)Visio.VisCellIndices.visCtlX},
													  // ID_SRC_CONTROLY
														{(int)Visio.VisSectionIndices.visSectionControls,
                            0,
                            (int)Visio.VisCellIndices.visCtlY},
													  // ID_SRC_CONTROLDINX
														{(int)Visio.VisSectionIndices.visSectionControls,
                            0,
                            (int)Visio.VisCellIndices.visCtlXDyn},
													  // ID_SRC_CONTROLDINY
														{(int)Visio.VisSectionIndices.visSectionControls,
                            0,
                            (int)Visio.VisCellIndices.visCtlYDyn},
													  // ID_SRC_CONTROLBEHAVX
														{(int)Visio.VisSectionIndices.visSectionControls,
                            0,
                            (int)Visio.VisCellIndices.visCtlXCon},
													  // ID_SRC_CONTROLBEHAVY
														{(int)Visio.VisSectionIndices.visSectionControls,
                            0,
                            (int)Visio.VisCellIndices.visCtlYCon},
													  // ID_SRC_CONTROLGLUE
														{(int)Visio.VisSectionIndices.visSectionControls,
                            0,
                            (int)Visio.VisCellIndices.visCtlGlue},
													  // ID_SRC_CONTROLTIP
														{(int)Visio.VisSectionIndices.visSectionControls,
                            0,
                            (int)Visio.VisCellIndices.visCtlTip},
													  // ID_SRC_PRINTPAGELEFTMARGIN
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesLeftMargin},
                            // ID_SRC_PRINTPAGERIGHTMARGIN
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesRightMargin},
                            // ID_SRC_PRINTPAGETOPMARGIN
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesTopMargin},
                            // ID_SRC_PRINTPAGEBOTTOMMARGIN
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesBottomMargin},
                            // ID_SRC_PRINTSCALEX
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesScaleX},
                            // ID_SRC_PRINTSCALEY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesScaleY},
                            // ID_SRC_PRINTPAGESX
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesPagesX},
                            // ID_SRC_PRINTPAGESY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesPagesY},
                            // ID_SRC_PRINTCENTERX
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesCenterX},
                            // ID_SRC_PRINTCENTERY
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesCenterY},
                            // ID_SRC_PRINTONPAGE
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesOnPage},
                            // ID_SRC_PRINTGRID
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesPrintGrid},
                            // ID_SRC_PRINTORIENTATION
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesPageOrientation},
                            // ID_SRC_PRINTPAPERKIND
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesPaperKind},
                            // ID_SRC_PRINTPAPERSOURCE
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowPrintProperties,
                             (int)Visio.VisCellIndices.visPrintPropertiesPaperSource},
                            // ID_SRC_GROUPSELECTMODE
													  {(int)Visio.VisSectionIndices.visSectionObject,
                             (int)Visio.VisRowIndices.visRowGroup,
                             (int)Visio.VisCellIndices.visGroupSelectMode},
                             // ID_SRC_BIDON
													  {0,
                             0,
                             0}
                            };

    #endregion

    #region nom cellules

    /// <summary>
    /// Valeur permettant d'accéder aux éléments du tableau
    /// des cellules particulières cellName
    /// </summary>
    public enum CNValue
      {
      /// <summary>
      /// Propriétés personnalisées
      /// </summary>
      ID_CN_CUSTOMPROP = 0,
      /// <summary>
      /// Cellules personnalisées
      /// </summary>
      ID_CN_USERCELL = 1
      }

    /// <summary>
    /// Tableau contenant les noms des cellules particulières
    /// </summary>
    public static string[] cellName = {
                                     "Prop.",
                                     "User."
                                     };
    #endregion

    #region valeurs menus visio
    /// <summary>
    /// Valeurs des menus de Visio
    /// </summary>
    public enum menuValue
      {
      // Menu
      // Menu fichier
      ID_MEN_STDMENUBAR = 37,
      ID_MEN_FILEMENUCLOSE = 106,
      ID_MEN_FILEMENUSAVEAS = 748,
      ID_MEN_FILEMENU = 30002,
      ID_MEN_EDITMENU = 30003,
      ID_MEN_DISPLAYMENU = 30004,
      ID_MEN_INSERTMENU = 30005,
      ID_MEN_FORMATMENU = 30006,
      ID_MEN_TOOLSMENU = 30007,
      ID_MEN_WINDOWMENU = 30009,
      ID_MEN_HELPMENU = 30010,
      /// <summary>
      /// Menu Contextuel Drawing
      /// </summary>
      ID_MEN_STDMENUDRAWINGMOUSERIGHT = 32777,
      ID_MEN_STDMENUGABARITMASTER = 32782,
      // Disable menu Gabarit sauf ID_MEN_STDMENUGSCLOSE,
      // ID_MEN_STDMENUGSICONNAME,ID_MEN_STDMENUGSICON et ID_MEN_STDMENUGSNAME
      ID_MEN_STDMENUGABARITSYSTEM = 32789,
      ID_MEN_STDONGLETPAGE = 32815,
      ID_MEN_STDMENUDHELP = 34196,
      ID_MEN_SHAPEMENU = 50170,
      ID_MEN_DATAMENU = 50244
      }

    public enum itemValue
      {
      // Menu
      // Menu fichier
      ID_MEN_FILEMENUSAVE = 3,
      ID_MEN_STDMENUDMRCOPY = 19,
      ID_MEN_STDMENUDMRCUT = 21,
      ID_MEN_STDMENUDMRPASTE = 22,
      ID_MEN_FILEMENUOPEN = 23,
      ID_MEN_FILEMENUCLOSE = 106,
      ID_MEN_FILEMENUSAVEAS = 748,
      ID_MEN_FILEMENUNEW = 30037,
      // Menu edit
      ID_MEN_STDMENUPRINTPREVIEW = 32775,
      /// <summary>
      /// Menu Contextuel Drawing
      /// </summary>
      // Disable menu Gabarit sauf ID_MEN_STDMENUGSCLOSE,
      // ID_MEN_STDMENUGSICONNAME,ID_MEN_STDMENUGSICON et ID_MEN_STDMENUGSNAME
      ID_MEN_STDMENUGSCLOSE = 34220,
      ID_MEN_STDMENUGSICONNAME = 34248,
      ID_MEN_STDMENUGSICON = 34249,
      ID_MEN_STDMENUGSNAME = 34250,
      ID_MEN_STDMENUDMRREPOSITIONNELIB = 34464,
      ID_MEN_STDMENUDMRREPOSITIONNEBESOIN = 34465,
      ID_MEN_STDMENUDMRREPOSITIONNENEVER = 34466,
      ID_MEN_STDMENUDMRREPOSITIONNE = 34597,
      ID_MEN_STDMENUDMRREPOSITIONNECROISEMENT = 34605,
      ID_MEN_STDMENUGSICONDETAILS = 34660,
      // Menu contextuel onglet Page
      ID_MEN_STDONGLETPAGE_INSERT = 33846,
      ID_MEN_STDONGLETPAGE_DELETE = 34460,
      ID_MEN_STDONGLETPAGE_RENAME = 34461,
      ID_MEN_STDONGLETPAGE_REORG = 33848
      }

    /// <summary>
    /// Valeurs des barres d'outil de Visio
    /// </summary>
    public enum toolbarValue
      {
      // Barre d'outils standard
      ID_TOO_STDTOOLSAVE = 3,
      ID_TOO_STDTOOLBAR = 9,
      ID_TOO_STDTOOLNEW = 18,
      ID_TOO_STDTOOLOPEN = 23
      }
    #endregion


    /// <summary>
    /// 
    /// </summary>
    public VLConstants()
      {
      //
      // TODO : ajoutez ici la logique du constructeur
      //
      }
    }
  }
