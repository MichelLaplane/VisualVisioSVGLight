using Microsoft.Office.Interop.Visio;
using Svg;
using Svg.Pathing;
using Svg.Transforms;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisualVisioSVGLight
  {
  internal class VisualVisioSVGLightUtil
    {

    public static void CreateRect(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblTranslateX, double dblTranslateY, double dblAngle, double dblWidthRatio, double dblHeightRatio,
                        double dblSVGWidth, double dblSVGHeight, bool bViewBox, string strFill, string strTrokeColor)
      {
      string strParamStrokeColor = "", strParamLocStrokeColor = "";
      string strParamStrokeWidth = "", strParamLocStrokeWidth = "";
      string strParamFill = "", strParamLocFill = "";
      string strRounding = "";
      int iRed = 0, iGreen = 0, iBlue = 0;
      Visio.Shape visShape;

      strParamStrokeColor = strTrokeColor;
      strParamFill = strFill;
      double dblX1 = ((SvgRectangle)element).X + dblTranslateX;
      double dblY1 = ((SvgRectangle)element).Y + dblTranslateY;
      double dblX2 = dblX1 + ((SvgRectangle)element).Width;
      double dblY2 = dblY1 + ((SvgRectangle)element).Height;
      ((SvgRectangle)element).TryGetAttribute("rx", out strRounding);
      ((SvgRectangle)element).TryGetAttribute("stroke", out strParamLocStrokeColor);
      ((SvgRectangle)element).TryGetAttribute("stroke-width", out strParamLocStrokeWidth);
      ((SvgRectangle)element).TryGetAttribute("fill", out strParamLocFill);
      if (strParamLocStrokeColor != null)
        strParamStrokeColor = strParamLocStrokeColor;
      if (strParamLocStrokeWidth != null)
        strParamStrokeWidth = strParamLocStrokeWidth;
      if (strParamLocFill != null)
        strParamFill = strParamLocFill;
      if (strTrokeColor != "")
        {
        strParamStrokeColor = strTrokeColor;
        if ((strParamLocStrokeColor != "") && (strParamLocStrokeColor != null))
          {
          strParamStrokeColor = strParamLocStrokeColor;
          }
        }
      if (strFill != "")
        {
        strParamFill = strFill;
        if ((strParamLocFill != "") && (strParamLocFill != null))
          {
          strParamFill = strParamLocFill;
          }
        }
      if (bViewBox)
        {
        dblX1 = visPage.Application.ConvertResult(dblX1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY1 = -visPage.Application.ConvertResult(dblY1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblX2 = visPage.Application.ConvertResult(dblX2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY2 = -visPage.Application.ConvertResult(dblY2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        visShape = visPage.DrawRectangle(dblX1 / dblWidthRatio, dblY1 / dblHeightRatio, dblX2 / dblWidthRatio, dblY2 / dblHeightRatio);
        // Rotation éventuelle
        if (dblAngle != 0)
          {
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
          // Centre de rotation à gauche au centre
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
          // Centre de rotation en haut à gauche
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
          // Rotation de la forme. Attention le signe de l'angle doit être inversé
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
          // Centre de rotation au centre en haut pour commencer à revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
          // Centre de rotation au centre au centre pour revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
          //VLMethods.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out dblPinXAfterRepos);
          //VLMethods.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out dblPinYAfterRepos);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          }
        // Recentrage par rapport a la shape SVG
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
        }
      else
        {
        dblX1 = visPage.Application.ConvertResult(dblX1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY1 = visPage.Application.ConvertResult(dblY1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblX2 = visPage.Application.ConvertResult(dblX2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY2 = visPage.Application.ConvertResult(dblY2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        // Y coordinate is in the inverse direction of SVG so Y must be negative
        visShape = visPage.DrawRectangle(dblX1 / dblWidthRatio, -dblY1 / dblHeightRatio, dblX2 / dblWidthRatio, -dblY2 / dblHeightRatio);
        //visShape = visPage.DrawRectangle(dblX1, dblY1, dblX2, dblY2);
        // Rotation éventuelle
        if (dblAngle != 0)
          {
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
          // Centre de rotation à gauche au centre
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
          // Centre de rotation en haut à gauche
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
          // Rotation de la forme. Attention le signe de l'angle doit être inversé
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
          // Centre de rotation au centre en haut pour commencer à revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
          // Centre de rotation au centre au centre pour revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
          //VLMethods.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out dblPinXAfterRepos);
          //VLMethods.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out dblPinYAfterRepos);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          }
        // Recentrage par rapport a la shape SVG
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblPinXValue + dblSVGPinXValue) - (dblSVGWidth * 0.5));
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblPinYValue + dblSVGPinYValue) + (dblSVGHeight * 0.5));
        }
      ApplyShapeStyles(visPage, visShape, strParamStrokeWidth, strRounding, strParamStrokeColor, strParamFill, dblWidthRatio);
      }

    public static void CreateCircle(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblTranslateX, double dblTranslateY, double dblAngle, double dblWidthRatio, double dblHeightRatio,
                          double dblSVGWidth, double dblSVGHeight, bool bViewBox, string strFill, string strTrokeColor)
      {
      string strParamStrokeColor = "", strParamLocStrokeColor = "";
      string strParamStrokeWidth = "", strParamLocStrokeWidth = "";
      string strParamFill = "", strParamLocFill = "";
      string strRounding = "";
      int iRed = 0, iGreen = 0, iBlue = 0;
      Visio.Shape visShape;

      double dblCenterX = ((SvgCircle)element).CenterX + dblTranslateX;
      double dblCenterY = ((SvgCircle)element).CenterY + dblTranslateY;
      double dblRadius = ((SvgCircle)element).Radius;
      double dblDiameter = dblRadius * 2.0F;
      double dblX1 = dblCenterX - (dblDiameter * 0.5);
      double dblY1 = dblCenterY - (dblDiameter * 0.5);
      double dblX2 = dblX1 + dblDiameter;
      double dblY2 = dblY1 + dblDiameter;

      ((SvgCircle)element).TryGetAttribute("stroke", out strParamLocStrokeColor);
      ((SvgCircle)element).TryGetAttribute("stroke-width", out strParamLocStrokeWidth);
      ((SvgCircle)element).TryGetAttribute("fill", out strParamLocFill);
      if (strParamLocStrokeColor != null)
        strParamStrokeColor = strParamLocStrokeColor;
      if (strParamLocStrokeWidth != null)
        strParamStrokeWidth = strParamLocStrokeWidth;
      if (strParamLocFill != null)
        strParamFill = strParamLocFill;
      if (strTrokeColor != "")
        {
        strParamStrokeColor = strTrokeColor;
        if ((strParamLocStrokeColor != "") && (strParamLocStrokeColor != null))
          {
          strParamStrokeColor = strParamLocStrokeColor;
          }
        }
      if (strFill != "")
        {
        strParamFill = strFill;
        if ((strParamLocFill != "") && (strParamLocFill != null))
          {
          strParamFill = strParamLocFill;
          }
        }
      if (bViewBox)
        {
        dblX1 = visPage.Application.ConvertResult(dblX1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY1 = -visPage.Application.ConvertResult(dblY1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblX2 = visPage.Application.ConvertResult(dblX2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY2 = -visPage.Application.ConvertResult(dblY2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        visShape = visPage.DrawOval(dblX1 / dblWidthRatio, dblY1 / dblHeightRatio, dblX2 / dblWidthRatio, dblY2 / dblHeightRatio);
        // Rotation éventuelle
        if (dblAngle != 0)
          {
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
          // Centre de rotation à gauche au centre
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
          // Centre de rotation en haut à gauche
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
          // Rotation de la forme. Attention le signe de l'angle doit être inversé
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
          // Centre de rotation au centre en haut pour commencer à revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
          // Centre de rotation au centre au centre pour revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          }
        // Recentrage par rapport a la shape SVG
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
        }
      else
        {
        dblX1 = visPage.Application.ConvertResult(dblX1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY1 = visPage.Application.ConvertResult(dblY1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblX2 = visPage.Application.ConvertResult(dblX2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY2 = visPage.Application.ConvertResult(dblY2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        // Y coordinate is in the inverse direction of SVG so Y must be negative
        visShape = visPage.DrawRectangle(dblX1 / dblWidthRatio, -dblY1 / dblHeightRatio, dblX2 / dblWidthRatio, -dblY2 / dblHeightRatio);
        //visShape = visPage.DrawRectangle(dblX1, dblY1, dblX2, dblY2);
        // Rotation éventuelle
        if (dblAngle != 0)
          {
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
          // Centre de rotation à gauche au centre
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
          // Centre de rotation en haut à gauche
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
          // Rotation de la forme. Attention le signe de l'angle doit être inversé
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
          // Centre de rotation au centre en haut pour commencer à revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
          // Centre de rotation au centre au centre pour revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
          //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out dblPinXAfterRepos);
          //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out dblPinYAfterRepos);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          }
        // Recentrage par rapport a la shape SVG
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblPinXValue + dblSVGPinXValue) - (dblSVGWidth * 0.5));
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblPinYValue + dblSVGPinYValue) + (dblSVGHeight * 0.5));
        }
      ApplyShapeStyles(visPage, visShape, strParamStrokeWidth, strRounding, strParamStrokeColor, strParamFill, dblWidthRatio);
      }

    public static void CreatePolyline(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblTranslateX, double dblTranslateY, double dblAngle, double dblWidthRatio, double dblHeightRatio,
                          double dblSVGWidth, double dblSVGHeight, bool bViewBox, string strFill, string strTrokeColor)
      {
      string strParamStrokeColor = "", strParamLocStrokeColor = "";
      string strParamStrokeWidth = "", strParamLocStrokeWidth = "";
      string strParamFill = "", strParamLocFill = "";
      Visio.Shape visShape;

      strParamStrokeColor = strTrokeColor;
      strParamFill = strFill;
      ((SvgPolyline)element).TryGetAttribute("stroke", out strParamLocStrokeColor);
      ((SvgPolyline)element).TryGetAttribute("stroke-width", out strParamLocStrokeWidth);
      ((SvgPolyline)element).TryGetAttribute("fill", out strParamLocFill);
      if (strParamLocStrokeColor != null)
        strParamStrokeColor = strParamLocStrokeColor;
      if (strParamLocStrokeWidth != null)
        strParamStrokeWidth = strParamLocStrokeWidth;
      if (strParamLocFill != null)
        strParamFill = strParamLocFill;
      if (strTrokeColor != "")
        {
        strParamStrokeColor = strTrokeColor;
        if ((strParamLocStrokeColor != "") && (strParamLocStrokeColor != null))
          {
          strParamStrokeColor = strParamLocStrokeColor;
          }
        }
      if (strFill != "")
        {
        strParamFill = strFill;
        if ((strParamLocFill != "") && (strParamLocFill != null))
          {
          strParamFill = strParamLocFill;
          }
        }
      int nbPoints = ((SvgPolyline)element).Points.Count;
      double[] arPoint = new double[nbPoints];
      for (int i = 0; i < nbPoints; i++)
        {
        if (i % 2 == 0)
          {
          arPoint[i] = visPage.Application.ConvertResult(((SvgPolyline)element).Points[i].Value + dblTranslateX,
                                                         (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches)
                                                         / dblWidthRatio;
          }
        else
          {
          arPoint[i] = visPage.Application.ConvertResult(-((SvgPolyline)element).Points[i].Value + dblTranslateY,
                                                         (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches)
                                                         / dblHeightRatio;
          }
        }
      if (bViewBox)
        {
        visShape = visPage.DrawPolyline(arPoint, (int)Visio.VisDrawSplineFlags.visPolyline1D);
        // Rotation éventuelle
        if (dblAngle != 0)
          {
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
          // Centre de rotation à gauche au centre
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
          // Centre de rotation en haut à gauche
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
          // Rotation de la forme. Attention le signe de l'angle doit être inversé
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
          // Centre de rotation au centre en haut pour commencer à revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
          // Centre de rotation au centre au centre pour revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
          //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out dblPinXAfterRepos);
          //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out dblPinYAfterRepos);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          }
        // Recentrage par rapport a la shape SVG
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
        }
      else
        {
        visShape = visPage.DrawPolyline(arPoint, (int)Visio.VisDrawSplineFlags.visPolyline1D);
        // Rotation éventuelle
        if (dblAngle != 0)
          {
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
          // Centre de rotation à gauche au centre
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
          // Centre de rotation en haut à gauche
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
          // Rotation de la forme. Attention le signe de l'angle doit être inversé
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
          // Centre de rotation au centre en haut pour commencer à revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
          // Centre de rotation au centre au centre pour revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
          //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out dblPinXAfterRepos);
          //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out dblPinYAfterRepos);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          }
        // Recentrage par rapport a la shape SVG
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblPinXValue + dblSVGPinXValue) - (dblSVGWidth * 0.5));
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblPinYValue + dblSVGPinYValue) + (dblSVGHeight * 0.5));
        }
      ApplyShapeStyles(visPage, visShape, strParamStrokeWidth, "", strParamStrokeColor, strParamFill, dblWidthRatio);
      }


    public static void Create2DPolylineFromPath(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblWidthRatio, double dblHeightRatio,
                                  double dblSVGWidth, double dblSVGHeight)
      {
      double dblBeginX = 0.0, dblBeginY = 0.0;
      double dblEndX = 0.0, dblEndY = 0.0;
      double dblOriginX = 0.0, dblOriginY = 0.0, dblExtremityX = 0.0, dblExtremityY = 0.0;
      string strStrokeColor, strStrokeWidth, strFill;

      ((SvgPath)element).TryGetAttribute("stroke", out strStrokeColor);
      ((SvgPath)element).TryGetAttribute("stroke-width", out strStrokeWidth);
      ((SvgPath)element).TryGetAttribute("fill", out strFill);
      SvgPath svgPath = ((SvgPath)element);
      Svg.Pathing.SvgPathSegmentList arData = svgPath.PathData;
      int iGeometry = 0;
      int iGeometryLine = 0;
      dblOriginX = arData[0].End.X;
      dblOriginY = arData[0].End.Y;
      if (arData[arData.Count - 1].ToString().Substring(0, 1) == "Z" || arData[arData.Count - 1].ToString().Substring(0, 1) == "z")
        {
        dblExtremityX = arData[arData.Count - 2].End.X;
        dblExtremityY = arData[arData.Count - 2].End.Y;
        }
      else
        {
        dblExtremityX = arData[arData.Count - 1].End.X;
        dblExtremityY = arData[arData.Count - 1].End.Y;
        }
      dblOriginX = visPage.Application.ConvertResult(dblOriginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblOriginY = -visPage.Application.ConvertResult(dblOriginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblExtremityX = visPage.Application.ConvertResult(dblExtremityX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblExtremityY = -visPage.Application.ConvertResult(dblExtremityY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      Visio.Shape visPathShape = visPage.DrawRectangle(dblOriginX / dblWidthRatio, dblOriginY / dblHeightRatio, dblExtremityX / dblWidthRatio, dblOriginY / dblHeightRatio);
      // Deleting of all LineTo
      int nbGeometry = visPathShape.GeometryCount;
      // Get the count of rows in the current Geometry section. 
      int nbRows = visPathShape.RowCount[(short)Visio.VisSectionIndices.visSectionFirstComponent];

      for (int iRow = 1; iRow < nbRows - 1; iRow++)
        {
        visPathShape.DeleteRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (short)(Visio.VisRowIndices.visRowVertex + 1));
        }
      //Visio.Shape visPathShape = visPage.DrawLine(dblOriginX / dblWidthRatio, dblOriginY / dblHeightRatio, dblExtremityX / dblWidthRatio, dblOriginY / dblHeightRatio);
      foreach (Svg.Pathing.SvgPathSegment pathSegment in arData)
        {
        string strPoint = pathSegment.ToString();
        switch (strPoint.Substring(0, 1))
          {
          case "M":
          case "m":
            // MoveTo
            if (iGeometryLine > 0)
              {
              // Create a new geometry line
              visPathShape.AddRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (int)Visio.VisRowIndices.visRowLast, (int)Visio.VisRowTags.visTagMoveTo);
              }
            dblBeginX = pathSegment.End.X;
            dblBeginY = pathSegment.End.Y;
            dblBeginX = visPage.Application.ConvertResult(dblBeginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblBeginY = visPage.Application.ConvertResult(dblBeginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            iGeometryLine++;
            break;
          case "L":
          case "l":
            // LineTo
            if (iGeometryLine > 0)
              {
              // Create a new geometry line
              visPathShape.AddRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (int)Visio.VisRowIndices.visRowLast, (int)Visio.VisRowTags.visTagLineTo);
              }
            dblBeginX = pathSegment.End.X;
            dblBeginY = pathSegment.End.Y;
            dblBeginX = visPage.Application.ConvertResult(dblBeginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblBeginY = -visPage.Application.ConvertResult(dblBeginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                          (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visX, (dblBeginX - dblOriginX) / dblWidthRatio);
            VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                          (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visY, (dblBeginY - dblOriginY) / dblHeightRatio);
            iGeometryLine++;
            break;
          case "C":
            // CubicBezierTo
            double dblStartX = dblBeginX;
            double dblStartY = dblBeginY;
            double dblFirstControlPointX = ((SvgCubicCurveSegment)pathSegment).FirstControlPoint.X;
            double dblFirstControlPointY = ((SvgCubicCurveSegment)pathSegment).FirstControlPoint.Y;
            double dblSecondControlPointX = ((SvgCubicCurveSegment)pathSegment).SecondControlPoint.X;
            double dblSecondControlPointY = ((SvgCubicCurveSegment)pathSegment).SecondControlPoint.Y;
            double dblEndPointX = ((PointF)pathSegment.End).X;
            double dblEndPointY = ((PointF)pathSegment.End).Y;
            dblFirstControlPointX = visPage.Application.ConvertResult(dblFirstControlPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblFirstControlPointY = -visPage.Application.ConvertResult(dblFirstControlPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblSecondControlPointX = visPage.Application.ConvertResult(dblSecondControlPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblSecondControlPointY = -visPage.Application.ConvertResult(dblSecondControlPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblEndPointX = visPage.Application.ConvertResult(dblEndPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblEndPointY = visPage.Application.ConvertResult(dblEndPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            double[] arControlPoint = new double[6];
            arControlPoint[0] = (dblFirstControlPointX) / dblWidthRatio;
            arControlPoint[1] = (dblFirstControlPointY + dblOriginY + (dblBeginY * 0.5)) / dblHeightRatio;
            arControlPoint[2] = (dblEndPointX) / dblWidthRatio;
            arControlPoint[3] = (dblEndPointY + dblOriginY) / dblHeightRatio;
            arControlPoint[4] = (dblSecondControlPointX) / dblWidthRatio;
            arControlPoint[5] = (dblSecondControlPointY + dblOriginY + (dblBeginY * 0.5)) / dblHeightRatio;
            //ConvertBezierToNurbs(visPage, arControlPoint);
            Visio.Shape visShapeBezier = visPage.DrawBezier(arControlPoint, 2, (int)VisDrawSplineFlags.visSpline1D);
            visShapeBezier.UpdateAlignmentBox();
            visShapeBezier.BoundingBox((int)VisBoundingBoxArgs.visBBoxExtents, out double dblLeft1, out double dblBottom1, out double dblRight1, out double dblTop1);
            // Recentrage par rapport a la shape SVG
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue1);
            VisualVisioUtil.GetDoubleCellVal(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue1);
            VisualVisioUtil.SetDoubleCellVal(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblPinXValue1 + dblSVGPinXValue1) - (dblSVGWidth * 0.5));
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue1);
            VisualVisioUtil.GetDoubleCellVal(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue1);
            VisualVisioUtil.SetDoubleCellVal(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblPinYValue1 + dblSVGPinYValue1) + (dblSVGHeight * 0.5));
            if (strStrokeWidth != "")
              {
              Double.TryParse(strStrokeWidth, out double dblLineWeight);
              dblLineWeight = dblLineWeight / dblWidthRatio;
              NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
              VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_LINEWEIGHT, dblLineWeight.ToString("0.00 pt", nfi));
              }
            switch (strStrokeColor)
              {
              case "Blue":
                VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, "RGB(0,112,192)");
                break;
              case "Green":
                VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, "RGB(0,176,80)");
                break;
              case "Purple":
                VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, "RGB(112,48,160)");
                break;
              case "Red":
                VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, "RGB(255,0,0)");
                break;
              case "none":
                VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_LINEPATTERN, "0");
                break;
              case null:
                break;
              default:
                if (strStrokeColor.StartsWith("#") && strStrokeColor.Length == 7)
                  {
                  int r = int.Parse(strStrokeColor.Substring(1, 2), System.Globalization.NumberStyles.HexNumber);
                  int g = int.Parse(strStrokeColor.Substring(3, 2), System.Globalization.NumberStyles.HexNumber);
                  int b = int.Parse(strStrokeColor.Substring(5, 2), System.Globalization.NumberStyles.HexNumber);
                  VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, $"RGB({r},{g},{b})");
                  }
                break;
              }
            switch (strFill)
              {
              case "Blue":
                VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(0,112,192)");
                break;
              case "Green":
                VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(0,176,80)");
                break;
              case "Purple":
                VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(112,48,160)");
                break;
              case "Red":
                VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(255,0,0)");
                break;
              case "Yellow":
                VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(255,255,0)");
                break;
              case "None":
                VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_FILLPATTERN, "0");
                break;
              case null:
                break;
              default:
                if (strFill.StartsWith("#") && strFill.Length == 7)
                  {
                  int r = int.Parse(strFill.Substring(1, 2), System.Globalization.NumberStyles.HexNumber);
                  int g = int.Parse(strFill.Substring(3, 2), System.Globalization.NumberStyles.HexNumber);
                  int b = int.Parse(strFill.Substring(5, 2), System.Globalization.NumberStyles.HexNumber);
                  VisualVisioUtil.SetFormulaCell(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, $"RGB({r},{g},{b})");
                  }
                break;
              }
            break;
          case "c":
            break;
          case "Z":
          case "z":
            // Fermeture de la forme
            // Create a new geometry line
            visPathShape.AddRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (int)Visio.VisRowIndices.visRowLast,
                                (int)Visio.VisRowTags.visTagLineTo);
            VisualVisioUtil.SetFormulaCell(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                          (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visX, "Geometry" + (iGeometry + 1).ToString() + ".X1");
            VisualVisioUtil.SetFormulaCell(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                          (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visY, "Geometry" + (iGeometry + 1).ToString() + ".Y1");
            VisualVisioUtil.SetFormulaCell(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                          (int)Visio.VisRowIndices.visRowFirst, (int)Visio.VisCellIndices.visCompNoFill, false.ToString());
            break;
          }
        }
      visPathShape.UpdateAlignmentBox();
      visPathShape.BoundingBox((int)VisBoundingBoxArgs.visBBoxExtents, out double dblLeft, out double dblBottom, out double dblRight, out double dblTop);
      // Recentrage par rapport a la shape SVG
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
      VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblPinXValue + dblSVGPinXValue) - (dblSVGWidth * 0.5));
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      VisualVisioUtil.GetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
      VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblPinYValue + dblSVGPinYValue) + (dblSVGHeight * 0.5));
      if (strStrokeWidth != "")
        {
        Double.TryParse(strStrokeWidth, out double dblLineWeight);
        dblLineWeight = dblLineWeight / dblWidthRatio;
        NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
        VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_LINEWEIGHT, dblLineWeight.ToString("0.00 pt", nfi));
        }
      switch (strStrokeColor)
        {
        case "Blue":
          VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, "RGB(0,112,192)");
          break;
        case "Green":
          VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, "RGB(0,176,80)");
          break;
        case "Purple":
          VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, "RGB(112,48,160)");
          break;
        case "Red":
          VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, "RGB(255,0,0)");
          break;
        case "none":
          VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_LINEPATTERN, "0");
          break;
        case null:
          break;
        default:
          if (strStrokeColor.StartsWith("#") && strStrokeColor.Length == 7)
            {
            int r = int.Parse(strStrokeColor.Substring(1, 2), System.Globalization.NumberStyles.HexNumber);
            int g = int.Parse(strStrokeColor.Substring(3, 2), System.Globalization.NumberStyles.HexNumber);
            int b = int.Parse(strStrokeColor.Substring(5, 2), System.Globalization.NumberStyles.HexNumber);
            VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, $"RGB({r},{g},{b})");
            }
          break;
        }
      switch (strFill)
        {
        case "Blue":
          VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(0,112,192)");
          break;
        case "Green":
          VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(0,176,80)");
          break;
        case "Purple":
          VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(112,48,160)");
          break;
        case "Red":
          VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(255,0,0)");
          break;
        case "Yellow":
          VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(255,255,0)");
          break;
        case "None":
          VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_FILLPATTERN, "0");
          break;
        case null:
          break;
        default:
          if (strFill.StartsWith("#") && strFill.Length == 7)
            {
            int r = int.Parse(strFill.Substring(1, 2), System.Globalization.NumberStyles.HexNumber);
            int g = int.Parse(strFill.Substring(3, 2), System.Globalization.NumberStyles.HexNumber);
            int b = int.Parse(strFill.Substring(5, 2), System.Globalization.NumberStyles.HexNumber);
            VisualVisioUtil.SetFormulaCell(visPathShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, $"RGB({r},{g},{b})");
            }
          break;
        }
      }

    public static void CreateLine(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblWidthRatio, double dblHeightRatio,
                                  double dblSVGWidth, double dblSVGHeight)
      {
      int iVisInUnit, iVisOutUnit;

      //GetDoubleCellVal(visPage, (int)VLConstants.SRCValue.ID_SRC_WIDTH, (int)Visio.VisUnitCodes.visInches, out double dblPageWidth);
      //GetDoubleCellVal(visPage, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, (int)Visio.VisUnitCodes.visInches, out double dblPageHeight);
      double dblSubBeginX = (((SvgLine)element).StartX).Value;
      double dblSubBeginY = (((SvgLine)element).StartY).Value;
      //double dblConvertedSubBeginX = visApp.ConvertResult(dblSubBeginX, strSvgUnit, "pt");
      //double dblSubBeginY = dblSVGHeight - (((SvgLine)element).StartY).Value;
      double dblSubEndX = (((SvgLine)element).EndX).Value;
      double dblSubEndY = (((SvgLine)element).EndY).Value;
      dblSubBeginX = visPage.Application.ConvertResult(dblSubBeginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblSubBeginY = -visPage.Application.ConvertResult(dblSubBeginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblSubEndX = visPage.Application.ConvertResult(dblSubEndX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblSubEndY = -visPage.Application.ConvertResult(dblSubEndY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      Visio.Shape visShape = visPage.DrawLine(dblSubBeginX / dblWidthRatio, dblSubBeginY / dblHeightRatio, dblSubEndX / dblWidthRatio, dblSubEndY / dblHeightRatio);
      // Recentrage par rapport a la shape SVG
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      // Recadrage X
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DBEGINX, out double dblBeginXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DBEGINX, dblBeginXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DENDX, out double dblEndXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DENDX, dblEndXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
      // Recadrage Y
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DBEGINY, out double dblBeginYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DBEGINY, dblBeginYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DENDY, out double dblEndYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DENDY, dblEndYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
      }

    public static void CreateText(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblWidthRatio, double dblHeightRatio,
                                  double dblSVGWidth, double dblSVGHeight, string strSvgUnit, string strOutUnit)
      {

      string strText = ((SvgText)element).Text;
      double dblFontSize = ((SvgUnit)((SvgText)element).FontSize).Value;
      string strFontFamily = ((SvgText)element).FontFamily;

      double dblPinX = ((SvgUnitCollection)(((SvgText)element).X))[0].Value;
      double dblPinY = -((SvgUnitCollection)(((SvgText)element).Y))[0].Value;
      dblPinX = visPage.Application.ConvertResult(dblPinX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      double dblInchesFontSize = visPage.Application.ConvertResult(dblFontSize, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      // We increment the Y with the dblFontSize
      dblPinY = visPage.Application.ConvertResult(dblPinY + dblFontSize, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      // Y coordinate is in the inverse direction of SVG so Y must be negative
      //Visio.Shape visShape = visPage.DrawRectangle(dblPinX / dblWidthRatio, -dblPinY / dblHeightRatio, (dblPinX + (strText.Length * 10)) / dblWidthRatio, -dblPinY / dblHeightRatio);
      Visio.Shape visShape = visPage.DrawRectangle(dblPinX / dblWidthRatio, -dblPinY / dblHeightRatio, dblPinX / dblWidthRatio, -dblPinY / dblHeightRatio);
      var iSize = visShape.get_CellsSRC((int)Visio.VisSectionIndices.visSectionCharacter, 0, (int)Visio.VisCellIndices.visCharacterSize).ResultIU;
      //ID_SRC_CHARSIZE
      visShape.Text = strText;
      double dblTxtWidth = (strText.Length - 1) * iSize;
      double dblTxtHeight = iSize * 2;
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, (int)Visio.VisUnitCodes.visInches, dblTxtWidth);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, (int)Visio.VisUnitCodes.visInches, dblTxtHeight);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, (int)Visio.VisUnitCodes.visInches, (dblPinX / dblWidthRatio) + (dblTxtWidth * 0.5));
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, (int)Visio.VisUnitCodes.visInches, (dblPinY / dblHeightRatio) - (dblTxtHeight * 0.5));
      VisualVisioUtil.SetIntCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LINEPATTERN, 0);
      VisualVisioUtil.SetIntCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLPATTERN, 0);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_CHARSIZE, (int)Visio.VisUnitCodes.visPoints, dblFontSize);
      Visio.Fonts visFonts = visPage.Document.Fonts;
      Visio.Font visFont = visFonts[strFontFamily];
      VisualVisioUtil.SetIntCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_CHARFONT, visFont.ID);
      // Recentrage par rapport a la shape SVG
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblPinXValue + dblSVGPinXValue) - (dblSVGWidth * 0.5));
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblPinYValue + dblSVGPinYValue) + (dblSVGHeight * 0.5));
      }

    public static void CreateRectangleWithText(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblTranslateX, double dblTranslateY, double dblAngle, double dblWidthRatio, double dblHeightRatio,
                          double dblSVGWidth, double dblSVGHeight, bool bViewBox, string strFill, string strTrokeColor)
      {
      string strParamStrokeColor = "", strParamLocStrokeColor = "";
      string strParamStrokeWidth = "", strParamLocStrokeWidth = "";
      string strParamFill = "", strParamLocFill = "";
      string strRounding = "";
      int iRed = 0, iGreen = 0, iBlue = 0;
      Visio.Shape visShape;

      strParamStrokeColor = strTrokeColor;
      strParamFill = strFill;
      double dblX1 = ((SvgRectangle)element).X + dblTranslateX;
      double dblY1 = ((SvgRectangle)element).Y + dblTranslateY;
      double dblX2 = dblX1 + ((SvgRectangle)element).Width;
      double dblY2 = dblY1 + ((SvgRectangle)element).Height;
      ((SvgRectangle)element).TryGetAttribute("rx", out strRounding);
      ((SvgRectangle)element).TryGetAttribute("stroke", out strParamLocStrokeColor);
      ((SvgRectangle)element).TryGetAttribute("stroke-width", out strParamLocStrokeWidth);
      ((SvgRectangle)element).TryGetAttribute("fill", out strParamLocFill);
      if (strParamLocStrokeColor != null)
        strParamStrokeColor = strParamLocStrokeColor;
      if (strParamLocStrokeWidth != null)
        strParamStrokeWidth = strParamLocStrokeWidth;
      if (strParamLocFill != null)
        strParamFill = strParamLocFill;
      if (strTrokeColor != "")
        {
        strParamStrokeColor = strTrokeColor;
        if ((strParamLocStrokeColor != "") && (strParamLocStrokeColor != null))
          {
          strParamStrokeColor = strParamLocStrokeColor;
          }
        }
      if (strFill != "")
        {
        strParamFill = strFill;
        if ((strParamLocFill != "") && (strParamLocFill != null))
          {
          strParamFill = strParamLocFill;
          }
        }
      if (bViewBox)
        {
        dblX1 = visPage.Application.ConvertResult(dblX1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY1 = -visPage.Application.ConvertResult(dblY1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblX2 = visPage.Application.ConvertResult(dblX2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY2 = -visPage.Application.ConvertResult(dblY2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        visShape = visPage.DrawRectangle(dblX1 / dblWidthRatio, dblY1 / dblHeightRatio, dblX2 / dblWidthRatio, dblY2 / dblHeightRatio);
        // Rotation éventuelle
        if (dblAngle != 0)
          {
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
          // Centre de rotation à gauche au centre
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
          // Centre de rotation en haut à gauche
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
          // Rotation de la forme. Attention le signe de l'angle doit être inversé
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
          // Centre de rotation au centre en haut pour commencer à revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
          // Centre de rotation au centre au centre pour revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
          //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out dblPinXAfterRepos);
          //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out dblPinYAfterRepos);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          }
        // Recentrage par rapport a la shape SVG
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
        }
      else
        {
        dblX1 = visPage.Application.ConvertResult(dblX1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY1 = -visPage.Application.ConvertResult(dblY1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblX2 = visPage.Application.ConvertResult(dblX2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblY2 = -visPage.Application.ConvertResult(dblY2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        // Y coordinate is in the inverse direction of SVG so Y must be negative
        visShape = visPage.DrawRectangle(dblX1 / dblWidthRatio, dblY1 / dblHeightRatio, dblX2 / dblWidthRatio, dblY2 / dblHeightRatio);
        //visShape = visPage.DrawRectangle(dblX1, dblY1, dblX2, dblY2);
        // Rotation éventuelle
        if (dblAngle != 0)
          {
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
          // Centre de rotation à gauche au centre
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
          // Centre de rotation en haut à gauche
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
          // Rotation de la forme. Attention le signe de l'angle doit être inversé
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
          // Centre de rotation au centre en haut pour commencer à revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
          VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
          // Centre de rotation au centre au centre pour revenir à la position originale
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
          //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out dblPinXAfterRepos);
          //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out dblPinYAfterRepos);
          // repositionnement de la forme en X
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
          // repositionnement de la forme en Y
          VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
          }
        // Recentrage par rapport a la shape SVG
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblPinXValue + dblSVGPinXValue) - (dblSVGWidth * 0.5));
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblPinYValue + dblSVGPinYValue) + (dblSVGHeight * 0.5));
        }
      ApplyShapeStyles(visPage, visShape, strParamStrokeWidth, strRounding, strParamStrokeColor, strParamFill, dblWidthRatio);
      // Rajout du texte
      foreach (SvgElement subElement in (element.Parent).Children)
        {
        string strTransform = "";
        float fltAngle = 0.0F, fltX = 0.0F, fltY = 0.0F;
        subElement.TryGetAttribute("transform", out strTransform);

        if (!string.IsNullOrEmpty(strTransform))
          {
          if (subElement.Transforms.Count >= 1 && subElement.Transforms.ElementAt(0).GetType().Name == "SvgTranslate")
            {
            fltX = ((SvgTranslate)subElement.Transforms.ElementAt(0)).X;
            fltY = ((SvgTranslate)subElement.Transforms.ElementAt(0)).Y;
            }
          if (subElement.Transforms.Count >= 2 && subElement.Transforms.ElementAt(1).GetType().Name == "SvgRotate")
            {
            fltAngle = ((SvgRotate)subElement.Transforms.ElementAt(1)).Angle;
            }
          }
        foreach (SvgElement subChildElement in subElement.Children)
          {
          switch (subChildElement.GetType().Name)
            {
            case "SvgRectangle":
              break;
            case "SvgForeignObject":
              SvgForeignObject svgForeignObject = (SvgForeignObject)subChildElement;
              string strXML = svgForeignObject.GetXML();
              XDocument doc = XDocument.Parse(strXML);
              XNamespace xhtml = "http://www.w3.org/1999/xhtml";
              string paragraphValue = doc.Descendants(xhtml + "p").First().Value;
              visShape.Text = paragraphValue;
              break;
            default:
              break;
            }
          }

        }
      }

    public static void Create2DPolylineFromMarker(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblWidthRatio, double dblHeightRatio,
                              double dblSVGWidth, double dblSVGHeight)
      {
      string strStrokeColor, strStrokeWidth, strFill;

      ((SvgMarker)element).TryGetAttribute("stroke", out strStrokeColor);
      ((SvgMarker)element).TryGetAttribute("stroke-width", out strStrokeWidth);
      ((SvgMarker)element).TryGetAttribute("fill", out strFill);
      switch (((SvgMarker)element).Children[0].GetType().Name)
        {
        case "SvgPath":
          Create2DPolylineFromPath(visPage, visSVGShape, ((SvgMarker)element).Children[0], dblWidthRatio, dblHeightRatio, dblSVGWidth, dblSVGHeight);
          break;
        case "SvgCircle":
          break;
        }
      }

    private static void ApplyShapeStyles(Visio.Page visPage, Visio.Shape visShape, string strParamStrokeWidth, string strRounding, string strParamStrokeColor, string strParamFill, double dblWidthRatio)
      {
      if (strParamStrokeWidth != "")
        {
        Double.TryParse(strParamStrokeWidth, out double dblLineWeight);
        dblLineWeight = dblLineWeight / dblWidthRatio;
        NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
        VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_LINEWEIGHT, dblLineWeight.ToString("0.00 pt", nfi));
        }
      if (strRounding != "")
        {
        Double.TryParse(strRounding, out double dblRounding);
        dblRounding = visPage.Application.ConvertResult(dblRounding, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblRounding = dblRounding / dblWidthRatio;
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LINEROUNDING, dblRounding);
        }
      switch (strParamStrokeColor)
        {
        case "Blue":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, "RGB(0,112,192)");
          break;
        case "Green":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, "RGB(0,176,80)");
          break;
        case "Purple":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, "RGB(112,48,160)");
          break;
        case "none":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_LINEPATTERN, "0");
          break;
        default:
          break;
        }
      switch (strParamFill)
        {
        case "Blue":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(0,112,192)");
          break;
        case "Green":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(0,176,80)");
          break;
        case "Purple":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(112,48,160)");
          break;
        case "Red":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(255,0,0)");
          break;
        case "Yellow":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, "RGB(255,255,0)");
          break;
        case "None":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLPATTERN, "0");
          break;
        default:
          break;
        }
      }




    }
  }
