using ExCSS;
using Microsoft.Office.Interop.Visio;
using Svg;
using Svg.Pathing;
using Svg.Transforms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static VisualVisioSVGLight.VLConstants;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisualVisioSVGLight
  {
  internal class VisualVisioSVGLightUtil
    {

    public static void CreateRect(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblTranslateX, double dblTranslateY, double dblAngle, double dblWidthRatio, double dblHeightRatio,
                        double dblSVGWidth, double dblSVGHeight, bool bViewBox, string strFill, string strStrokeColor, string strOpacity)
      {
      string strParamStrokeColor = "", strParamLocStrokeColor = "";
      string strParamStrokeWidth = "", strParamLocStrokeWidth = "";
      string strParamFill = "", strParamLocFill = "";
      string strParamOpacity = "", strParamLocOpacity = "";
      string strRounding = "";
      int iRed = 0, iGreen = 0, iBlue = 0;
      Visio.Shape visShape;

      strParamStrokeColor = strStrokeColor;
      strParamFill = strFill;
      double dblX1 = ((SvgRectangle)element).X + dblTranslateX;
      double dblY1 = ((SvgRectangle)element).Y + dblTranslateY;
      double dblX2 = dblX1 + ((SvgRectangle)element).Width;
      double dblY2 = dblY1 + ((SvgRectangle)element).Height;
      ((SvgRectangle)element).TryGetAttribute("rx", out strRounding);
      ((SvgRectangle)element).TryGetAttribute("stroke", out strParamLocStrokeColor);
      ((SvgRectangle)element).TryGetAttribute("stroke-width", out strParamLocStrokeWidth);
      ((SvgRectangle)element).TryGetAttribute("fill", out strParamLocFill);
      ((SvgRectangle)element).TryGetAttribute("opacity", out strParamLocOpacity);
      if (strParamLocStrokeColor != null)
        strParamStrokeColor = strParamLocStrokeColor;
      if (strParamLocStrokeWidth != null)
        strParamStrokeWidth = strParamLocStrokeWidth;
      if (strParamLocFill != null)
        strParamFill = strParamLocFill;
      if (strParamLocOpacity != null)
        strParamOpacity = strParamLocOpacity;
      if (strStrokeColor != "")
        {
        strParamStrokeColor = strStrokeColor;
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
      if (strOpacity != "")
        {
        strParamOpacity = strOpacity;
        if ((strParamLocOpacity != "") && (strParamLocOpacity != null))
          {
          strParamOpacity = strParamLocOpacity;
          }
        }
      dblX1 = visPage.Application.ConvertResult(dblX1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblY1 = -visPage.Application.ConvertResult(dblY1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblX2 = visPage.Application.ConvertResult(dblX2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblY2 = -visPage.Application.ConvertResult(dblY2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      visShape = visPage.DrawRectangle(dblX1 / dblWidthRatio, dblY1 / dblHeightRatio, dblX2 / dblWidthRatio, dblY2 / dblHeightRatio);
      // Possible rotation
      if (dblAngle != 0)
        {
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
        // Rotation center shifted to the left-center
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
        // Rotation center at the top-left
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
        // Rotating the shape. Note that the angle sign must be inverted
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
        // Centre de rotation au centre en haut pour commencer à revenir à la position originale
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
        // Rotation center at the center-top to start returning to the original position
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        }
      // Re-centering relative to the SVG shape
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
      ApplyShapeStyles(visPage, visShape, strParamStrokeWidth, strRounding, strParamStrokeColor, strParamFill, strParamOpacity, dblWidthRatio);
      }

    public static Visio.Shape CreateCircle(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblTranslateX, double dblTranslateY, double dblAngle, double dblWidthRatio, double dblHeightRatio,
                          double dblSVGWidth, double dblSVGHeight, double dblMarkerWidthRatio, double dblMarkerHeightRatio, bool bViewBox, string strFill, string strStrokeColor, string strOpacity)
      {
      string strParamStrokeColor = "", strParamLocStrokeColor = "";
      string strParamStrokeWidth = "", strParamLocStrokeWidth = "";
      string strParamFill = "", strParamLocFill = "";
      string strParamOpacity = "", strParamLocOpacity = "";
      string strRounding = "";
      Visio.Shape visShape = null;

      double dblCenterX = ((SvgCircle)element).CenterX + dblTranslateX;
      double dblCenterY = ((SvgCircle)element).CenterY + dblTranslateY;
      double dblRadius = ((SvgCircle)element).Radius;
      double dblDiameter = (dblRadius * 2.0F) * dblMarkerWidthRatio;
      double dblX1 = dblCenterX - (dblDiameter * 0.5);
      double dblY1 = dblCenterY - (dblDiameter * 0.5);
      double dblX2 = dblX1 + dblDiameter;
      double dblY2 = dblY1 + dblDiameter;

      ((SvgCircle)element).TryGetAttribute("stroke", out strParamLocStrokeColor);
      ((SvgCircle)element).TryGetAttribute("stroke-width", out strParamLocStrokeWidth);
      ((SvgCircle)element).TryGetAttribute("fill", out strParamLocFill);
      ((SvgCircle)element).TryGetAttribute("opacity", out strParamLocOpacity);
      if (strParamLocStrokeColor != null)
        strParamStrokeColor = strParamLocStrokeColor;
      if (strParamLocStrokeWidth != null)
        strParamStrokeWidth = strParamLocStrokeWidth;
      if (strParamLocFill != null)
        strParamFill = strParamLocFill;
      if (strParamLocOpacity != null)
        strParamOpacity = strParamLocOpacity;
      if (!string.IsNullOrEmpty(strStrokeColor))
        {
        strParamStrokeColor = strStrokeColor;
        if (!string.IsNullOrEmpty(strParamLocStrokeColor))
          {
          strParamStrokeColor = strParamLocStrokeColor;
          }
        }
      if (!string.IsNullOrEmpty(strFill))
        {
        strParamFill = strFill;
        if (!string.IsNullOrEmpty(strParamLocFill))
          {
          strParamFill = strParamLocFill;
          }
        }
      if (!string.IsNullOrEmpty(strOpacity))
        {
        strParamOpacity = strOpacity;
        if (!string.IsNullOrEmpty(strParamLocOpacity))
          {
          strParamOpacity = strParamLocOpacity;
          }
        }
      dblX1 = visPage.Application.ConvertResult(dblX1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblY1 = -visPage.Application.ConvertResult(dblY1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblX2 = visPage.Application.ConvertResult(dblX2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblY2 = -visPage.Application.ConvertResult(dblY2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      visShape = visPage.DrawOval(dblX1 / dblWidthRatio, dblY1 / dblHeightRatio, dblX2 / dblWidthRatio, dblY2 / dblHeightRatio);
      // Possible rotation
      if (dblAngle != 0)
        {
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
        // Rotation center shifted to the left-center
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
        // Rotation center at the top-left
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
        // Rotating the shape. Note that the angle sign must be inverted
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
        // Centre de rotation au centre en haut pour commencer à revenir à la position originale
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
        // Rotation center at the center-top to start returning to the original position
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        }
      // Re-centering relative to the SVG shape
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
      ApplyShapeStyles(visPage, visShape, strParamStrokeWidth, strRounding, strParamStrokeColor, strParamFill, strParamOpacity, dblWidthRatio);
      return visShape;
      }

    public static void CreateEllipse(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblTranslateX, double dblTranslateY, double dblAngle, double dblWidthRatio, double dblHeightRatio,
                          double dblSVGWidth, double dblSVGHeight, bool bViewBox, string strFill, string strStrokeColor, string strOpacity)
      {
      string strParamStrokeColor = "", strParamLocStrokeColor = "";
      string strParamStrokeWidth = "", strParamLocStrokeWidth = "";
      string strParamFill = "", strParamLocFill = "";
      string strParamOpacity = "", strParamLocOpacity = "";
      string strTransform = "";
      string strRounding = "";
      Visio.Shape visShape;

      ((SvgEllipse)element).TryGetAttribute("transform", out strTransform);
      ((SvgEllipse)element).TryGetAttribute("stroke", out strParamLocStrokeColor);
      ((SvgEllipse)element).TryGetAttribute("stroke-width", out strParamLocStrokeWidth);
      ((SvgEllipse)element).TryGetAttribute("fill", out strParamLocFill);
      ((SvgEllipse)element).TryGetAttribute("opacity", out strParamLocOpacity);
      if (!string.IsNullOrEmpty(strTransform))
        {
        if (element.Transforms.Count >= 1 && element.Transforms.ElementAt(0).GetType().Name == "SvgTranslate")
          {
          dblTranslateX = ((SvgTranslate)element.Transforms.ElementAt(0)).X;
          dblTranslateY = ((SvgTranslate)element.Transforms.ElementAt(0)).Y;
          }
        if (element.Transforms.Count >= 2 && element.Transforms.ElementAt(1).GetType().Name == "SvgRotate")
          {
          dblAngle = ((SvgRotate)element.Transforms.ElementAt(1)).Angle;
          }
        }
      double dblCenterX = ((SvgEllipse)element).CenterX;
      double dblCenterY = ((SvgEllipse)element).CenterY;
      double dblRadiusX = ((SvgEllipse)element).RadiusX;
      double dblRadiusY = ((SvgEllipse)element).RadiusY;
      double dblDiameter = dblRadiusX * 2.0F;
      double dblX1 = dblTranslateX - dblRadiusX;
      double dblY1 = dblTranslateY - dblRadiusY;
      double dblX2 = dblX1 + (dblRadiusX * 2);
      double dblY2 = dblY1 + (dblRadiusY * 2);
      if (strParamLocStrokeColor != null)
        strParamStrokeColor = strParamLocStrokeColor;
      if (strParamLocStrokeWidth != null)
        strParamStrokeWidth = strParamLocStrokeWidth;
      if (strParamLocFill != null)
        strParamFill = strParamLocFill;
      if (strParamLocOpacity != null)
        strParamOpacity = strParamLocOpacity;
      if (strStrokeColor != "")
        {
        strParamStrokeColor = strStrokeColor;
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
      if (strOpacity != "")
        {
        strParamOpacity = strOpacity;
        if ((strParamLocOpacity != "") && (strParamLocOpacity != null))
          {
          strParamOpacity = strParamLocOpacity;
          }
        }
      dblX1 = visPage.Application.ConvertResult(dblX1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblY1 = -visPage.Application.ConvertResult(dblY1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblX2 = visPage.Application.ConvertResult(dblX2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblY2 = -visPage.Application.ConvertResult(dblY2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      visShape = visPage.DrawOval(dblX1 / dblWidthRatio, dblY1 / dblHeightRatio, dblX2 / dblWidthRatio, dblY2 / dblHeightRatio);
      // Possible rotation
      if (dblAngle != 0)
        {
        //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
        //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
        //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
        //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
        //// Rotation center shifted to the left-center
        //VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
        //// repositioning the shape along the X-axis
        //VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
        //// Rotation center at the top-left
        //VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
        //// Repositioning the shape along the Y-axis
        //VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
        // Rotating the shape. Note that the angle sign must be inverted
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
        //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
        //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
        //// Centre de rotation au centre en haut pour commencer à revenir à la position originale
        //VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
        //// repositioning the shape along the X-axis
        //VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        //// Repositioning the shape along the Y-axis
        //VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
        //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
        //// Rotation center at the center-top to start returning to the original position
        //VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
        //// repositioning the shape along the X-axis
        //VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        //// Repositioning the shape along the Y-axis
        //VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        }
      // Re-centering relative to the SVG shape
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
      ApplyShapeStyles(visPage, visShape, strParamStrokeWidth, strRounding, strParamStrokeColor, strParamFill, strParamOpacity, dblWidthRatio);
      }

    public static void CreatePolyline(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblTranslateX, double dblTranslateY, double dblAngle, double dblWidthRatio, double dblHeightRatio,
                          double dblSVGWidth, double dblSVGHeight, bool bViewBox, string strFill, string strStrokeColor, string strOpacity)
      {
      string strParamStrokeColor = "", strParamLocStrokeColor = "";
      string strParamStrokeWidth = "", strParamLocStrokeWidth = "";
      string strParamFill = "", strParamLocFill = "";
      string strParamOpacity = "", strParamLocOpacity = "";
      Visio.Shape visShape;

      strParamStrokeColor = strStrokeColor;
      strParamFill = strFill;
      ((SvgPolyline)element).TryGetAttribute("stroke", out strParamLocStrokeColor);
      ((SvgPolyline)element).TryGetAttribute("stroke-width", out strParamLocStrokeWidth);
      ((SvgPolyline)element).TryGetAttribute("fill", out strParamLocFill);
      ((SvgPolyline)element).TryGetAttribute("opacity", out strParamLocOpacity);
      if (strParamLocStrokeColor != null)
        strParamStrokeColor = strParamLocStrokeColor;
      if (strParamLocStrokeWidth != null)
        strParamStrokeWidth = strParamLocStrokeWidth;
      if (strParamLocFill != null)
        strParamFill = strParamLocFill;
      if (strParamLocOpacity != null)
        strParamOpacity = strParamLocOpacity;
      if (strStrokeColor != "")
        {
        strParamStrokeColor = strStrokeColor;
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
      if (strOpacity != "")
        {
        strParamOpacity = strOpacity;
        if ((strParamLocOpacity != "") && (strParamLocOpacity != null))
          {
          strParamOpacity = strParamLocOpacity;
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
      visShape = visPage.DrawPolyline(arPoint, (int)Visio.VisDrawSplineFlags.visPolyline1D);
      // Possible rotation
      if (dblAngle != 0)
        {
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
        // Rotation center shifted to the left-center
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
        // Rotation center at the top-left
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
        // Rotating the shape. Note that the angle sign must be inverted
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
        // Centre de rotation au centre en haut pour commencer à revenir à la position originale
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
        // Rotation center at the center-top to start returning to the original position
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        }
      // Re-centering relative to the SVG shape
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
      ApplyShapeStyles(visPage, visShape, strParamStrokeWidth, "", strParamStrokeColor, strParamFill, strParamOpacity, dblWidthRatio);
      }


    public static void CreatePolygon(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblTranslateX, double dblTranslateY, double dblAngle, double dblWidthRatio, double dblHeightRatio,
                        double dblSVGWidth, double dblSVGHeight, bool bViewBox, string strFill, string strStrokeColor, string strOpacity)
      {
      string strParamStrokeColor = "", strParamLocStrokeColor = "";
      string strParamStrokeWidth = "", strParamLocStrokeWidth = "";
      string strParamFill = "", strParamLocFill = "";
      string strParamOpacity = "", strParamLocOpacity = "";
      Visio.Shape visShape;

      strParamStrokeColor = strStrokeColor;
      strParamFill = strFill;
      ((SvgPolygon)element).TryGetAttribute("stroke", out strParamLocStrokeColor);
      ((SvgPolygon)element).TryGetAttribute("stroke-width", out strParamLocStrokeWidth);
      ((SvgPolygon)element).TryGetAttribute("fill", out strParamLocFill);
      ((SvgPolygon)element).TryGetAttribute("opacity", out strParamLocOpacity);
      if (strParamLocStrokeColor != null)
        strParamStrokeColor = strParamLocStrokeColor;
      if (strParamLocStrokeWidth != null)
        strParamStrokeWidth = strParamLocStrokeWidth;
      if (strParamLocFill != null)
        strParamFill = strParamLocFill;
      if (strParamLocOpacity != null)
        strParamOpacity = strParamLocOpacity;
      if (strStrokeColor != "")
        {
        strParamStrokeColor = strStrokeColor;
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
      if (strOpacity != "")
        {
        strParamOpacity = strOpacity;
        if ((strParamLocOpacity != "") && (strParamLocOpacity != null))
          {
          strParamOpacity = strParamLocOpacity;
          }
        }
      int nbPoints = ((SvgPolygon)element).Points.Count;
      double[] arPoint = new double[nbPoints];
      for (int i = 0; i < nbPoints; i++)
        {
        if (i % 2 == 0)
          {
          arPoint[i] = visPage.Application.ConvertResult(((SvgPolygon)element).Points[i].Value + dblTranslateX,
                                                         (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches)
                                                         / dblWidthRatio;
          }
        else
          {
          arPoint[i] = visPage.Application.ConvertResult(-((SvgPolygon)element).Points[i].Value + dblTranslateY,
                                                         (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches)
                                                         / dblHeightRatio;
          }
        }
      visShape = visPage.DrawPolyline(arPoint, (int)Visio.VisDrawSplineFlags.visPolyline1D);
      // Get the count of rows in the Geometry section. 
      int iGeometryLineCount = visShape.RowCount[(short)Visio.VisSectionIndices.visSectionFirstComponent];
      // Close the polyline
      // We must add a rellineto as the DrawPolyline create a RelMoveTo
      visShape.AddRow((short)Visio.VisSectionIndices.visSectionFirstComponent, (int)Visio.VisRowIndices.visRowLast,
                          (int)Visio.VisRowTags.visTagRelLineTo);
      VisualVisioUtil.SetFormulaCell(visShape, (int)Visio.VisSectionIndices.visSectionFirstComponent,
                    (int)Visio.VisRowIndices.visRowVertex + iGeometryLineCount - 1, (int)Visio.VisCellIndices.visX, "Geometry1.X1");
      VisualVisioUtil.SetFormulaCell(visShape, (int)Visio.VisSectionIndices.visSectionFirstComponent,
                    (int)Visio.VisRowIndices.visRowVertex + iGeometryLineCount - 1, (int)Visio.VisCellIndices.visY, "Geometry1.Y1");
      VisualVisioUtil.SetFormulaCell(visShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent),
                    (int)Visio.VisRowIndices.visRowFirst, (int)Visio.VisCellIndices.visCompNoFill, false.ToString());
      // Possible rotation
      if (dblAngle != 0)
        {
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
        // Rotation center shifted to the left-center
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
        // Rotation center at the top-left
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
        // Rotating the shape. Note that the angle sign must be inverted
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
        // Centre de rotation au centre en haut pour commencer à revenir à la position originale
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
        // Rotation center at the center-top to start returning to the original position
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        }
      // Re-centering relative to the SVG shape
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
      ApplyShapeStyles(visPage, visShape, strParamStrokeWidth, "", strParamStrokeColor, strParamFill, strParamOpacity, dblWidthRatio);
      }

    public static void Create2DPolylineFromPath(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, string styleContent, double dblWidthRatio, double dblHeightRatio,
                                  double dblSVGWidth, double dblSVGHeight, bool bHide)
      {
      double dblBeginX = 0.0, dblBeginY = 0.0;
      double dblOriginX = 0.0, dblOriginY = 0.0, dblExtremityX = 0.0, dblExtremityY = 0.0;
      double dblPathOriginX = 0.0, dblPathOriginY = 0.0;
      double dblRelOriginX = 0.0, dblRelOriginY = 0.0;
      double dblArcRelOriginX = 0, dblArcRelOriginY = 0;
      Visio.Shape visPathShape=null;
      string strStrokeColor, strStrokeWidth, strFill, strOpacity;
      string strStartMarker = "", strMidMarker = "", strEndMarker = "";
      Visio.Shape visStartMarkerShape = null, visMidMarkerShape = null, visEndMarkerShape = null;
      string strStyleFill = "", strStyleStrokeWidth="", strStyleMarker = "";

      using (StringReader reader = new StringReader(styleContent))
        {
        string strLine;
        bool bPathFound = false;
        while ((strLine = reader.ReadLine()) != null)
          {
          strLine = strLine.Trim();
          // Parse specific CSS properties
          if (strLine.StartsWith("path"))
            {
            bPathFound = true;
            }
          else if (strLine.Contains("fill:") && bPathFound)
            {
            strStyleFill = strLine.Replace("fill:","");
            strStyleFill = strStyleFill.Replace(";", "");
            strStyleFill = strStyleFill.Trim();
            }
          else if (strLine.Contains("stroke-width:"))
            {
            strStyleStrokeWidth = strLine.Replace("stroke-width:", ""); ;
            strStyleStrokeWidth = strStyleStrokeWidth.Replace(";", "");
            strStyleStrokeWidth = strStyleStrokeWidth.Replace("px", "");
            strStyleStrokeWidth = strStyleStrokeWidth.Trim();
            }
          else if (strLine.Contains("marker:"))
            {
            strStyleMarker = strLine.Replace("marker:", ""); ;
            strStyleMarker = strStyleMarker.Replace(";", "");
            strStyleMarker = strStyleMarker.Trim();
            }
          }
        }
      ((SvgPath)element).TryGetAttribute("stroke", out strStrokeColor);
      ((SvgPath)element).TryGetAttribute("stroke-width", out strStrokeWidth);
      if (String.IsNullOrEmpty(strStrokeWidth))
        strStrokeWidth = strStyleStrokeWidth;
      ((SvgPath)element).TryGetAttribute("fill", out strFill);
      if (String.IsNullOrEmpty(strFill))
        strFill = strStyleFill;
      ((SvgPath)element).TryGetAttribute("opacity", out strOpacity);
      ((SvgPath)element).TryGetAttribute("marker-start", out strStartMarker);
      ((SvgPath)element).TryGetAttribute("marker-mid", out strMidMarker);
      if (String.IsNullOrEmpty(strMidMarker))
        strMidMarker = strStyleMarker;
      ((SvgPath)element).TryGetAttribute("marker-end", out strEndMarker);
      if (!String.IsNullOrEmpty(strStartMarker))
        {
        strStartMarker = strStartMarker.Replace("url(#", "");
        strStartMarker = strStartMarker.Replace(")", "");
        VisualVisioUtil.GetVisShape(visPage, strStartMarker, out visStartMarkerShape);
        }
      if (!String.IsNullOrEmpty(strMidMarker))
        {
        strMidMarker = strMidMarker.Replace("url(#", "");
        strMidMarker = strMidMarker.Replace(")", "");
        VisualVisioUtil.GetVisShape(visPage, strMidMarker, out visMidMarkerShape);
        }
      if (!String.IsNullOrEmpty(strEndMarker))
        {
        strEndMarker = strEndMarker.Replace("url(#", "");
        strEndMarker = strEndMarker.Replace(")", "");
        VisualVisioUtil.GetVisShape(visPage, strEndMarker, out visEndMarkerShape);
        }
      SvgPath svgPath = ((SvgPath)element);
      Svg.Pathing.SvgPathSegmentList arData = svgPath.PathData;
      // Search the number of begin vertices
      string strPath = svgPath.PathData.ToString();
      int iCountBeginVertices = strPath.Count(c => c == 'M' || c == 'm');
      int iGeometry = 0;
      int iGeometryLine = 0;
      dblOriginX = arData[0].End.X;
      dblOriginY = arData[0].End.Y;
      dblPathOriginX = arData[0].End.X;
      dblPathOriginY = arData[0].End.Y;
      dblRelOriginX = dblOriginX;
      dblRelOriginY = dblOriginY;
      dblArcRelOriginX = dblOriginX;
      dblArcRelOriginY = dblOriginY;
      //if (arData[arData.Count - 1].ToString().Substring(0, 1) == "Z" || arData[arData.Count - 1].ToString().Substring(0, 1) == "z")
      //  {
      //  dblExtremityX = arData[arData.Count - 2].End.X;
      //  dblExtremityY = arData[arData.Count - 2].End.Y;
      //  }
      //else
      //  {
      //  dblExtremityX = arData[arData.Count - 1].End.X;
      //  dblExtremityY = arData[arData.Count - 1].End.Y;
      //  }
      dblOriginX = visPage.Application.ConvertResult(dblOriginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblOriginY = -visPage.Application.ConvertResult(dblOriginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblExtremityX = visPage.Application.ConvertResult(dblExtremityX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblExtremityY = -visPage.Application.ConvertResult(dblExtremityY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      bool bCubicBezier = false;
      bool bSmoothCubicBezier = false;
      bool bPathZEndFound = false;
      double dblCurrentPointX = 0.0, dblCurrentPointY = 0.0;
      double dblReflexionPointX = 0.0, dblReflexionPointY = 0.0;
      int iBeginVertices = 0;
      foreach (Svg.Pathing.SvgPathSegment pathSegment in arData)
        {
        string strPoint = pathSegment.ToString();
        double dblCubicEndPointX = 0.0, dblCubicEndPointY = 0.0;
        switch (strPoint.Substring(0, 1))
          {
          case "M":
            if (bPathZEndFound)
              {
              if (bHide)
                {
                if(visPathShape != null)
                  VisualVisioUtil.SetGeometryVisibility(visPathShape, false, false);
                }
              // Create new PathShape
              iGeometry = 0;
              iGeometryLine = 0;
              dblOriginX = visPage.Application.ConvertResult(dblCurrentPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblOriginY = -visPage.Application.ConvertResult(dblCurrentPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              visPathShape = visPage.DrawRectangle(0, 0, 0, 0);
              bPathZEndFound = false;
              }
            else
              {
              visPathShape = visPage.DrawRectangle(dblOriginX / dblWidthRatio, dblOriginY / dblHeightRatio, dblOriginX / dblWidthRatio, dblOriginY / dblHeightRatio);
              // Deleting of all LineTo
              int nbGeometry = visPathShape.GeometryCount;
              // Get the count of rows in the current Geometry section. 
              int nbRows = visPathShape.RowCount[(short)Visio.VisSectionIndices.visSectionFirstComponent];

              for (int iRow = 1; iRow < nbRows - 1; iRow++)
                {
                visPathShape.DeleteRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (short)(Visio.VisRowIndices.visRowVertex + 1));
                }
              }
            // MoveTo
            if (iGeometryLine > 0)
              {
              // Create a new geometry line
              visPathShape.AddRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (int)Visio.VisRowIndices.visRowLast, (int)Visio.VisRowTags.visTagMoveTo);
              }
            dblBeginX = pathSegment.End.X;
            dblBeginY = pathSegment.End.Y;
            dblCurrentPointX = dblBeginX;
            dblCurrentPointY = dblBeginY;
            dblBeginX = visPage.Application.ConvertResult(dblBeginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblBeginY = -visPage.Application.ConvertResult(dblBeginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            if (InsertStartMarker(visPage, visSVGShape, strStrokeColor, dblSVGWidth, dblSVGHeight, visStartMarkerShape, visMidMarkerShape, visEndMarkerShape, iBeginVertices, iCountBeginVertices,
                                 dblRelOriginX, dblRelOriginY))
              iBeginVertices++;
            iGeometryLine++;
            break;
          case "m":
            if (bPathZEndFound)
              {
              // We must create another Path Shape
              // Fisrt finishing correctly previous path shape
              if (visPathShape != null)
                {
                visPathShape.UpdateAlignmentBox();
                // Re-centering relative to the SVG shape
                VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblEndFoundSVGPinXValue);
                VisualVisioUtil.GetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblEndFoundPinXValue);
                VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblEndFoundPinXValue + dblEndFoundSVGPinXValue) - (dblSVGWidth * 0.5));
                VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblEndFoundSVGPinYValue);
                VisualVisioUtil.GetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblEndFoundPinYValue);
                VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblEndFoundPinYValue + dblEndFoundSVGPinYValue) + (dblSVGHeight * 0.5));
                ApplyShapeStyles(visPage, visPathShape, strStrokeWidth, "", strStrokeColor, strFill, strOpacity, dblWidthRatio);
                if (bHide)
                  {
                  VisualVisioUtil.SetGeometryVisibility(visPathShape, false, false);
                  }
                }
              // Create new PathShape
              iGeometry = 0;
              iGeometryLine = 0;
              dblBeginX = dblPathOriginX + pathSegment.End.X;
              dblBeginY = dblPathOriginY + pathSegment.End.Y;
              dblBeginX = visPage.Application.ConvertResult(dblBeginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblBeginY = -visPage.Application.ConvertResult(dblBeginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              visPathShape = visPage.DrawRectangle(dblBeginX / dblWidthRatio, dblBeginY / dblHeightRatio, dblBeginX / dblWidthRatio, dblBeginY / dblHeightRatio);
              dblBeginX = dblPathOriginX + pathSegment.End.X;
              dblBeginY = dblPathOriginY + pathSegment.End.Y;
              dblRelOriginX = 0;
              dblRelOriginY = 0;
              dblCurrentPointX = 0;
              dblCurrentPointY = 0;
              dblPathOriginX = dblBeginX;
              dblPathOriginY = dblBeginY;
              // Deleting of all LineTo
              int nGeometry = visPathShape.GeometryCount;
              // Get the count of rows in the current Geometry section. 
              int nbRows = visPathShape.RowCount[(short)Visio.VisSectionIndices.visSectionFirstComponent];
              for (int iRow = 1; iRow < nbRows - 1; iRow++)
                {
                visPathShape.DeleteRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (short)(Visio.VisRowIndices.visRowVertex + 1));
                }
              bPathZEndFound = false;
              }
            else
              {
              dblBeginX = pathSegment.End.X;
              dblBeginY = pathSegment.End.Y;
              }
            // MoveTo
            if (iGeometryLine > 0)
              {
              // Create a new geometry line
              visPathShape.AddRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (int)Visio.VisRowIndices.visRowLast, (int)Visio.VisRowTags.visTagMoveTo);
              }
            if (InsertStartMarker(visPage, visSVGShape, strStrokeColor, dblSVGWidth, dblSVGHeight, visStartMarkerShape, visMidMarkerShape, visEndMarkerShape, iBeginVertices, iCountBeginVertices,
                                 dblBeginX, dblBeginY))
              iBeginVertices++;
            iGeometryLine++;
            break;
          case "L":
            // Absolute line
            if (iGeometryLine > 0)
              {
              // Create a new geometry line
              visPathShape.AddRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (int)Visio.VisRowIndices.visRowLast, (int)Visio.VisRowTags.visTagLineTo);
              }
            dblBeginX = pathSegment.End.X;
            dblBeginY = pathSegment.End.Y;
            dblCurrentPointX = dblBeginX;
            dblCurrentPointY = dblBeginY;
            dblBeginX = visPage.Application.ConvertResult(dblBeginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblBeginY = -visPage.Application.ConvertResult(dblBeginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                          (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visX, (dblBeginX - dblOriginX) / dblWidthRatio);
            VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                          (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visY, (dblBeginY - dblOriginY) / dblHeightRatio);
            iGeometryLine++; break;
          case "l":
            // Rel LineTo
            if (iGeometryLine > 0)
              {
              // Create a new geometry line
              visPathShape.AddRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (int)Visio.VisRowIndices.visRowLast, (int)Visio.VisRowTags.visTagLineTo);
              }
            dblBeginX = (dblCurrentPointX - dblRelOriginX) + pathSegment.End.X;
            dblBeginY = (dblCurrentPointY - dblRelOriginY) + pathSegment.End.Y;
            dblCurrentPointX = dblBeginX;
            dblCurrentPointY = dblBeginY;
            dblBeginX = visPage.Application.ConvertResult(dblBeginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblBeginY = -visPage.Application.ConvertResult(dblBeginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                          (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visX, (dblBeginX) / dblWidthRatio);
            VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                          (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visY, (dblBeginY) / dblHeightRatio);
            iGeometryLine++;
            break;
          case "C":
            // CubicBezier curveto
            // Current point
            double dblStartX = dblCurrentPointX;
            double dblStartY = dblCurrentPointY;
            // First control point
            double dblFirstControlPointX = ((SvgCubicCurveSegment)pathSegment).FirstControlPoint.X;
            double dblFirstControlPointY = ((SvgCubicCurveSegment)pathSegment).FirstControlPoint.Y;
            // Second control point
            double dblSecondControlPointX = ((SvgCubicCurveSegment)pathSegment).SecondControlPoint.X;
            double dblSecondControlPointY = ((SvgCubicCurveSegment)pathSegment).SecondControlPoint.Y;
            dblReflexionPointX = dblSecondControlPointX;
            dblReflexionPointY = dblSecondControlPointY;
            // End point
            dblCubicEndPointX = ((PointF)pathSegment.End).X;
            dblCubicEndPointY = ((PointF)pathSegment.End).Y;
            dblCurrentPointX = dblCubicEndPointX;
            dblCurrentPointY = dblCubicEndPointY;
            dblStartX = visPage.Application.ConvertResult(dblStartX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblStartY = -visPage.Application.ConvertResult(dblStartY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblFirstControlPointX = visPage.Application.ConvertResult(dblFirstControlPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblFirstControlPointY = -visPage.Application.ConvertResult(dblFirstControlPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblSecondControlPointX = visPage.Application.ConvertResult(dblSecondControlPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblSecondControlPointY = -visPage.Application.ConvertResult(dblSecondControlPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblCubicEndPointX = visPage.Application.ConvertResult(dblCubicEndPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblCubicEndPointY = -visPage.Application.ConvertResult(dblCubicEndPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            double[] arControlPoint = new double[8];
            // Current point
            arControlPoint[0] = dblStartX / dblWidthRatio;
            arControlPoint[1] = dblStartY / dblHeightRatio;
            // First control point
            arControlPoint[2] = (dblFirstControlPointX) / dblWidthRatio;
            //arControlPoint[3] = (dblFirstControlPointY + dblOriginY + (dblBeginY * 0.5)) / dblHeightRatio;
            arControlPoint[3] = (dblFirstControlPointY) / dblHeightRatio;
            // Second control point
            arControlPoint[4] = (dblSecondControlPointX) / dblWidthRatio;
            arControlPoint[5] = (dblSecondControlPointY) / dblHeightRatio;
            // End point
            arControlPoint[6] = (dblCubicEndPointX) / dblWidthRatio;
            arControlPoint[7] = (dblCubicEndPointY) / dblHeightRatio;
            Visio.Shape visShapeBezier = visPage.DrawBezier(arControlPoint, 3, (int)VisDrawSplineFlags.visSpline1D);
            visShapeBezier.UpdateAlignmentBox();
            visShapeBezier.BoundingBox((int)VisBoundingBoxArgs.visBBoxExtents, out double dblLeft1, out double dblBottom1, out double dblRight1, out double dblTop1);
            // Re-centering relative to the SVG shape
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue1);
            VisualVisioUtil.GetDoubleCellVal(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue1);
            VisualVisioUtil.SetDoubleCellVal(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblPinXValue1 + dblSVGPinXValue1) - (dblSVGWidth * 0.5));
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue1);
            VisualVisioUtil.GetDoubleCellVal(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue1);
            VisualVisioUtil.SetDoubleCellVal(visShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblPinYValue1 + dblSVGPinYValue1) + (dblSVGHeight * 0.5));
            ApplyShapeStyles(visPage, visShapeBezier, strStrokeWidth, "", strStrokeColor, strFill, strOpacity, dblWidthRatio);
            bCubicBezier = true;
            break;
          case "c":
            break;
          case "S":
            // CubicBezier shorthand/smooth curveto
            double dblSmoothFirstControlPointX, dblSmoothFirstControlPointY;
            double dblSmoothSecondControlPointX, dblSmoothSecondControlPointY;
            double dblSmoothEndPointX, dblSmoothEndPointY;
            double[] arSmoothControlPoint = null;
            if (bCubicBezier)
              {
              // Current point
              double dblSmoothStartX = dblCurrentPointX;
              double dblSmoothStartY = dblCurrentPointY;
              // The first control point is the reflection on the previous second control point
              dblSmoothFirstControlPointX = dblCurrentPointX;
              dblSmoothFirstControlPointY = dblCurrentPointY + dblReflexionPointY;
              // Second control point
              dblSmoothSecondControlPointX = ((SvgCubicCurveSegment)pathSegment).SecondControlPoint.X;
              dblSmoothSecondControlPointY = ((SvgCubicCurveSegment)pathSegment).SecondControlPoint.Y;
              // End point
              dblSmoothEndPointX = ((PointF)pathSegment.End).X;
              dblSmoothEndPointY = ((PointF)pathSegment.End).Y;
              dblCurrentPointX = dblSmoothEndPointX;
              dblCurrentPointY = dblSmoothEndPointY;
              dblSmoothStartX = visPage.Application.ConvertResult(dblSmoothStartX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblSmoothStartY = -visPage.Application.ConvertResult(dblSmoothStartY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblSmoothFirstControlPointX = visPage.Application.ConvertResult(dblSmoothFirstControlPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblSmoothFirstControlPointY = -visPage.Application.ConvertResult(dblSmoothFirstControlPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblSmoothSecondControlPointX = visPage.Application.ConvertResult(dblSmoothSecondControlPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblSmoothSecondControlPointY = -visPage.Application.ConvertResult(dblSmoothSecondControlPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblSmoothEndPointX = visPage.Application.ConvertResult(dblSmoothEndPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblSmoothEndPointY = -visPage.Application.ConvertResult(dblSmoothEndPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              arSmoothControlPoint = new double[8];
              // Current point
              arSmoothControlPoint[0] = (dblSmoothStartX) / dblWidthRatio;
              arSmoothControlPoint[1] = (dblSmoothStartY) / dblHeightRatio;
              // First control point
              arSmoothControlPoint[2] = (dblSmoothFirstControlPointX) / dblWidthRatio;
              arSmoothControlPoint[3] = (dblSmoothFirstControlPointY) / dblHeightRatio;
              // Second control point
              arSmoothControlPoint[4] = (dblSmoothSecondControlPointX) / dblWidthRatio;
              arSmoothControlPoint[5] = (dblSmoothSecondControlPointY) / dblHeightRatio;
              // End point
              arSmoothControlPoint[6] = (dblSmoothEndPointX) / dblWidthRatio;
              arSmoothControlPoint[7] = (dblSmoothEndPointY) / dblHeightRatio;
              }
            else
              {
              dblSmoothFirstControlPointX = dblReflexionPointX;
              dblSmoothFirstControlPointY = dblReflexionPointY;
              dblSmoothFirstControlPointX = ((SvgCubicCurveSegment)pathSegment).FirstControlPoint.X;
              dblSmoothFirstControlPointY = ((SvgCubicCurveSegment)pathSegment).FirstControlPoint.Y;
              dblSmoothSecondControlPointX = ((SvgCubicCurveSegment)pathSegment).SecondControlPoint.X;
              dblSmoothSecondControlPointY = ((SvgCubicCurveSegment)pathSegment).SecondControlPoint.Y;
              dblSmoothEndPointX = ((PointF)pathSegment.End).X;
              dblSmoothEndPointY = ((PointF)pathSegment.End).Y;
              dblSmoothFirstControlPointX = visPage.Application.ConvertResult(dblSmoothFirstControlPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblSmoothFirstControlPointY = -visPage.Application.ConvertResult(dblSmoothFirstControlPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblSmoothSecondControlPointX = visPage.Application.ConvertResult(dblSmoothSecondControlPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblSmoothSecondControlPointY = -visPage.Application.ConvertResult(dblSmoothSecondControlPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              arSmoothControlPoint = new double[4];
              arSmoothControlPoint[0] = (dblSmoothFirstControlPointX) / dblWidthRatio;
              arSmoothControlPoint[1] = (dblSmoothFirstControlPointY + dblOriginY + (dblBeginY * 0.5)) / dblHeightRatio;
              arSmoothControlPoint[2] = (dblSmoothFirstControlPointX) / dblWidthRatio;
              arSmoothControlPoint[3] = (dblSmoothFirstControlPointY + dblOriginY) / dblHeightRatio;
              }
            Visio.Shape visShapeSmoothBezier = visPage.DrawBezier(arSmoothControlPoint, 3, (int)VisDrawSplineFlags.visSpline1D);
            visShapeSmoothBezier.UpdateAlignmentBox();
            visShapeSmoothBezier.BoundingBox((int)VisBoundingBoxArgs.visBBoxExtents, out double dblSmoothLeft, out double dblSmoothBottom, out double dblSmoothRight, out double dblSmoothTop);
            // Re-centering relative to the SVG shape
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSmoothSVGPinXValue1);
            VisualVisioUtil.GetDoubleCellVal(visShapeSmoothBezier, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSmoothPinXValue1);
            VisualVisioUtil.SetDoubleCellVal(visShapeSmoothBezier, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblSmoothPinXValue1 + dblSmoothSVGPinXValue1) - (dblSVGWidth * 0.5));
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSmoothSVGPinYValue1);
            VisualVisioUtil.GetDoubleCellVal(visShapeSmoothBezier, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSmoothPinYValue1);
            VisualVisioUtil.SetDoubleCellVal(visShapeSmoothBezier, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblSmoothPinYValue1 + dblSmoothSVGPinYValue1) + (dblSVGHeight * 0.5));
            ApplyShapeStyles(visPage, visShapeSmoothBezier, strStrokeWidth, "", strStrokeColor, strFill, strOpacity, dblWidthRatio);
            bSmoothCubicBezier = true;
            break;
          case "s":
            break;
          case "Q":
            // Quadratic Bezier curve
            // Current point
            double dblQuadStartX = dblCurrentPointX;
            double dblQuadStartY = dblCurrentPointY;
            // Control point
            double dblControlPointX = ((SvgQuadraticCurveSegment)pathSegment).ControlPoint.X;
            double dblControlPointY = ((SvgQuadraticCurveSegment)pathSegment).ControlPoint.Y;
            dblReflexionPointX = dblCurrentPointX;
            dblReflexionPointY = dblCurrentPointY - dblControlPointY;
            // End point
            double dblEndPointX = ((SvgQuadraticCurveSegment)pathSegment).End.X;
            double dblEndPointY = ((SvgQuadraticCurveSegment)pathSegment).End.Y;
            dblCurrentPointX = dblEndPointX;
            dblCurrentPointY = dblEndPointY;
            dblQuadStartX = visPage.Application.ConvertResult(dblQuadStartX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblQuadStartY = -visPage.Application.ConvertResult(dblQuadStartY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblControlPointX = visPage.Application.ConvertResult(dblControlPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblControlPointY = -visPage.Application.ConvertResult(dblControlPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblEndPointX = visPage.Application.ConvertResult(dblEndPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblEndPointY = -visPage.Application.ConvertResult(dblEndPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            double[] arQuadraticControlPoint = new double[6];
            // Current point
            arQuadraticControlPoint[0] = dblQuadStartX / dblWidthRatio;
            arQuadraticControlPoint[1] = dblQuadStartY / dblHeightRatio;
            // Control point
            arQuadraticControlPoint[2] = (dblControlPointX) / dblWidthRatio;
            arQuadraticControlPoint[3] = (dblControlPointY) / dblHeightRatio;
            // End point
            arQuadraticControlPoint[4] = (dblEndPointX) / dblWidthRatio;
            arQuadraticControlPoint[5] = (dblEndPointY) / dblHeightRatio;
            Visio.Shape visQuadraticShapeBezier = visPage.DrawBezier(arQuadraticControlPoint, 2, (int)VisDrawSplineFlags.visSpline1D);
            visQuadraticShapeBezier.UpdateAlignmentBox();
            // Re-centering relative to the SVG shape
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblQuadraticSVGPinXValue);
            VisualVisioUtil.GetDoubleCellVal(visQuadraticShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblQuadraticPinXValue);
            VisualVisioUtil.SetDoubleCellVal(visQuadraticShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblQuadraticPinXValue + dblQuadraticSVGPinXValue) - (dblSVGWidth * 0.5));
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblQuadraticSVGPinYValue);
            VisualVisioUtil.GetDoubleCellVal(visQuadraticShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblQuadraticPinYValue);
            VisualVisioUtil.SetDoubleCellVal(visQuadraticShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblQuadraticPinYValue + dblQuadraticSVGPinYValue) + (dblSVGHeight * 0.5));
            ApplyShapeStyles(visPage, visQuadraticShapeBezier, strStrokeWidth, "", strStrokeColor, strFill, strOpacity, dblWidthRatio);
            bCubicBezier = true;
            break;
          case "q":
            break;
          case "T":
            // Quadratic Bezier curve
            // Current point
            double dblSmoothQuadStartX = dblCurrentPointX;
            double dblSmoothQuadStartY = dblCurrentPointY;
            // The control point is the reflection on the previous control point
            double dblSmoothControlPointX = dblCurrentPointX + dblReflexionPointX;
            double dblSmoothControlPointY = dblCurrentPointY + dblReflexionPointY;
            // End point
            double dblSmoothQuadEndPointX = ((SvgQuadraticCurveSegment)pathSegment).End.X;
            double dblSmoothQuadEndPointY = ((SvgQuadraticCurveSegment)pathSegment).End.Y;
            dblCurrentPointX = dblCubicEndPointX;
            dblCurrentPointY = dblCubicEndPointY;
            dblReflexionPointX = dblSmoothControlPointX;
            dblReflexionPointY = dblSmoothControlPointY;
            dblSmoothQuadStartX = visPage.Application.ConvertResult(dblSmoothQuadStartX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblSmoothQuadStartY = -visPage.Application.ConvertResult(dblSmoothQuadStartY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblSmoothControlPointX = visPage.Application.ConvertResult(dblSmoothControlPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblSmoothControlPointY = -visPage.Application.ConvertResult(dblSmoothControlPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblSmoothQuadEndPointX = visPage.Application.ConvertResult(dblSmoothQuadEndPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblSmoothQuadEndPointY = -visPage.Application.ConvertResult(dblSmoothQuadEndPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            double[] arSmoothQuadraticControlPoint = new double[6];
            // Current point
            arSmoothQuadraticControlPoint[0] = dblSmoothQuadStartX / dblWidthRatio;
            arSmoothQuadraticControlPoint[1] = dblSmoothQuadStartY / dblHeightRatio;
            // Control point
            arSmoothQuadraticControlPoint[2] = (dblSmoothControlPointX) / dblWidthRatio;
            arSmoothQuadraticControlPoint[3] = (dblSmoothControlPointY) / dblHeightRatio;
            // End point
            arSmoothQuadraticControlPoint[4] = (dblSmoothQuadEndPointX) / dblWidthRatio;
            arSmoothQuadraticControlPoint[5] = (dblSmoothQuadEndPointY) / dblHeightRatio;
            Visio.Shape visSmoothQuadraticShapeBezier = visPage.DrawBezier(arSmoothQuadraticControlPoint, 2, (int)VisDrawSplineFlags.visSpline1D);
            visSmoothQuadraticShapeBezier.UpdateAlignmentBox();
            // Re-centering relative to the SVG shape
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSmoothQuadraticSVGPinXValue);
            VisualVisioUtil.GetDoubleCellVal(visSmoothQuadraticShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSmoothQuadraticPinXValue);
            VisualVisioUtil.SetDoubleCellVal(visSmoothQuadraticShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblSmoothQuadraticPinXValue + dblSmoothQuadraticSVGPinXValue) - (dblSVGWidth * 0.5));
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSmoothQuadraticSVGPinYValue);
            VisualVisioUtil.GetDoubleCellVal(visSmoothQuadraticShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSmoothQuadraticPinYValue);
            VisualVisioUtil.SetDoubleCellVal(visSmoothQuadraticShapeBezier, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblSmoothQuadraticPinYValue + dblSmoothQuadraticSVGPinYValue) + (dblSVGHeight * 0.5));
            ApplyShapeStyles(visPage, visSmoothQuadraticShapeBezier, strStrokeWidth, "", strStrokeColor, strFill, strOpacity, dblWidthRatio);
            bCubicBezier = true;
            break;
          case "t":
            break;
          case "A":
            // elliptical arc
            break;
          case "a":
              double dblEllipticArcControlPointX = 0.0, dblEllipticArcControlPointY = 0.0;
              double dblEllipticalArcRadiusX = ((SvgArcSegment)pathSegment).RadiusX;
              double dblEllipticalArcRadiusY = ((SvgArcSegment)pathSegment).RadiusY;
              double dblAngle = ((SvgArcSegment)pathSegment).Angle;
              dblAngle = dblAngle * (Math.PI / 180.0);
              double dblEllipticalArcEndX = dblCurrentPointX + ((SvgArcSegment)pathSegment).End.X;
              double dblEllipticalArcEndY = dblCurrentPointY + ((SvgArcSegment)pathSegment).End.Y;
              SvgArcSize arcSize = ((SvgArcSegment)pathSegment).Size;
              SvgArcSweep arcSweep = ((SvgArcSegment)pathSegment).Sweep;
              // LineTo
              if (iGeometryLine > 0)
                {
                // Create a new geometry line
                visPathShape.AddRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (int)Visio.VisRowIndices.visRowLast, (int)Visio.VisRowTags.visTagEllipticalArcTo);
                }
              double dblEllipticArcEndingVertexX = dblEllipticalArcEndX;
              double dblEllipticArcEndingVertexY = dblEllipticalArcEndY;
              if (arcSize == SvgArcSize.Large)
                {
                // Large arc
                if (arcSweep == SvgArcSweep.Positive)
                  {
                  }
                else
                  {
                  dblEllipticArcControlPointX = dblCurrentPointX + ((SvgArcSegment)pathSegment).RadiusX;
                  dblEllipticArcControlPointY = dblCurrentPointY + ((SvgArcSegment)pathSegment).RadiusY;
                  }
                }
              else
                {
                // Small arc
                if (arcSweep == SvgArcSweep.Positive)
                  {
                  dblEllipticArcControlPointX = dblCurrentPointX - ((SvgArcSegment)pathSegment).RadiusX;
                  dblEllipticArcControlPointY = (dblCurrentPointY - ((SvgArcSegment)pathSegment).RadiusY);

                  }
                else
                  {
                  dblEllipticArcControlPointX = (dblArcRelOriginX - ((SvgArcSegment)pathSegment).RadiusX) - ((SvgArcSegment)pathSegment).RadiusX;
                  dblEllipticArcControlPointY = -((SvgArcSegment)pathSegment).RadiusY;
                  }
                }
              dblCurrentPointX = dblEllipticalArcEndX;
              dblCurrentPointY = dblEllipticalArcEndY;
              dblArcRelOriginX = 0;
              dblRelOriginX = 0;
              dblRelOriginY = 0;
              double dblEllipticArcAngle = dblAngle;
              dblEllipticArcEndingVertexX = visPage.Application.ConvertResult(dblEllipticArcEndingVertexX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblEllipticArcEndingVertexY = -visPage.Application.ConvertResult(dblEllipticArcEndingVertexY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblEllipticArcControlPointX = visPage.Application.ConvertResult(dblEllipticArcControlPointX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              dblEllipticArcControlPointY = -visPage.Application.ConvertResult(dblEllipticArcControlPointY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                            (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visX, (dblEllipticArcEndingVertexX) / dblWidthRatio);
              VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                            (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visY, (dblEllipticArcEndingVertexY) / dblHeightRatio);
              VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                            (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visControlX, (dblEllipticArcControlPointX) / dblWidthRatio);
              VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                            (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visControlY, (dblEllipticArcControlPointY) / dblHeightRatio);
              VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                            (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visEccentricityAngle, dblEllipticArcAngle);
            iGeometryLine++;
            break;
          case "H":
            // Horizontal lineto
            double dblAbsHorizontalX = pathSegment.End.X;
            break;
          case "h":
            double dblRelHorizontalX = (dblCurrentPointX - dblRelOriginX) + pathSegment.End.X;
            double dblRelHorizontalY = dblCurrentPointY - dblRelOriginY;
            if (iGeometryLine > 0)
              {
              // Create a new geometry line
              visPathShape.AddRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (int)Visio.VisRowIndices.visRowLast, (int)Visio.VisRowTags.visTagLineTo);
              }
            dblCurrentPointX = dblRelHorizontalX;
            dblCurrentPointY = dblRelHorizontalY;
            dblRelOriginX = 0.0;
            dblRelOriginY = 0.0;
            InsertMidMarker(visPage, visSVGShape,strStrokeColor, dblSVGWidth, dblSVGHeight, visMidMarkerShape, iBeginVertices, iCountBeginVertices,
                                 dblPathOriginX + dblRelHorizontalX, dblPathOriginY + dblRelHorizontalY);
            dblRelHorizontalX = visPage.Application.ConvertResult(dblRelHorizontalX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblRelHorizontalY = -visPage.Application.ConvertResult(dblRelHorizontalY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                        (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visX, (dblRelHorizontalX) / dblWidthRatio);
            VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                          (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visY, (dblRelHorizontalY) / dblHeightRatio);
            //}
            iGeometryLine++;
            break;
          case "v":
            double dblRelVerticalX = (dblCurrentPointX - dblRelOriginX);
            double dblRelVerticalY = (dblCurrentPointY - dblRelOriginY) + pathSegment.End.Y;
            if (iGeometryLine > 0)
              {
              // Create a new geometry line
              visPathShape.AddRow((short)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry), (int)Visio.VisRowIndices.visRowLast, (int)Visio.VisRowTags.visTagLineTo);
              }
            dblCurrentPointX = dblRelVerticalX;
            dblCurrentPointY = dblRelVerticalY;
            dblRelOriginX = 0.0;
            dblRelOriginY = 0.0;
            dblRelVerticalX = visPage.Application.ConvertResult(dblRelVerticalX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblRelVerticalY = -visPage.Application.ConvertResult(dblRelVerticalY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                        (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visX, (dblRelVerticalX) / dblWidthRatio);
            VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)(Visio.VisSectionIndices.visSectionFirstComponent + iGeometry),
                          (int)Visio.VisRowIndices.visRowVertex + iGeometryLine, (int)Visio.VisCellIndices.visY, (dblRelVerticalY) / dblHeightRatio);
            //  }
            //InsertMidMarker(visPage, visSVGShape, strStrokeColor, dblSVGWidth, dblSVGHeight, visMidMarkerShape, iBeginVertices, iCountBeginVertices,
            //                     dblPathOriginX + dblRelVerticalX, dblPathOriginY + dblRelVerticalY);
            InsertMidMarker(visPage, visSVGShape, strStrokeColor, dblSVGWidth, dblSVGHeight, visMidMarkerShape, iBeginVertices, iCountBeginVertices,
                                 dblPathOriginX + dblRelVerticalX + dblCurrentPointX, dblPathOriginY + dblRelVerticalY + dblCurrentPointY);
            //}
            iGeometryLine++;
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
            bPathZEndFound = true;
            break;
          }
        }
      visPathShape.UpdateAlignmentBox();
      // Re-centering relative to the SVG shape
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
      VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblPinXValue + dblSVGPinXValue) - (dblSVGWidth * 0.5));
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      VisualVisioUtil.GetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
      VisualVisioUtil.SetDoubleCellVal(visPathShape, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblPinYValue + dblSVGPinYValue) + (dblSVGHeight * 0.5));
      ApplyShapeStyles(visPage, visPathShape, strStrokeWidth, "", strStrokeColor, strFill, strOpacity, dblWidthRatio);
      if (bHide)
        {
        VisualVisioUtil.SetGeometryVisibility(visPathShape, false, false);
        }
      }

    public static void CreateLine(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, string strStrokeColor, string strStrokeWidth, string strFill,
                                  string strOpacity, double dblWidthRatio, double dblHeightRatio, double dblSVGWidth, double dblSVGHeight)
      {
      string strParamStrokeColor = "", strParamLocStrokeColor = "";
      string strParamStrokeWidth = "", strParamLocStrokeWidth = "";
      string strParamFill = "", strParamLocFill = "";
      string strParamOpacity = "", strParamLocOpacity = "";

      ((SvgLine)element).TryGetAttribute("stroke", out strParamLocStrokeColor);
      ((SvgLine)element).TryGetAttribute("stroke-width", out strParamLocStrokeWidth);
      ((SvgLine)element).TryGetAttribute("fill", out strParamLocFill);
      ((SvgLine)element).TryGetAttribute("opacity", out strParamLocOpacity);

      if (strParamLocStrokeColor != null)
        strParamStrokeColor = strParamLocStrokeColor;
      if (strParamLocStrokeWidth != null)
        strParamStrokeWidth = strParamLocStrokeWidth;
      if (strParamLocFill != null)
        strParamFill = strParamLocFill;
      if (strParamLocOpacity != null)
        strParamOpacity = strParamLocOpacity;
      if (strStrokeColor != "")
        {
        strParamStrokeColor = strStrokeColor;
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
      if (strOpacity != "")
        {
        strParamOpacity = strOpacity;
        if ((strParamLocOpacity != "") && (strParamLocOpacity != null))
          {
          strParamOpacity = strParamLocOpacity;
          }
        }
      double dblSubBeginX = (((SvgLine)element).StartX).Value;
      double dblSubBeginY = (((SvgLine)element).StartY).Value;
      double dblSubEndX = (((SvgLine)element).EndX).Value;
      double dblSubEndY = (((SvgLine)element).EndY).Value;
      dblSubBeginX = visPage.Application.ConvertResult(dblSubBeginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblSubBeginY = -visPage.Application.ConvertResult(dblSubBeginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblSubEndX = visPage.Application.ConvertResult(dblSubEndX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblSubEndY = -visPage.Application.ConvertResult(dblSubEndY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      Visio.Shape visShape = visPage.DrawLine(dblSubBeginX / dblWidthRatio, dblSubBeginY / dblHeightRatio, dblSubEndX / dblWidthRatio, dblSubEndY / dblHeightRatio);
      // Re-centering relative to the SVG shape
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      // Cropping X
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DBEGINX, out double dblBeginXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DBEGINX, dblBeginXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DENDX, out double dblEndXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DENDX, dblEndXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
      // Cropping Y
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DBEGINY, out double dblBeginYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DBEGINY, dblBeginYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DENDY, out double dblEndYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_1DENDY, dblEndYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
      ApplyShapeStyles(visPage, visShape, strParamStrokeWidth, "", strParamStrokeColor, strParamFill, strParamOpacity, dblWidthRatio);
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
      Visio.Shape visShape = visPage.DrawRectangle(dblPinX / dblWidthRatio, -dblPinY / dblHeightRatio, dblPinX / dblWidthRatio, -dblPinY / dblHeightRatio);
      var iSize = visShape.get_CellsSRC((int)Visio.VisSectionIndices.visSectionCharacter, 0, (int)Visio.VisCellIndices.visCharacterSize).ResultIU;
      double dblBlocTextMargin = 4.0;
      dblBlocTextMargin = visPage.Application.ConvertResult(dblBlocTextMargin, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      iSize = (iSize + dblBlocTextMargin) / dblHeightRatio;
      visShape.Text = strText;
      double dblTxtWidth = (strText.Length - 1) * iSize;
      double dblTxtHeight = iSize * 2;
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, (int)Visio.VisUnitCodes.visInches, dblTxtWidth);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, (int)Visio.VisUnitCodes.visInches, dblTxtHeight);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, (int)Visio.VisUnitCodes.visInches, (dblPinX / dblWidthRatio) + (dblTxtWidth * 0.5));
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, (int)Visio.VisUnitCodes.visInches, (dblPinY / dblHeightRatio) - (dblTxtHeight * 0.5));
      VisualVisioUtil.SetIntCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LINEPATTERN, 0);
      VisualVisioUtil.SetIntCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLPATTERN, 0);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_CHARSIZE, (int)Visio.VisUnitCodes.visPoints, dblFontSize / dblHeightRatio);
      Visio.Fonts visFonts = visPage.Document.Fonts;
      Visio.Font visFont = visFonts[strFontFamily];
      VisualVisioUtil.SetIntCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_CHARFONT, visFont.ID);
      // Re-centering relative to the SVG shape
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblPinXValue + dblSVGPinXValue) - (dblSVGWidth * 0.5));
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblPinYValue + dblSVGPinYValue) + (dblSVGHeight * 0.5));
      }

    public static void CreateRectangleWithText(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, double dblTranslateX, double dblTranslateY, double dblAngle, double dblWidthRatio, double dblHeightRatio,
                          double dblSVGWidth, double dblSVGHeight, bool bViewBox, string strFill, string strStrokeColor, string strOpacity)
      {
      string strParamStrokeColor = "", strParamLocStrokeColor = "";
      string strParamStrokeWidth = "", strParamLocStrokeWidth = "";
      string strParamFill = "", strParamLocFill = "";
      string strParamOpacity = "", strParamLocOpacity = "";
      string strRounding = "";
      Visio.Shape visShape;

      strParamStrokeColor = strStrokeColor;
      strParamFill = strFill;
      double dblX1 = ((SvgRectangle)element).X + dblTranslateX;
      double dblY1 = ((SvgRectangle)element).Y + dblTranslateY;
      double dblX2 = dblX1 + ((SvgRectangle)element).Width;
      double dblY2 = dblY1 + ((SvgRectangle)element).Height;
      ((SvgRectangle)element).TryGetAttribute("rx", out strRounding);
      ((SvgRectangle)element).TryGetAttribute("stroke", out strParamLocStrokeColor);
      ((SvgRectangle)element).TryGetAttribute("stroke-width", out strParamLocStrokeWidth);
      ((SvgRectangle)element).TryGetAttribute("fill", out strParamLocFill);
      ((SvgRectangle)element).TryGetAttribute("opacity", out strParamLocOpacity);
      if (strParamLocStrokeColor != null)
        strParamStrokeColor = strParamLocStrokeColor;
      if (strParamLocStrokeWidth != null)
        strParamStrokeWidth = strParamLocStrokeWidth;
      if (strParamLocFill != null)
        strParamFill = strParamLocFill;
      if (strParamLocOpacity != null)
        strParamOpacity = strParamLocOpacity;
      if (strStrokeColor != "")
        {
        strParamStrokeColor = strStrokeColor;
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
      if (strOpacity != "")
        {
        strParamOpacity = strOpacity;
        if ((strParamLocOpacity != "") && (strParamLocOpacity != null))
          {
          strParamOpacity = strParamLocOpacity;
          }
        }
      dblX1 = visPage.Application.ConvertResult(dblX1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblY1 = -visPage.Application.ConvertResult(dblY1, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblX2 = visPage.Application.ConvertResult(dblX2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      dblY2 = -visPage.Application.ConvertResult(dblY2, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
      visShape = visPage.DrawRectangle(dblX1 / dblWidthRatio, dblY1 / dblHeightRatio, dblX2 / dblWidthRatio, dblY2 / dblHeightRatio);
      // Possible rotation
      if (dblAngle != 0)
        {
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXBeforeCenterRotationChanges);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYBeforeCenterRotationChanges);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_WIDTH, out double dblWidth);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_HEIGHT, out double dblHeight);
        // Rotation center shifted to the left-center
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, 0.0);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXBeforeCenterRotationChanges - (dblWidth * 0.5));
        // Rotation center at the top-left
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight);
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYBeforeCenterRotationChanges + (dblHeight * 0.5));
        // Rotating the shape. Note that the angle sign must be inverted
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_ANGLE, (int)Visio.VisUnitCodes.visDegrees, -dblAngle);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRotation);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRotation);
        // Centre de rotation au centre en haut pour commencer à revenir à la position originale
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINX, dblWidth * 0.5);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRotation + ((dblWidth * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRotation + ((dblWidth * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXAfterRepos);
        VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYAfterRepos);
        // Rotation center at the center-top to start returning to the original position
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LOCPINY, dblHeight * 0.5);
        //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out dblPinXAfterRepos);
        //VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out dblPinYAfterRepos);
        // repositioning the shape along the X-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXAfterRepos + ((dblHeight * 0.5) * Math.Sin(-dblAngle * (Math.PI / 180.0))));
        // Repositioning the shape along the Y-axis
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYAfterRepos - ((dblHeight * 0.5) * Math.Cos(-dblAngle * (Math.PI / 180.0))));
        }
      // Re-centering relative to the SVG shape
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblSVGPinXValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblPinXValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINX, dblPinXValue + (dblSVGPinXValue - (dblSVGWidth * 0.5)));
      VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblSVGPinYValue);
      VisualVisioUtil.GetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblPinYValue);
      VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_PINY, dblPinYValue + (dblSVGPinYValue + (dblSVGHeight * 0.5)));
      ApplyShapeStyles(visPage, visShape, strParamStrokeWidth, strRounding, strParamStrokeColor, strParamFill, strParamOpacity, dblWidthRatio);
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

    public static void Create2DPolylineFromMarker(Visio.Page visPage, Visio.Shape visSVGShape, SvgElement element, string styleContent, double dblWidthRatio, double dblHeightRatio,
                              double dblSVGWidth, double dblSVGHeight, bool bHide)
      {
      string strStrokeColor, strStrokeWidth = "", strFill = "", strOpacity = "", strMarkerWidth, strMarkerHeight;
      Visio.Shape visMarkerShape;
      double dblMarkerWidthRatio = 1.0, dblMarkerHeightRatio = 1.0;
      double dblMarkerWidth = 1.0, dblMarkerHeight = 1.0;

      ((SvgMarker)element).TryGetAttribute("stroke", out strStrokeColor);
      ((SvgMarker)element).TryGetAttribute("stroke-width", out strStrokeWidth);
      ((SvgMarker)element).TryGetAttribute("fill", out strFill);
      ((SvgMarker)element).TryGetAttribute("opacity", out strOpacity);
      ((SvgMarker)element).TryGetAttribute("markerWidth", out strMarkerWidth);
      dblMarkerWidth = ((SvgMarker)element).ViewBox.Width;
      dblMarkerHeight = ((SvgMarker)element).ViewBox.Height;
      if (!string.IsNullOrEmpty(strMarkerWidth) && (dblMarkerWidth != 0))
        {
        dblMarkerWidth = ((SvgMarker)element).ViewBox.Width;
        NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
        if (Double.TryParse(strMarkerWidth, NumberStyles.AllowDecimalPoint, nfi, out double dblValue))
          dblMarkerWidthRatio = dblValue / dblMarkerWidth;
        }
      ((SvgMarker)element).TryGetAttribute("markerHeight", out strMarkerHeight);
      if (!string.IsNullOrEmpty(strMarkerHeight) && (dblMarkerHeight != 0))
        {
        dblMarkerHeight = ((SvgMarker)element).ViewBox.Height;
        NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
        if (Double.TryParse(strMarkerHeight, NumberStyles.AllowDecimalPoint, nfi, out double dblValue))
          dblMarkerHeightRatio = dblValue / dblMarkerHeight;
        }
      switch (((SvgMarker)element).Children[0].GetType().Name)
        {
        case "SvgPath":
          Create2DPolylineFromPath(visPage, visSVGShape, ((SvgMarker)element).Children[0], styleContent, dblWidthRatio, dblHeightRatio, dblSVGWidth, dblSVGHeight, bHide);
          break;
        case "SvgCircle":
          visMarkerShape = CreateCircle(visPage, visSVGShape, ((SvgMarker)element).Children[0], 0.0, 0.0, 0, dblWidthRatio, dblHeightRatio,
                                    dblSVGWidth, dblSVGHeight, dblMarkerWidthRatio, dblMarkerHeightRatio, true, strFill, strStrokeColor, strOpacity);
          if (visMarkerShape != null)
            {
            visMarkerShape.Name = ((SvgMarker)element).ID;
            VisualVisioUtil.SetGeometryVisibility(visMarkerShape, false, false);
            }
          break;
        }
      }


    private static bool InsertStartMarker(Visio.Page visPage, Visio.Shape visSVGShape, string strStrokeColor, double dblSVGWidth, double dblSVGHeight,
                                           Visio.Shape visStartMarkerShape, Visio.Shape visMidMarkerShape, Visio.Shape visEndMarkerShape,
                                           int iBeginVertices,
                                      int iCountBeginVertices, double dblRelOriginX, double dblRelOriginY)
      {
      if (visStartMarkerShape != null)
        {
        if (iBeginVertices == 0)
          {
          Visio.Shape visMarker = visStartMarkerShape.Duplicate();
          VisualVisioUtil.SetGeometryVisibility(visMarker, true, true);
          double dblMarkerOriginX = visPage.Application.ConvertResult(dblRelOriginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
          double dblMarkerOriginY = -visPage.Application.ConvertResult(dblRelOriginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
          VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, dblMarkerOriginX);
          VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, dblMarkerOriginY);
          // Re-centering relative to the SVG shape
          VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblEndFoundSVGPinXValue);
          VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblEndMarkerPinXValue);
          VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblEndMarkerPinXValue + dblEndFoundSVGPinXValue) - (dblSVGWidth * 0.5));
          VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblEndFoundSVGPinYValue);
          VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblEndMarkerPinYValue);
          VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblEndMarkerPinYValue + dblEndFoundSVGPinYValue) + (dblSVGHeight * 0.5));
          if (visMidMarkerShape != null)
            {
            visMarker = visMidMarkerShape.Duplicate();
            VisualVisioUtil.SetGeometryVisibility(visMarker, true, true);
            dblMarkerOriginX = visPage.Application.ConvertResult(dblRelOriginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblMarkerOriginY = -visPage.Application.ConvertResult(dblRelOriginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, dblMarkerOriginX);
            VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, dblMarkerOriginY);
            // Re-centering relative to the SVG shape
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblMidFoundSVGPinXValue);
            VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblMidMarkerPinXValue);
            VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblMidMarkerPinXValue + dblMidFoundSVGPinXValue) - (dblSVGWidth * 0.5));
            VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblMidFoundSVGPinYValue);
            VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblMidMarkerPinYValue);
            VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblMidMarkerPinYValue + dblMidFoundSVGPinYValue) + (dblSVGHeight * 0.5));
            }
          }
        else
          {
          if (iBeginVertices == iCountBeginVertices - 1)
            {
            if (visMidMarkerShape != null)
              {
              Visio.Shape visMarker = visMidMarkerShape.Duplicate();
              VisualVisioUtil.SetGeometryVisibility(visMarker, true, true);
              double dblMarkerOriginX = visPage.Application.ConvertResult(dblRelOriginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              double dblMarkerOriginY = -visPage.Application.ConvertResult(dblRelOriginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, dblMarkerOriginX);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, dblMarkerOriginY);
              // Re-centering relative to the SVG shape
              VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblMidFoundSVGPinXValue);
              VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblMidMarkerPinXValue);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblMidMarkerPinXValue + dblMidFoundSVGPinXValue) - (dblSVGWidth * 0.5));
              VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblMidFoundSVGPinYValue);
              VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblMidMarkerPinYValue);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblMidMarkerPinYValue + dblMidFoundSVGPinYValue) + (dblSVGHeight * 0.5));
              }
            if (visEndMarkerShape != null)
              {
              Visio.Shape visMarker = visEndMarkerShape.Duplicate();
              VisualVisioUtil.SetGeometryVisibility(visMarker, true, true);
              double dblMarkerOriginX = visPage.Application.ConvertResult(dblRelOriginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              double dblMarkerOriginY = -visPage.Application.ConvertResult(dblRelOriginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, dblMarkerOriginX);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, dblMarkerOriginY);
              // Re-centering relative to the SVG shape
              VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblEndFoundSVGPinXValue);
              VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblEndMarkerPinXValue);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblEndMarkerPinXValue + dblEndFoundSVGPinXValue) - (dblSVGWidth * 0.5));
              VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblEndFoundSVGPinYValue);
              VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblEndMarkerPinYValue);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblEndMarkerPinYValue + dblEndFoundSVGPinYValue) + (dblSVGHeight * 0.5));
              }
            }
          else
            {
            if (visMidMarkerShape != null)
              {
              Visio.Shape visMarker = visMidMarkerShape.Duplicate();
              VisualVisioUtil.SetGeometryVisibility(visMarker, true, true);
              double dblMarkerOriginX = visPage.Application.ConvertResult(dblRelOriginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              double dblMarkerOriginY = -visPage.Application.ConvertResult(dblRelOriginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, dblMarkerOriginX);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, dblMarkerOriginY);
              // Re-centering relative to the SVG shape
              VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblMidFoundSVGPinXValue);
              VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblMidMarkerPinXValue);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblMidMarkerPinXValue + dblMidFoundSVGPinXValue) - (dblSVGWidth * 0.5));
              VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblMidFoundSVGPinYValue);
              VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblMidMarkerPinYValue);
              VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblMidMarkerPinYValue + dblMidFoundSVGPinYValue) + (dblSVGHeight * 0.5));
              }
            }
          }
        return true;
        }
      else
        {
        if (visMidMarkerShape != null)
          {
          Visio.Shape visMarker = visMidMarkerShape.Duplicate();
          if (strStrokeColor != "")
            ApplyShapeStyles(visPage, visMarker, "", "", strStrokeColor, "", "", 1.0);
          VisualVisioUtil.SetGeometryVisibility(visMarker, true, true);
          double dblMarkerOriginX = visPage.Application.ConvertResult(dblRelOriginX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
          double dblMarkerOriginY = -visPage.Application.ConvertResult(dblRelOriginY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
          VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, dblMarkerOriginX);
          VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, dblMarkerOriginY);
          // Re-centering relative to the SVG shape
          VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblMidFoundSVGPinXValue);
          VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblMidMarkerPinXValue);
          VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblMidMarkerPinXValue + dblMidFoundSVGPinXValue) - (dblSVGWidth * 0.5));
          VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblMidFoundSVGPinYValue);
          VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblMidMarkerPinYValue);
          VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblMidMarkerPinYValue + dblMidFoundSVGPinYValue) + (dblSVGHeight * 0.5));
          }
        return false;
        }
      }

    private static void InsertMidMarker(Visio.Page visPage, Visio.Shape visSVGShape,string strStrokeColor, double dblSVGWidth, double dblSVGHeight, Visio.Shape visMidMarkerShape, int iBeginVertices,
                                      int iCountBeginVertices, double dblCoordX, double dblCoordY)
      {
      if (visMidMarkerShape != null)
        {
        Visio.Shape visMarker = visMidMarkerShape.Duplicate();
        if(strStrokeColor != "")
          ApplyShapeStyles(visPage, visMarker, "", "", strStrokeColor, "", "", 1.0);
        VisualVisioUtil.SetGeometryVisibility(visMarker, true, true);
        double dblMarkerOriginX = visPage.Application.ConvertResult(dblCoordX, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        double dblMarkerOriginY = -visPage.Application.ConvertResult(dblCoordY, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, dblMarkerOriginX);
        VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, dblMarkerOriginY);
        // Re-centering relative to the SVG shape
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblMidFoundSVGPinXValue);
        VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, out double dblMidMarkerPinXValue);
        VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINX, (dblMidMarkerPinXValue + dblMidFoundSVGPinXValue) - (dblSVGWidth * 0.5));
        VisualVisioUtil.GetDoubleCellVal(visSVGShape, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblMidFoundSVGPinYValue);
        VisualVisioUtil.GetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, out double dblMidMarkerPinYValue);
        VisualVisioUtil.SetDoubleCellVal(visMarker, (int)VLConstants.SRCValue.ID_SRC_PINY, (dblMidMarkerPinYValue + dblMidFoundSVGPinYValue) + (dblSVGHeight * 0.5));
        }

      }

    private static void ApplyShapeStyles(Visio.Page visPage, Visio.Shape visShape, string strParamStrokeWidth, string strRounding, string strParamStrokeColor, string strParamFill, string strParamOpacity, double dblWidthRatio)
      {
      if (!string.IsNullOrEmpty(strParamStrokeWidth))
        {
        Double.TryParse(strParamStrokeWidth, out double dblLineWeight);
        dblLineWeight = dblLineWeight / dblWidthRatio;
        NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
        VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_LINEWEIGHT, dblLineWeight.ToString("0.00 pt", nfi));
        }
      if (!string.IsNullOrEmpty(strRounding))
        {
        Double.TryParse(strRounding, out double dblRounding);
        dblRounding = visPage.Application.ConvertResult(dblRounding, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
        dblRounding = dblRounding / dblWidthRatio;
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_LINEROUNDING, dblRounding);
        }
      switch (strParamStrokeColor)
        {
        case "":
        case null:
          break;
        case "None":
        case "none":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_LINEPATTERN, "0");
          break;
        default:
          if (strParamStrokeColor.StartsWith("#") && strParamStrokeColor.Length == 7)
            {
            int r = int.Parse(strParamStrokeColor.Substring(1, 2), System.Globalization.NumberStyles.HexNumber);
            int g = int.Parse(strParamStrokeColor.Substring(3, 2), System.Globalization.NumberStyles.HexNumber);
            int b = int.Parse(strParamStrokeColor.Substring(5, 2), System.Globalization.NumberStyles.HexNumber);
            VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, $"RGB({r},{g},{b})");
            }
          else
            {
            System.Drawing.Color color = System.Drawing.Color.FromName(strParamStrokeColor);
            string strColor = "RGB(" + color.R.ToString() + "," + color.G.ToString() + "," + color.B.ToString() + ")";
            VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_LINECOLOR, strColor);
            }
          break;
        }
      switch (strParamFill)
        {
        case "":
        case null:
          break;
        case "None":
        case "none":
          VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLPATTERN, "0");
          break;
        default:
          if (strParamFill.StartsWith("#") && strParamFill.Length == 7)
            {
            int r = int.Parse(strParamFill.Substring(1, 2), System.Globalization.NumberStyles.HexNumber);
            int g = int.Parse(strParamFill.Substring(3, 2), System.Globalization.NumberStyles.HexNumber);
            int b = int.Parse(strParamFill.Substring(5, 2), System.Globalization.NumberStyles.HexNumber);
            VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, $"RGB({r},{g},{b})");
            }
          else
            {
            System.Drawing.Color color = System.Drawing.Color.FromName(strParamFill);
            string strColor = "RGB(" + color.R.ToString() + "," + color.G.ToString() + "," + color.B.ToString() + ")";
            VisualVisioUtil.SetFormulaCell(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNB, strColor);
            }
          break;
        }
      if (!string.IsNullOrEmpty(strParamOpacity))
        {
        Double.TryParse(strParamOpacity, out double dblOpacity);
        dblOpacity = (1 - dblOpacity) * 100.0;
        VisualVisioUtil.SetDoubleCellVal(visShape, (int)VLConstants.SRCValue.ID_SRC_FILLFOREGNBTRANS, (int)Visio.VisUnitCodes.visPercent, dblOpacity);
        }
      }




    }
  }
