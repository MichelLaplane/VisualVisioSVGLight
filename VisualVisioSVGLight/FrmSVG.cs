using Microsoft.Web.WebView2.Core;
using Svg;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Svg;
using static System.Windows.Forms.AxHost;
using Visio = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Interop.Visio;
using System.Xml.Linq;
using System.Web.UI.HtmlControls;
using System.Drawing.Drawing2D;
using Svg.Transforms;

namespace VisualVisioSVGLight
  {
  public partial class FrmSVG : Form
    {
    Microsoft.Office.Interop.Visio.Application visApp;
    private static string mermaidHtmlFileName = "VisualMermaidVisio.html";
    private static string svgFileName = "VisualVisioSVG.svg";
    private static string pngFileName = "VisualVisioSVG.png";

    public FrmSVG(Microsoft.Office.Interop.Visio.Application visParamApp)
      {
      InitializeComponent();
      visApp = visParamApp;
      }

    private void btnOpenPng_Click(object sender, EventArgs e)
      {
      OpenFileDialog openFileDialog1 = new OpenFileDialog();
      if (openFileDialog1.ShowDialog() == DialogResult.OK)
        {
        string strFullPath = openFileDialog1.FileName;
        edSVG.Text = System.IO.File.ReadAllText(strFullPath);
        }
      }

    private void edSVG_TextChanged(object sender, EventArgs e)
      {
      if (edSVG.Text.Split('\n').Length > 15)
        edSVG.ScrollBars = ScrollBars.Vertical;
      else
        edSVG.ScrollBars = ScrollBars.None;
      }
    private void btnVisioPngInsert_Click(object sender, EventArgs e)
      {
      string strFullPath;
      Microsoft.Office.Interop.Visio.Page visActivePage = visApp.ActivePage;
      if (visActivePage == null)
        { MessageBox.Show("No active page in Visio document");
        return;
        }
      strFullPath = System.IO.Path.Combine(VisualVisioSVGLight.strProjectPath, svgFileName);
      System.IO.File.WriteAllText(strFullPath, edSVG.Text);
      var svgDocument = SvgDocument.Open(strFullPath);
      var pngImage = svgDocument.Draw();
      pngImage.Save(System.IO.Path.Combine(VisualVisioSVGLight.strProjectPath, pngFileName));
      // Import in Visio document
      visActivePage.Import(System.IO.Path.Combine(VisualVisioSVGLight.strProjectPath, pngFileName));
      }

    /// <summary>
    /// Insert SVG elements into Visio
    /// https://www.w3.org/TR/SVG2/coords.html#Introduction
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void btnSvgNativeInsert_Click(object sender, EventArgs e)
      {
      string strFullPath = System.IO.Path.Combine(VisualVisioSVGLight.strProjectPath, svgFileName);
      double dblViewBoxX = 0.0;
      double dblViewBoxY = 0.0;
      double dblViewBoxWidth = 0.0;
      double dblViewBoxHeight = 0.0;
      bool bViewBox = false;
      double dblWidthRatio = 0.0;
      double dblHeightRatio = 0.0;
      double dblSVGInchesWidth = 0.0, dblSVGInchesHeight = 0.0;
      string strWidthUnit="", strHeightUnit="",strWidth, strHeight;
      double dblSVGWidth=0.0,dblSVGHeight=0.0;
      float fltAngle = 0.0F, fltX = 0.0F, fltY = 0.0F;
      string strStrokeColor = "";
      string strFill = "";

      var svgDocument = SvgDocument.Open(strFullPath);
      Microsoft.Office.Interop.Visio.Page visActivePage = visApp.ActivePage;
      if (visActivePage != null)
        {
        svgDocument.TryGetAttribute("width", out string strSvgWidth);
        svgDocument.TryGetAttribute("height", out string strSvgHeight);
        if (strSvgWidth != null)
          {
          if (strSvgWidth.EndsWith("%"))
            {
            strSvgWidth = "254px";
            strWidthUnit = strSvgWidth.Remove(0, (strSvgWidth.Length - 2));
            }
          else
            {
            strWidthUnit = strSvgWidth.Remove(0, (strSvgWidth.Length - 2));
            }
          }
        if (strSvgHeight != null)
          {
          if (strSvgHeight.EndsWith("%"))
            {
            strSvgHeight = "254px";
            strHeightUnit = strSvgHeight.Remove(0, (strSvgHeight.Length - 2));
            }
          else
            {
            strHeightUnit = strSvgHeight.Remove(0, (strSvgHeight.Length - 2));
            }
          }
        string strSvgUnit = strWidthUnit;
        switch (strWidthUnit)
          {
          case "px":
            strSvgUnit = "px";
            break;
          case "cm":
            strSvgUnit = "cm";
            break;
          }
        if ((strSvgWidth != "") && (strSvgWidth != null))
          {
          strWidth = strSvgWidth.Replace(strWidthUnit, "");
          dblSVGWidth = Convert.ToDouble(strWidth);
          //if(strWidthUnit == "%")
          //  dblSVGWidth *= 5;
          }
        if ((strSvgHeight != "") && (strSvgHeight != null))
          {
          strHeight = strSvgHeight.Replace(strWidthUnit, "");
          dblSVGHeight = Convert.ToDouble(strHeight);
          //if (strHeightUnit == "%")
          //  dblSVGHeight *= 5;
          }
        if(dblSVGHeight == 0.0)
          dblSVGHeight = dblSVGWidth;
        // Rectangle du SVG
        switch (strWidthUnit)
          {
          case "px":
            strSvgUnit = "px";
            dblSVGInchesWidth = visActivePage.Application.ConvertResult(dblSVGWidth, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            dblSVGInchesHeight = visActivePage.Application.ConvertResult(dblSVGHeight, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches);
            break;
          case "cm":
            strSvgUnit = "cm";
            dblSVGInchesWidth = visActivePage.Application.ConvertResult(dblSVGWidth, (int)Visio.VisUnitCodes.visCentimeters, (int)Visio.VisUnitCodes.visInches);
            dblSVGInchesHeight = visActivePage.Application.ConvertResult(dblSVGHeight, (int)Visio.VisUnitCodes.visCentimeters, (int)Visio.VisUnitCodes.visInches);
            break;
          case "%":
            strSvgUnit = "px";
            dblSVGInchesWidth = visActivePage.Application.ConvertResult(dblSVGWidth, (int)Visio.VisUnitCodes.visCentimeters, (int)Visio.VisUnitCodes.visInches);
            dblSVGInchesHeight = visActivePage.Application.ConvertResult(dblSVGHeight, (int)Visio.VisUnitCodes.visCentimeters, (int)Visio.VisUnitCodes.visInches);
            break;
          }

        dblWidthRatio = visActivePage.Application.ConvertResult(dblSVGWidth, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches) / dblSVGInchesWidth;
        dblHeightRatio = visActivePage.Application.ConvertResult(dblSVGHeight, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches) / dblSVGInchesHeight;
        svgDocument.TryGetAttribute("viewBox", out string strViewbox);
        if (strViewbox == "Svg.SvgViewBox")
          {
          bViewBox = true;
          SvgViewBox svgViewBox = svgDocument.ViewBox;
          dblViewBoxX = svgViewBox.MinX;
          dblViewBoxY = svgViewBox.MinY;
          dblViewBoxWidth = svgViewBox.Width;
          dblViewBoxHeight = svgViewBox.Height;
          dblWidthRatio = visActivePage.Application.ConvertResult(dblViewBoxWidth, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches) / dblSVGInchesWidth;
          dblHeightRatio = visActivePage.Application.ConvertResult(dblViewBoxHeight, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches) / dblSVGInchesHeight;
          }


        //visApp.ConvertResult(strWidth, strWidth.Remove(0, (strWidth.Length - 2)),"mm");
        //

        // Valeur de dblPageWidth et dblPageHeight en pouces (inches)
        double dblPageWidth = visActivePage.PageSheet.get_CellsSRC((int)Visio.VisSectionIndices.visSectionObject,
            (int)Visio.VisRowIndices.visRowPage,
            (int)Visio.VisCellIndices.visPageWidth).ResultIU;
        double dblPageHeight = visActivePage.PageSheet.get_CellsSRC((int)Visio.VisSectionIndices.visSectionObject,
                  (int)Visio.VisRowIndices.visRowPage,
                  (int)Visio.VisCellIndices.visPageHeight).ResultIU;

        //double dblSVGWidthRatio = visActivePage.Application.ConvertResult(dblSVGWidth, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches) / dblSVGInchesWidth;
        //double dblSVGHeightRatio = visActivePage.Application.ConvertResult(dblSVGHeight, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches) / dblSVGInchesHeight;
        Visio.Shape visSVGShape = visActivePage.DrawRectangle(0, 0, dblSVGInchesWidth, -dblSVGInchesHeight);
        // centrage du dessin
        visActivePage.CenterDrawing();
        //double dblWidthRatio = visActivePage.Application.ConvertResult(dblViewBoxWidth, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches) / dblSVGInchesWidth;
        //double dblHeightRatio = visActivePage.Application.ConvertResult(dblViewBoxHeight, (int)Visio.VisUnitCodes.visPoints, (int)Visio.VisUnitCodes.visInches) / dblSVGInchesHeight;

        // Access SVG elements

        foreach (SvgElement element in svgDocument.Children)
          {
          // Perform actions on each element
          var symbol = element.GetType();
          switch (symbol.Name)
            {
            case "SvgTitle":
              break;
            case "SvgDescription":
              break;
            case "SvgRectangle":
              VisualVisioSVGLightUtil.CreateRect(visActivePage, visSVGShape, element, fltX, fltY, fltAngle, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight, bViewBox, strFill, strStrokeColor);
              break;
            case "SvgCircle":
              VisualVisioSVGLightUtil.CreateCircle(visActivePage, visSVGShape, element, fltX, fltY, fltAngle, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight, bViewBox, strFill, strStrokeColor);
              break;
            case "SvgLine":
              double dblBeginX = ((SvgLine)element).StartX * ((1 / dblSVGWidth) * 100);
              double dblBeginY = ((SvgLine)element).StartY * ((1 / dblSVGHeight) * 100);
              double dblEndX = ((SvgLine)element).EndX * ((1 / dblSVGWidth) * 100);
              double dblEndY = ((SvgLine)element).EndY * ((1 / dblSVGHeight) * 100);
              visActivePage.DrawLine(dblBeginX, dblBeginY, dblEndX, dblEndY);
              break;
            case "SvgPolyline":
              VisualVisioSVGLightUtil.CreatePolyline(visActivePage, visSVGShape, element, fltX, fltY, fltAngle, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight, bViewBox, strFill, strStrokeColor);
              break;
            case "SvgPath":
              SvgPath svgPath = ((SvgPath)element);
              Svg.Pathing.SvgPathSegmentList arData = svgPath.PathData;
              VisualVisioSVGLightUtil.Create2DPolylineFromPath(visActivePage, visSVGShape, element, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight);
              break;
            case "SvgGroup":
              ProcessSvgElement(element, visActivePage, visSVGShape, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight, strSvgUnit, bViewBox);
              break;
            case "SvgUnknownElement":
              break;
            default:
              break;
            }
          string strElement = symbol.ToString();
          }
        }
      else
        {
        MessageBox.Show("No active page in Visio document");
        return;
        }
      }

    private void ProcessSvgElement(SvgElement element, Visio.Page visActivePage, Visio.Shape visSVGShape, double dblWidthRatio, double dblHeightRatio, double dblSVGInchesWidth, double dblSVGInchesHeight, string strSvgUnit, bool bViewBox)
      {
      float fltAngle = 0.0F, fltX = 0.0F, fltY = 0.0F;
      string strTransform = "";
      string strStrokeColor = "";
      string strStrokeWidth = "";
      string strFill = "";

      element.TryGetAttribute("transform", out strTransform);
      element.TryGetAttribute("stroke", out strStrokeColor);
      element.TryGetAttribute("stroke-width", out strStrokeWidth);
      element.TryGetAttribute("fill", out strFill);

      if (!string.IsNullOrEmpty(strTransform))
        {
        if (element.Transforms.Count >= 1 && element.Transforms.ElementAt(0).GetType().Name == "SvgTranslate")
          {
          fltX = ((SvgTranslate)element.Transforms.ElementAt(0)).X;
          fltY = ((SvgTranslate)element.Transforms.ElementAt(0)).Y;
          }
        if (element.Transforms.Count >= 2 && element.Transforms.ElementAt(1).GetType().Name == "SvgRotate")
          {
          fltAngle = ((SvgRotate)element.Transforms.ElementAt(1)).Angle;
          }
        }
      foreach (SvgElement subElement in element.Children)
        {
        switch (subElement.GetType().Name)
          {
          case "SvgLine":
            VisualVisioSVGLightUtil.CreateLine(visActivePage, visSVGShape, subElement, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight);
            break;
          case "SvgText":
            VisualVisioSVGLightUtil.CreateText(visActivePage, visSVGShape, subElement, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight, strSvgUnit, "pt");
            break;
          case "SvgRectangle":
            SvgCustomAttributeCollection arAttribCollection = subElement.CustomAttributes;
            arAttribCollection.TryGetValue("class", out string strClass);
            switch(strClass)
              {
              case "basic label-container":
                VisualVisioSVGLightUtil.CreateRectangleWithText(visActivePage, visSVGShape, subElement, fltX, fltY, fltAngle, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight, bViewBox, strFill, strStrokeColor);
                break;
              default:
                VisualVisioSVGLightUtil.CreateRect(visActivePage, visSVGShape, subElement, fltX, fltY, fltAngle, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight, bViewBox, strFill, strStrokeColor);
                break;
              }
            break;
          case "SvgCircle":
            VisualVisioSVGLightUtil.CreateCircle(visActivePage, visSVGShape, subElement, fltX, fltY, fltAngle, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight, bViewBox, strFill, strStrokeColor);
            break;
          case "SvgPath":
            SvgPath svgPath = ((SvgPath)subElement);
            Svg.Pathing.SvgPathSegmentList arData = svgPath.PathData;
            VisualVisioSVGLightUtil.Create2DPolylineFromPath(visActivePage, visSVGShape, subElement, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight);
            break;
          case "SvgMarker":
            SvgMarker svgMarker = ((SvgMarker)subElement);
            VisualVisioSVGLightUtil.Create2DPolylineFromMarker(visActivePage, visSVGShape, subElement, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight);
            break;
          case "SvgForeignObject":
            break;
          case "SvgGroup":
            ProcessSvgElement(subElement, visActivePage, visSVGShape, dblWidthRatio, dblHeightRatio, dblSVGInchesWidth, dblSVGInchesHeight, strSvgUnit, bViewBox);
            break;
          }
        }
      }

    private void btnClose_Click(object sender, EventArgs e)
      {
      this.Close();
      }
    }
  }
