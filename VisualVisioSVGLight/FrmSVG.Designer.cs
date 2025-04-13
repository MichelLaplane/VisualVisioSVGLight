namespace VisualVisioSVGLight
  {
  partial class FrmSVG
    {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
      {
      if (disposing && (components != null))
        {
        components.Dispose();
        }
      base.Dispose(disposing);
      }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
      {
      this.btnSvgNativeInsert = new System.Windows.Forms.Button();
      this.btnVisioInsert = new System.Windows.Forms.Button();
      this.edSVG = new System.Windows.Forms.TextBox();
      this.btnOpenPng = new System.Windows.Forms.Button();
      this.btnClose = new System.Windows.Forms.Button();
      this.SuspendLayout();
      // 
      // btnSvgNativeInsert
      // 
      this.btnSvgNativeInsert.Location = new System.Drawing.Point(264, 427);
      this.btnSvgNativeInsert.Name = "btnSvgNativeInsert";
      this.btnSvgNativeInsert.Size = new System.Drawing.Size(120, 23);
      this.btnSvgNativeInsert.TabIndex = 5;
      this.btnSvgNativeInsert.Text = "Insert SVG as Visio";
      this.btnSvgNativeInsert.UseVisualStyleBackColor = true;
      this.btnSvgNativeInsert.Click += new System.EventHandler(this.btnSvgNativeInsert_Click);
      // 
      // btnVisioInsert
      // 
      this.btnVisioInsert.Location = new System.Drawing.Point(138, 427);
      this.btnVisioInsert.Name = "btnVisioInsert";
      this.btnVisioInsert.Size = new System.Drawing.Size(120, 23);
      this.btnVisioInsert.TabIndex = 0;
      this.btnVisioInsert.Text = "Insert SVG as PNG";
      this.btnVisioInsert.UseVisualStyleBackColor = true;
      this.btnVisioInsert.Click += new System.EventHandler(this.btnVisioPngInsert_Click);
      // 
      // edSVG
      // 
      this.edSVG.Location = new System.Drawing.Point(12, 12);
      this.edSVG.Multiline = true;
      this.edSVG.Name = "edSVG";
      this.edSVG.ScrollBars = System.Windows.Forms.ScrollBars.Both;
      this.edSVG.Size = new System.Drawing.Size(491, 409);
      this.edSVG.TabIndex = 2;
      this.edSVG.TextChanged += new System.EventHandler(this.edSVG_TextChanged);
      // 
      // btnOpenPng
      // 
      this.btnOpenPng.Location = new System.Drawing.Point(12, 427);
      this.btnOpenPng.Name = "btnOpenPng";
      this.btnOpenPng.Size = new System.Drawing.Size(120, 23);
      this.btnOpenPng.TabIndex = 4;
      this.btnOpenPng.Text = "Open SVG";
      this.btnOpenPng.UseVisualStyleBackColor = true;
      this.btnOpenPng.Click += new System.EventHandler(this.btnOpenPng_Click);
      // 
      // btnClose
      // 
      this.btnClose.Location = new System.Drawing.Point(428, 427);
      this.btnClose.Name = "btnClose";
      this.btnClose.Size = new System.Drawing.Size(75, 23);
      this.btnClose.TabIndex = 6;
      this.btnClose.Text = "Close";
      this.btnClose.UseVisualStyleBackColor = true;
      this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
      // 
      // FrmSVG
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(515, 456);
      this.Controls.Add(this.btnClose);
      this.Controls.Add(this.btnSvgNativeInsert);
      this.Controls.Add(this.btnVisioInsert);
      this.Controls.Add(this.edSVG);
      this.Controls.Add(this.btnOpenPng);
      this.Name = "FrmSVG";
      this.Text = "Manage SVG files";
      this.ResumeLayout(false);
      this.PerformLayout();

      }

    #endregion
    private System.Windows.Forms.Button btnSvgNativeInsert;
    private System.Windows.Forms.Button btnVisioInsert;
    private System.Windows.Forms.TextBox edSVG;
    private System.Windows.Forms.Button btnOpenPng;
    private System.Windows.Forms.Button btnClose;
    }
  }