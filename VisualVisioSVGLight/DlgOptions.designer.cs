
namespace VisualVisioSVGLight
  {
  partial class DlgOptions
    {
    /// <summary>
    /// Variable nécessaire au concepteur.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Nettoyage des ressources utilisées.
    /// </summary>
    /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
    protected override void Dispose(bool disposing)
      {
      if (disposing && (components != null))
        {
        components.Dispose();
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
      System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DlgOptions));
      this.edTemplatePath = new System.Windows.Forms.TextBox();
      this.labTemplatePath = new System.Windows.Forms.Label();
      this.btnStencilExplore = new System.Windows.Forms.Button();
      this.btnOk = new System.Windows.Forms.Button();
      this.btnCancel = new System.Windows.Forms.Button();
      this.edStencilPath = new System.Windows.Forms.TextBox();
      this.labStencilPath = new System.Windows.Forms.Label();
      this.btnTemplateExplore = new System.Windows.Forms.Button();
      this.edProjectPath = new System.Windows.Forms.TextBox();
      this.labProjectPath = new System.Windows.Forms.Label();
      this.btnProjectExplore = new System.Windows.Forms.Button();
      this.tabControl1 = new System.Windows.Forms.TabControl();
      this.tabPage2 = new System.Windows.Forms.TabPage();
      this.tabControl1.SuspendLayout();
      this.tabPage2.SuspendLayout();
      this.SuspendLayout();
      // 
      // edTemplatePath
      // 
      resources.ApplyResources(this.edTemplatePath, "edTemplatePath");
      this.edTemplatePath.Name = "edTemplatePath";
      // 
      // labTemplatePath
      // 
      resources.ApplyResources(this.labTemplatePath, "labTemplatePath");
      this.labTemplatePath.Name = "labTemplatePath";
      // 
      // btnStencilExplore
      // 
      resources.ApplyResources(this.btnStencilExplore, "btnStencilExplore");
      this.btnStencilExplore.Name = "btnStencilExplore";
      this.btnStencilExplore.UseVisualStyleBackColor = true;
      this.btnStencilExplore.Click += new System.EventHandler(this.btnStencilExplore_Click);
      // 
      // btnOk
      // 
      this.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK;
      resources.ApplyResources(this.btnOk, "btnOk");
      this.btnOk.Name = "btnOk";
      this.btnOk.UseVisualStyleBackColor = true;
      this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
      // 
      // btnCancel
      // 
      this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
      resources.ApplyResources(this.btnCancel, "btnCancel");
      this.btnCancel.Name = "btnCancel";
      this.btnCancel.UseVisualStyleBackColor = true;
      this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
      // 
      // edStencilPath
      // 
      resources.ApplyResources(this.edStencilPath, "edStencilPath");
      this.edStencilPath.Name = "edStencilPath";
      // 
      // labStencilPath
      // 
      resources.ApplyResources(this.labStencilPath, "labStencilPath");
      this.labStencilPath.Name = "labStencilPath";
      // 
      // btnTemplateExplore
      // 
      resources.ApplyResources(this.btnTemplateExplore, "btnTemplateExplore");
      this.btnTemplateExplore.Name = "btnTemplateExplore";
      this.btnTemplateExplore.UseVisualStyleBackColor = true;
      this.btnTemplateExplore.Click += new System.EventHandler(this.btnTemplateExplore_Click);
      // 
      // edProjectPath
      // 
      resources.ApplyResources(this.edProjectPath, "edProjectPath");
      this.edProjectPath.Name = "edProjectPath";
      // 
      // labProjectPath
      // 
      resources.ApplyResources(this.labProjectPath, "labProjectPath");
      this.labProjectPath.Name = "labProjectPath";
      // 
      // btnProjectExplore
      // 
      resources.ApplyResources(this.btnProjectExplore, "btnProjectExplore");
      this.btnProjectExplore.Name = "btnProjectExplore";
      this.btnProjectExplore.UseVisualStyleBackColor = true;
      this.btnProjectExplore.Click += new System.EventHandler(this.btnProjectExplore_Click);
      // 
      // tabControl1
      // 
      this.tabControl1.Controls.Add(this.tabPage2);
      resources.ApplyResources(this.tabControl1, "tabControl1");
      this.tabControl1.Name = "tabControl1";
      this.tabControl1.SelectedIndex = 0;
      // 
      // tabPage2
      // 
      this.tabPage2.Controls.Add(this.edTemplatePath);
      this.tabPage2.Controls.Add(this.labTemplatePath);
      this.tabPage2.Controls.Add(this.edStencilPath);
      this.tabPage2.Controls.Add(this.btnStencilExplore);
      this.tabPage2.Controls.Add(this.edProjectPath);
      this.tabPage2.Controls.Add(this.labStencilPath);
      this.tabPage2.Controls.Add(this.btnTemplateExplore);
      this.tabPage2.Controls.Add(this.btnProjectExplore);
      this.tabPage2.Controls.Add(this.labProjectPath);
      resources.ApplyResources(this.tabPage2, "tabPage2");
      this.tabPage2.Name = "tabPage2";
      this.tabPage2.UseVisualStyleBackColor = true;
      // 
      // DlgOptions
      // 
      resources.ApplyResources(this, "$this");
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.Controls.Add(this.tabControl1);
      this.Controls.Add(this.btnCancel);
      this.Controls.Add(this.btnOk);
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
      this.Name = "DlgOptions";
      this.tabControl1.ResumeLayout(false);
      this.tabPage2.ResumeLayout(false);
      this.tabPage2.PerformLayout();
      this.ResumeLayout(false);

      }

    #endregion

    private System.Windows.Forms.TextBox edTemplatePath;
    private System.Windows.Forms.Label labTemplatePath;
    private System.Windows.Forms.Button btnStencilExplore;
    private System.Windows.Forms.Button btnOk;
    private System.Windows.Forms.Button btnCancel;
    private System.Windows.Forms.TextBox edStencilPath;
    private System.Windows.Forms.Label labStencilPath;
    private System.Windows.Forms.Button btnTemplateExplore;
    private System.Windows.Forms.TextBox edProjectPath;
    private System.Windows.Forms.Label labProjectPath;
    private System.Windows.Forms.Button btnProjectExplore;
    private System.Windows.Forms.TabControl tabControl1;
    private System.Windows.Forms.TabPage tabPage2;
    }
  }