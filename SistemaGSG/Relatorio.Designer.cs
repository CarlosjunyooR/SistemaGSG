namespace SistemaGSG
{
    partial class Relatorio
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
            this.components = new System.ComponentModel.Container();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.dBSGSGAcessoBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dBSGSGAcesso = new SistemaGSG.DBSGSGAcesso();
            ((System.ComponentModel.ISupportInitialize)(this.dBSGSGAcessoBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dBSGSGAcesso)).BeginInit();
            this.SuspendLayout();
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "AcessoDB";
            reportDataSource1.Value = this.dBSGSGAcessoBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "SistemaGSG.Relatorios.Acesso.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(0, 0);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.ServerReport.BearerToken = null;
            this.reportViewer1.Size = new System.Drawing.Size(800, 450);
            this.reportViewer1.TabIndex = 0;
            // 
            // dBSGSGAcessoBindingSource
            // 
            this.dBSGSGAcessoBindingSource.DataSource = this.dBSGSGAcesso;
            this.dBSGSGAcessoBindingSource.Position = 0;
            // 
            // dBSGSGAcesso
            // 
            this.dBSGSGAcesso.DataSetName = "DBSGSGAcesso";
            this.dBSGSGAcesso.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // Relatorio
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.reportViewer1);
            this.Name = "Relatorio";
            this.Text = "Relatorio";
            this.Load += new System.EventHandler(this.Relatorio_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dBSGSGAcessoBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dBSGSGAcesso)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource dBSGSGAcessoBindingSource;
        private DBSGSGAcesso dBSGSGAcesso;
    }
}