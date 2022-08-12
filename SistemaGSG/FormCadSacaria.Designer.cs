
namespace SistemaGSG
{
    partial class FormCadSacaria
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtInicSacaria = new System.Windows.Forms.TextBox();
            this.txtFimSacaria = new System.Windows.Forms.TextBox();
            this.btnInsert = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(53, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Início";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(161, 33);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(23, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Fim";
            // 
            // txtInicSacaria
            // 
            this.txtInicSacaria.Location = new System.Drawing.Point(23, 49);
            this.txtInicSacaria.Name = "txtInicSacaria";
            this.txtInicSacaria.Size = new System.Drawing.Size(100, 20);
            this.txtInicSacaria.TabIndex = 2;
            this.txtInicSacaria.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtFimSacaria
            // 
            this.txtFimSacaria.Location = new System.Drawing.Point(129, 49);
            this.txtFimSacaria.Name = "txtFimSacaria";
            this.txtFimSacaria.Size = new System.Drawing.Size(100, 20);
            this.txtFimSacaria.TabIndex = 3;
            this.txtFimSacaria.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // btnInsert
            // 
            this.btnInsert.Location = new System.Drawing.Point(154, 94);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(75, 23);
            this.btnInsert.TabIndex = 4;
            this.btnInsert.Text = "&Inserir";
            this.btnInsert.UseVisualStyleBackColor = true;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // FormCadSacaria
            // 
            this.ClientSize = new System.Drawing.Size(252, 140);
            this.Controls.Add(this.btnInsert);
            this.Controls.Add(this.txtFimSacaria);
            this.Controls.Add(this.txtInicSacaria);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(312, 140);
            this.MinimizeBox = false;
            this.Name = "FormCadSacaria";
            this.Style = MetroFramework.MetroColorStyle.Orange;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtInicSacaria;
        private System.Windows.Forms.TextBox txtFimSacaria;
        private System.Windows.Forms.Button btnInsert;
    }
}
