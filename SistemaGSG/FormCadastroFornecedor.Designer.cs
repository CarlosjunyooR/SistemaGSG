
namespace SistemaGSG
{
    partial class FormCadastroFornecedor
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
            this.lbl1 = new System.Windows.Forms.Label();
            this.txtCodigoFornec = new System.Windows.Forms.TextBox();
            this.txtNomeFornec = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbl1
            // 
            this.lbl1.AutoSize = true;
            this.lbl1.Location = new System.Drawing.Point(15, 24);
            this.lbl1.Name = "lbl1";
            this.lbl1.Size = new System.Drawing.Size(40, 13);
            this.lbl1.TabIndex = 1;
            this.lbl1.Text = "&Código";
            // 
            // txtCodigoFornec
            // 
            this.txtCodigoFornec.Enabled = false;
            this.txtCodigoFornec.Location = new System.Drawing.Point(61, 21);
            this.txtCodigoFornec.Name = "txtCodigoFornec";
            this.txtCodigoFornec.Size = new System.Drawing.Size(100, 20);
            this.txtCodigoFornec.TabIndex = 2;
            // 
            // txtNomeFornec
            // 
            this.txtNomeFornec.Location = new System.Drawing.Point(61, 47);
            this.txtNomeFornec.Name = "txtNomeFornec";
            this.txtNomeFornec.Size = new System.Drawing.Size(237, 20);
            this.txtNomeFornec.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "&Nome";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(18, 96);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(280, 87);
            this.button1.TabIndex = 5;
            this.button1.Text = "&Cadastrar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // FormCadastroFornecedor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(321, 206);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtNomeFornec);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtCodigoFornec);
            this.Controls.Add(this.lbl1);
            this.Name = "FormCadastroFornecedor";
            this.Style = MetroFramework.MetroColorStyle.Orange;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl1;
        private System.Windows.Forms.TextBox txtCodigoFornec;
        private System.Windows.Forms.TextBox txtNomeFornec;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
    }
}