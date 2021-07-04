
namespace SistemaGSG
{
    partial class FormDetalhe
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormDetalhe));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.txtCodigoFornecedor = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtContratoFornecedor = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dataDesconto = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.ricTexto = new System.Windows.Forms.RichTextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.txtDesconto = new System.Windows.Forms.TextBox();
            this.txtSafra = new System.Windows.Forms.TextBox();
            this.txtValorDesconto = new SistemaGSG.textValor();
            this.txtIDdetalhe = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4,
            this.Column5,
            this.Column6});
            this.dataGridView1.Location = new System.Drawing.Point(23, 92);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(898, 383);
            this.dataGridView1.TabIndex = 0;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Column1";
            this.Column1.Name = "Column1";
            this.Column1.Visible = false;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Contrato Fornecedor";
            this.Column2.Name = "Column2";
            this.Column2.Width = 150;
            // 
            // Column3
            // 
            this.Column3.HeaderText = "Código Fornecedor";
            this.Column3.Name = "Column3";
            this.Column3.Width = 150;
            // 
            // Column4
            // 
            dataGridViewCellStyle1.Format = "N2";
            dataGridViewCellStyle1.NullValue = null;
            this.Column4.DefaultCellStyle = dataGridViewCellStyle1;
            this.Column4.HeaderText = "Valor Liquido";
            this.Column4.Name = "Column4";
            this.Column4.Width = 150;
            // 
            // Column5
            // 
            this.Column5.HeaderText = "Column5";
            this.Column5.Name = "Column5";
            this.Column5.Visible = false;
            // 
            // Column6
            // 
            this.Column6.HeaderText = "Desconto";
            this.Column6.Name = "Column6";
            this.Column6.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(846, 556);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "&SAP";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtCodigoFornecedor
            // 
            this.txtCodigoFornecedor.Enabled = false;
            this.txtCodigoFornecedor.Location = new System.Drawing.Point(81, 481);
            this.txtCodigoFornecedor.Name = "txtCodigoFornecedor";
            this.txtCodigoFornecedor.Size = new System.Drawing.Size(100, 20);
            this.txtCodigoFornecedor.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 484);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "&Código";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 510);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "&Contrato";
            // 
            // txtContratoFornecedor
            // 
            this.txtContratoFornecedor.Location = new System.Drawing.Point(81, 507);
            this.txtContratoFornecedor.Name = "txtContratoFornecedor";
            this.txtContratoFornecedor.Size = new System.Drawing.Size(100, 20);
            this.txtContratoFornecedor.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(28, 562);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "&Valor";
            // 
            // dataDesconto
            // 
            this.dataDesconto.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dataDesconto.Location = new System.Drawing.Point(81, 533);
            this.dataDesconto.Name = "dataDesconto";
            this.dataDesconto.Size = new System.Drawing.Size(100, 20);
            this.dataDesconto.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(28, 539);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(30, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "&Data";
            // 
            // ricTexto
            // 
            this.ricTexto.Enabled = false;
            this.ricTexto.Location = new System.Drawing.Point(187, 481);
            this.ricTexto.Name = "ricTexto";
            this.ricTexto.Size = new System.Drawing.Size(217, 99);
            this.ricTexto.TabIndex = 11;
            this.ricTexto.Text = "Digite aqui, Informações referente ao desconto...";
            this.ricTexto.DoubleClick += new System.EventHandler(this.ricTexto_DoubleClick);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(765, 556);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 12;
            this.button2.Text = "&Gravar";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // txtDesconto
            // 
            this.txtDesconto.Location = new System.Drawing.Point(821, 504);
            this.txtDesconto.Name = "txtDesconto";
            this.txtDesconto.Size = new System.Drawing.Size(100, 20);
            this.txtDesconto.TabIndex = 13;
            this.txtDesconto.Visible = false;
            // 
            // txtSafra
            // 
            this.txtSafra.Location = new System.Drawing.Point(821, 530);
            this.txtSafra.Name = "txtSafra";
            this.txtSafra.Size = new System.Drawing.Size(100, 20);
            this.txtSafra.TabIndex = 15;
            this.txtSafra.Visible = false;
            // 
            // txtValorDesconto
            // 
            this.txtValorDesconto.Location = new System.Drawing.Point(81, 559);
            this.txtValorDesconto.Name = "txtValorDesconto";
            this.txtValorDesconto.Size = new System.Drawing.Size(100, 20);
            this.txtValorDesconto.TabIndex = 8;
            this.txtValorDesconto.Tag = "";
            this.txtValorDesconto.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtValorDesconto.DoubleClick += new System.EventHandler(this.textValor1_DoubleClick);
            // 
            // txtIDdetalhe
            // 
            this.txtIDdetalhe.Location = new System.Drawing.Point(821, 481);
            this.txtIDdetalhe.Name = "txtIDdetalhe";
            this.txtIDdetalhe.Size = new System.Drawing.Size(100, 20);
            this.txtIDdetalhe.TabIndex = 16;
            this.txtIDdetalhe.Visible = false;
            // 
            // FormDetalhe
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(944, 602);
            this.Controls.Add(this.txtIDdetalhe);
            this.Controls.Add(this.txtSafra);
            this.Controls.Add(this.txtDesconto);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.ricTexto);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.dataDesconto);
            this.Controls.Add(this.txtValorDesconto);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtContratoFornecedor);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtCodigoFornecedor);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(944, 602);
            this.MinimumSize = new System.Drawing.Size(944, 602);
            this.Name = "FormDetalhe";
            this.Style = MetroFramework.MetroColorStyle.Orange;
            this.Text = "Adicionar Desconto";
            this.Theme = MetroFramework.MetroThemeStyle.Default;
            this.Load += new System.EventHandler(this.FormDetalhe_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtCodigoFornecedor;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtContratoFornecedor;
        private System.Windows.Forms.Label label3;
        private textValor txtValorDesconto;
        private System.Windows.Forms.DateTimePicker dataDesconto;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.RichTextBox ricTexto;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox txtDesconto;
        private System.Windows.Forms.TextBox txtSafra;
        private System.Windows.Forms.TextBox txtIDdetalhe;
    }
}