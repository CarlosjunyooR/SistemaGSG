namespace SistemaGSG
{
    partial class frmOrdemCarreg
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmOrdemCarreg));
            this.ComBoxNomeMotorista = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.TextBoxVez = new System.Windows.Forms.TextBox();
            this.StatusBar = new System.Windows.Forms.StatusStrip();
            this.StatusOnline = new System.Windows.Forms.ToolStripStatusLabel();
            this.StatusOffline = new System.Windows.Forms.ToolStripStatusLabel();
            this.StatusSAP = new System.Windows.Forms.ToolStripStatusLabel();
            this.StatusProgressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.Temporizador = new System.Windows.Forms.Timer(this.components);
            this.DT_SAP = new System.Windows.Forms.DataGridView();
            this.btnConsult = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.NomeMotorista = new System.Windows.Forms.TextBox();
            this.btnPosicao = new System.Windows.Forms.Button();
            this.StatusBar.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DT_SAP)).BeginInit();
            this.SuspendLayout();
            // 
            // ComBoxNomeMotorista
            // 
            this.ComBoxNomeMotorista.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.ComBoxNomeMotorista.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.ComBoxNomeMotorista.FormattingEnabled = true;
            this.ComBoxNomeMotorista.Location = new System.Drawing.Point(12, 66);
            this.ComBoxNomeMotorista.Name = "ComBoxNomeMotorista";
            this.ComBoxNomeMotorista.Size = new System.Drawing.Size(264, 21);
            this.ComBoxNomeMotorista.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(96, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Nome do Motorista";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(84, 111);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(118, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Vez de Carregamento...";
            // 
            // TextBoxVez
            // 
            this.TextBoxVez.Enabled = false;
            this.TextBoxVez.Location = new System.Drawing.Point(90, 127);
            this.TextBoxVez.Name = "TextBoxVez";
            this.TextBoxVez.Size = new System.Drawing.Size(100, 20);
            this.TextBoxVez.TabIndex = 3;
            this.TextBoxVez.Text = "Aguardando...";
            this.TextBoxVez.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // StatusBar
            // 
            this.StatusBar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StatusOnline,
            this.StatusOffline,
            this.StatusSAP,
            this.StatusProgressBar});
            this.StatusBar.Location = new System.Drawing.Point(20, 193);
            this.StatusBar.Name = "StatusBar";
            this.StatusBar.Size = new System.Drawing.Size(264, 22);
            this.StatusBar.TabIndex = 4;
            this.StatusBar.Text = "StatusBarrr";
            // 
            // StatusOnline
            // 
            this.StatusOnline.BackColor = System.Drawing.Color.Green;
            this.StatusOnline.ForeColor = System.Drawing.Color.White;
            this.StatusOnline.Name = "StatusOnline";
            this.StatusOnline.Size = new System.Drawing.Size(146, 17);
            this.StatusOnline.Text = "Status do Servidor.: Online";
            this.StatusOnline.Visible = false;
            // 
            // StatusOffline
            // 
            this.StatusOffline.BackColor = System.Drawing.Color.Red;
            this.StatusOffline.ForeColor = System.Drawing.Color.White;
            this.StatusOffline.Name = "StatusOffline";
            this.StatusOffline.Size = new System.Drawing.Size(147, 17);
            this.StatusOffline.Text = "Status do Servidor.: Offline";
            this.StatusOffline.Visible = false;
            // 
            // StatusSAP
            // 
            this.StatusSAP.BackColor = System.Drawing.Color.Purple;
            this.StatusSAP.ForeColor = System.Drawing.Color.White;
            this.StatusSAP.Name = "StatusSAP";
            this.StatusSAP.Size = new System.Drawing.Size(79, 17);
            this.StatusSAP.Text = "Abrir o SAP....";
            this.StatusSAP.Visible = false;
            // 
            // StatusProgressBar
            // 
            this.StatusProgressBar.Name = "StatusProgressBar";
            this.StatusProgressBar.Size = new System.Drawing.Size(100, 16);
            // 
            // Temporizador
            // 
            this.Temporizador.Enabled = true;
            this.Temporizador.Tick += new System.EventHandler(this.Temporizador_Tick);
            // 
            // DT_SAP
            // 
            this.DT_SAP.AllowUserToAddRows = false;
            this.DT_SAP.AllowUserToDeleteRows = false;
            this.DT_SAP.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DT_SAP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DT_SAP.Location = new System.Drawing.Point(348, 37);
            this.DT_SAP.Name = "DT_SAP";
            this.DT_SAP.Size = new System.Drawing.Size(0, 134);
            this.DT_SAP.TabIndex = 6;
            // 
            // btnConsult
            // 
            this.btnConsult.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnConsult.Location = new System.Drawing.Point(201, 156);
            this.btnConsult.Name = "btnConsult";
            this.btnConsult.Size = new System.Drawing.Size(75, 23);
            this.btnConsult.TabIndex = 7;
            this.btnConsult.Text = "&Consultar";
            this.btnConsult.UseVisualStyleBackColor = true;
            this.btnConsult.Click += new System.EventHandler(this.btnConsult_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(12, 12);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(15, 14);
            this.checkBox1.TabIndex = 8;
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // NomeMotorista
            // 
            this.NomeMotorista.Location = new System.Drawing.Point(348, 12);
            this.NomeMotorista.Name = "NomeMotorista";
            this.NomeMotorista.Size = new System.Drawing.Size(579, 20);
            this.NomeMotorista.TabIndex = 9;
            // 
            // btnPosicao
            // 
            this.btnPosicao.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPosicao.Location = new System.Drawing.Point(201, 125);
            this.btnPosicao.Name = "btnPosicao";
            this.btnPosicao.Size = new System.Drawing.Size(75, 23);
            this.btnPosicao.TabIndex = 10;
            this.btnPosicao.Text = "&Posição";
            this.btnPosicao.UseVisualStyleBackColor = true;
            this.btnPosicao.Click += new System.EventHandler(this.btnPosicao_Click);
            // 
            // frmOrdemCarreg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(304, 235);
            this.Controls.Add(this.btnPosicao);
            this.Controls.Add(this.NomeMotorista);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.btnConsult);
            this.Controls.Add(this.DT_SAP);
            this.Controls.Add(this.StatusBar);
            this.Controls.Add(this.TextBoxVez);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ComBoxNomeMotorista);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(304, 235);
            this.MinimumSize = new System.Drawing.Size(304, 235);
            this.Name = "frmOrdemCarreg";
            this.Style = MetroFramework.MetroColorStyle.Orange;
            this.Load += new System.EventHandler(this.frmOrdemCarreg_Load);
            this.StatusBar.ResumeLayout(false);
            this.StatusBar.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DT_SAP)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox ComBoxNomeMotorista;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TextBoxVez;
        private System.Windows.Forms.StatusStrip StatusBar;
        private System.Windows.Forms.ToolStripStatusLabel StatusOnline;
        private System.Windows.Forms.ToolStripStatusLabel StatusOffline;
        private System.Windows.Forms.ToolStripStatusLabel StatusSAP;
        private System.Windows.Forms.ToolStripProgressBar StatusProgressBar;
        private System.Windows.Forms.Timer Temporizador;
        private System.Windows.Forms.DataGridView DT_SAP;
        private System.Windows.Forms.Button btnConsult;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.TextBox NomeMotorista;
        private System.Windows.Forms.Button btnPosicao;
    }
}

