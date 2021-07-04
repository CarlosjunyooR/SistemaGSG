namespace SistemaGSG
{
    partial class FormPedido
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormPedido));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label19 = new System.Windows.Forms.Label();
            this.MesRef = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnSAP = new System.Windows.Forms.Button();
            this.btnPedidoNormal = new System.Windows.Forms.Button();
            this.btnPedido = new System.Windows.Forms.Button();
            this.chboxMigo = new MetroFramework.Controls.MetroCheckBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.btnFilterMigo = new System.Windows.Forms.Button();
            this.btnMigo = new System.Windows.Forms.Button();
            this.chboxMiro = new MetroFramework.Controls.MetroCheckBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtPedido = new System.Windows.Forms.TextBox();
            this.dtLanc = new System.Windows.Forms.DateTimePicker();
            this.dtDoc = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtNf = new System.Windows.Forms.TextBox();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.btnCriarMiro = new System.Windows.Forms.Button();
            this.btnFilterPedidoNormal = new System.Windows.Forms.Button();
            this.groupBox11 = new System.Windows.Forms.GroupBox();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.label8 = new System.Windows.Forms.Label();
            this.dtMiroFatura = new System.Windows.Forms.DateTimePicker();
            this.label7 = new System.Windows.Forms.Label();
            this.btnVoltar = new System.Windows.Forms.Button();
            this.ProgBar = new System.Windows.Forms.ProgressBar();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.tBBOLETOBindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.tBBOLETOBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.bancoDadosDataSetBoletosBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.groupBox11.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tBBOLETOBindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tBBOLETOBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bancoDadosDataSetBoletosBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label19);
            this.groupBox1.Controls.Add(this.MesRef);
            this.groupBox1.Controls.Add(this.label21);
            this.groupBox1.Location = new System.Drawing.Point(6, 22);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(314, 169);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Dados de Criação";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(43, 82);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(27, 13);
            this.label19.TabIndex = 32;
            this.label19.Text = "Mês";
            // 
            // MesRef
            // 
            this.MesRef.Location = new System.Drawing.Point(76, 79);
            this.MesRef.Name = "MesRef";
            this.MesRef.Size = new System.Drawing.Size(85, 20);
            this.MesRef.TabIndex = 27;
            this.MesRef.Text = "02/2021";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(43, 56);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(27, 13);
            this.label21.TabIndex = 11;
            this.label21.Text = "Item";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnSAP);
            this.groupBox3.Controls.Add(this.btnPedidoNormal);
            this.groupBox3.Controls.Add(this.btnPedido);
            this.groupBox3.Controls.Add(this.chboxMigo);
            this.groupBox3.Controls.Add(this.groupBox1);
            this.groupBox3.Location = new System.Drawing.Point(23, 29);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(436, 203);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Criar Pedido";
            // 
            // btnSAP
            // 
            this.btnSAP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSAP.Location = new System.Drawing.Point(355, 113);
            this.btnSAP.Name = "btnSAP";
            this.btnSAP.Size = new System.Drawing.Size(75, 23);
            this.btnSAP.TabIndex = 28;
            this.btnSAP.Text = "&Abrir SAP";
            this.btnSAP.UseVisualStyleBackColor = true;
            this.btnSAP.Click += new System.EventHandler(this.btnSAP_Click);
            // 
            // btnPedidoNormal
            // 
            this.btnPedidoNormal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPedidoNormal.Location = new System.Drawing.Point(355, 145);
            this.btnPedidoNormal.Name = "btnPedidoNormal";
            this.btnPedidoNormal.Size = new System.Drawing.Size(75, 23);
            this.btnPedidoNormal.TabIndex = 26;
            this.btnPedidoNormal.Text = "&Filtrar";
            this.btnPedidoNormal.UseVisualStyleBackColor = true;
            this.btnPedidoNormal.Click += new System.EventHandler(this.btnPedidoNormal_Click);
            // 
            // btnPedido
            // 
            this.btnPedido.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPedido.Location = new System.Drawing.Point(355, 174);
            this.btnPedido.Name = "btnPedido";
            this.btnPedido.Size = new System.Drawing.Size(75, 23);
            this.btnPedido.TabIndex = 24;
            this.btnPedido.Text = "&Criar Pedido";
            this.btnPedido.UseVisualStyleBackColor = true;
            this.btnPedido.Click += new System.EventHandler(this.btnPedido_Click_1);
            // 
            // chboxMigo
            // 
            this.chboxMigo.AutoSize = true;
            this.chboxMigo.Location = new System.Drawing.Point(379, 13);
            this.chboxMigo.Name = "chboxMigo";
            this.chboxMigo.Size = new System.Drawing.Size(51, 15);
            this.chboxMigo.TabIndex = 4;
            this.chboxMigo.Text = "Migo";
            this.chboxMigo.UseSelectable = true;
            this.chboxMigo.CheckedChanged += new System.EventHandler(this.metroCheckBox1_CheckedChanged);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.btnFilterMigo);
            this.groupBox5.Controls.Add(this.btnMigo);
            this.groupBox5.Controls.Add(this.chboxMiro);
            this.groupBox5.Controls.Add(this.groupBox7);
            this.groupBox5.Location = new System.Drawing.Point(465, 29);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(436, 203);
            this.groupBox5.TabIndex = 4;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Criar Migo";
            // 
            // btnFilterMigo
            // 
            this.btnFilterMigo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFilterMigo.Location = new System.Drawing.Point(355, 134);
            this.btnFilterMigo.Name = "btnFilterMigo";
            this.btnFilterMigo.Size = new System.Drawing.Size(75, 23);
            this.btnFilterMigo.TabIndex = 25;
            this.btnFilterMigo.Text = "&Filtrar Migo";
            this.btnFilterMigo.UseVisualStyleBackColor = true;
            this.btnFilterMigo.Click += new System.EventHandler(this.btnFilterMigo_Click);
            // 
            // btnMigo
            // 
            this.btnMigo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMigo.Location = new System.Drawing.Point(355, 162);
            this.btnMigo.Name = "btnMigo";
            this.btnMigo.Size = new System.Drawing.Size(75, 23);
            this.btnMigo.TabIndex = 24;
            this.btnMigo.Text = "&Criar Migo";
            this.btnMigo.UseVisualStyleBackColor = true;
            this.btnMigo.Click += new System.EventHandler(this.btnMigo_Click);
            // 
            // chboxMiro
            // 
            this.chboxMiro.AutoSize = true;
            this.chboxMiro.Location = new System.Drawing.Point(382, 13);
            this.chboxMiro.Name = "chboxMiro";
            this.chboxMiro.Size = new System.Drawing.Size(48, 15);
            this.chboxMiro.TabIndex = 1;
            this.chboxMiro.Text = "Miro";
            this.chboxMiro.UseSelectable = true;
            this.chboxMiro.CheckedChanged += new System.EventHandler(this.chboxMiro_CheckedChanged);
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.label3);
            this.groupBox7.Controls.Add(this.txtPedido);
            this.groupBox7.Controls.Add(this.dtLanc);
            this.groupBox7.Controls.Add(this.dtDoc);
            this.groupBox7.Controls.Add(this.label4);
            this.groupBox7.Controls.Add(this.label5);
            this.groupBox7.Controls.Add(this.label6);
            this.groupBox7.Controls.Add(this.txtNf);
            this.groupBox7.Location = new System.Drawing.Point(6, 22);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(314, 169);
            this.groupBox7.TabIndex = 0;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Geral";
            this.groupBox7.Enter += new System.EventHandler(this.groupBox7_Enter);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(40, 114);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Nº Pedido";
            // 
            // txtPedido
            // 
            this.txtPedido.Location = new System.Drawing.Point(138, 110);
            this.txtPedido.Name = "txtPedido";
            this.txtPedido.Size = new System.Drawing.Size(97, 20);
            this.txtPedido.TabIndex = 6;
            this.txtPedido.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dtLanc
            // 
            this.dtLanc.CustomFormat = "dd.MM.yyyy";
            this.dtLanc.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtLanc.Location = new System.Drawing.Point(138, 58);
            this.dtLanc.Name = "dtLanc";
            this.dtLanc.Size = new System.Drawing.Size(97, 20);
            this.dtLanc.TabIndex = 5;
            // 
            // dtDoc
            // 
            this.dtDoc.CustomFormat = "dd.MM.yyyy";
            this.dtDoc.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtDoc.Location = new System.Drawing.Point(138, 32);
            this.dtDoc.Name = "dtDoc";
            this.dtDoc.Size = new System.Drawing.Size(97, 20);
            this.dtDoc.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(40, 87);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(75, 13);
            this.label4.TabIndex = 2;
            this.label4.Text = "Nº Nota Fiscal";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(40, 61);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(92, 13);
            this.label5.TabIndex = 1;
            this.label5.Text = "Data Lançamento";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(40, 36);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(86, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "Data documento";
            // 
            // txtNf
            // 
            this.txtNf.Location = new System.Drawing.Point(138, 84);
            this.txtNf.Name = "txtNf";
            this.txtNf.Size = new System.Drawing.Size(97, 20);
            this.txtNf.TabIndex = 1;
            this.txtNf.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.btnCriarMiro);
            this.groupBox9.Controls.Add(this.btnFilterPedidoNormal);
            this.groupBox9.Controls.Add(this.groupBox11);
            this.groupBox9.Location = new System.Drawing.Point(907, 29);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(436, 203);
            this.groupBox9.TabIndex = 5;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "Criar Miro";
            // 
            // btnCriarMiro
            // 
            this.btnCriarMiro.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCriarMiro.Location = new System.Drawing.Point(355, 162);
            this.btnCriarMiro.Name = "btnCriarMiro";
            this.btnCriarMiro.Size = new System.Drawing.Size(75, 23);
            this.btnCriarMiro.TabIndex = 12;
            this.btnCriarMiro.Text = "&Criar Miro";
            this.btnCriarMiro.UseVisualStyleBackColor = true;
            this.btnCriarMiro.Click += new System.EventHandler(this.btnCriarMiro_Click);
            // 
            // btnFilterPedidoNormal
            // 
            this.btnFilterPedidoNormal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFilterPedidoNormal.Location = new System.Drawing.Point(355, 132);
            this.btnFilterPedidoNormal.Name = "btnFilterPedidoNormal";
            this.btnFilterPedidoNormal.Size = new System.Drawing.Size(75, 23);
            this.btnFilterPedidoNormal.TabIndex = 11;
            this.btnFilterPedidoNormal.Text = "&Filtrar";
            this.btnFilterPedidoNormal.UseVisualStyleBackColor = true;
            this.btnFilterPedidoNormal.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // groupBox11
            // 
            this.groupBox11.Controls.Add(this.dateTimePicker2);
            this.groupBox11.Controls.Add(this.label8);
            this.groupBox11.Controls.Add(this.dtMiroFatura);
            this.groupBox11.Controls.Add(this.label7);
            this.groupBox11.Location = new System.Drawing.Point(6, 22);
            this.groupBox11.Name = "groupBox11";
            this.groupBox11.Size = new System.Drawing.Size(314, 169);
            this.groupBox11.TabIndex = 0;
            this.groupBox11.TabStop = false;
            this.groupBox11.Text = "Dados Básicos";
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.CustomFormat = "dd.MM.yyyy";
            this.dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker2.Location = new System.Drawing.Point(138, 52);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(97, 20);
            this.dateTimePicker2.TabIndex = 9;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(40, 58);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(92, 13);
            this.label8.TabIndex = 8;
            this.label8.Text = "Data Lançamento";
            // 
            // dtMiroFatura
            // 
            this.dtMiroFatura.CustomFormat = "dd.MM.yyyy";
            this.dtMiroFatura.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtMiroFatura.Location = new System.Drawing.Point(138, 26);
            this.dtMiroFatura.Name = "dtMiroFatura";
            this.dtMiroFatura.Size = new System.Drawing.Size(97, 20);
            this.dtMiroFatura.TabIndex = 8;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(40, 32);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(78, 13);
            this.label7.TabIndex = 1;
            this.label7.Text = "Data da Fatura";
            // 
            // btnVoltar
            // 
            this.btnVoltar.Location = new System.Drawing.Point(1267, 10);
            this.btnVoltar.Name = "btnVoltar";
            this.btnVoltar.Size = new System.Drawing.Size(75, 23);
            this.btnVoltar.TabIndex = 24;
            this.btnVoltar.Text = "&Voltar";
            this.btnVoltar.UseVisualStyleBackColor = true;
            this.btnVoltar.Click += new System.EventHandler(this.btnVoltar_Click);
            // 
            // ProgBar
            // 
            this.ProgBar.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ProgBar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(66)))), ((int)(((byte)(135)))), ((int)(((byte)(245)))));
            this.ProgBar.ForeColor = System.Drawing.Color.Red;
            this.ProgBar.Location = new System.Drawing.Point(22, 238);
            this.ProgBar.Maximum = 1000;
            this.ProgBar.Name = "ProgBar";
            this.ProgBar.Size = new System.Drawing.Size(1320, 10);
            this.ProgBar.TabIndex = 26;
            // 
            // groupBox4
            // 
            this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox4.Controls.Add(this.dataGridView2);
            this.groupBox4.Controls.Add(this.dataGridView1);
            this.groupBox4.Location = new System.Drawing.Point(17, 254);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(1326, 511);
            this.groupBox4.TabIndex = 27;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Notas Fiscais de Energia";
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(21, 268);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(1299, 237);
            this.dataGridView2.TabIndex = 29;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(21, 19);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1299, 243);
            this.dataGridView1.TabIndex = 28;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(107, 56);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(19, 13);
            this.label1.TabIndex = 33;
            this.label1.Text = "de";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(146, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(13, 13);
            this.label2.TabIndex = 34;
            this.label2.Text = "2";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(76, 56);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(13, 13);
            this.label9.TabIndex = 35;
            this.label9.Text = "1";
            // 
            // FormPedido
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1365, 788);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.ProgBar);
            this.Controls.Add(this.btnVoltar);
            this.Controls.Add(this.groupBox9);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox3);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(1365, 788);
            this.MinimumSize = new System.Drawing.Size(1078, 788);
            this.Name = "FormPedido";
            this.Style = MetroFramework.MetroColorStyle.Orange;
            this.Load += new System.EventHandler(this.FormPedido_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.groupBox9.ResumeLayout(false);
            this.groupBox11.ResumeLayout(false);
            this.groupBox11.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tBBOLETOBindingSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tBBOLETOBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bancoDadosDataSetBoletosBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.DateTimePicker dtDoc;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtNf;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.GroupBox groupBox11;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtPedido;
        private System.Windows.Forms.DateTimePicker dtLanc;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.DateTimePicker dtMiroFatura;
        private System.Windows.Forms.Label label7;
        private MetroFramework.Controls.MetroCheckBox chboxMigo;
        private MetroFramework.Controls.MetroCheckBox chboxMiro;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Button btnPedido;
        private System.Windows.Forms.Button btnMigo;
        private System.Windows.Forms.Button btnFilterMigo;
        private System.Windows.Forms.Button btnPedidoNormal;
        private System.Windows.Forms.Button btnFilterPedidoNormal;
        private System.Windows.Forms.Button btnCriarMiro;
        private System.Windows.Forms.Button btnVoltar;
        private System.Windows.Forms.ProgressBar ProgBar;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.BindingSource bancoDadosDataSetBoletosBindingSource;
        private System.Windows.Forms.BindingSource tBBOLETOBindingSource;
        private System.Windows.Forms.BindingSource tBBOLETOBindingSource1;
        private System.Windows.Forms.DataGridViewTextBoxColumn idDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn materialDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn descitemDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn qtdDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn centroDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn custoDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn codimpDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn basecalculoDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn vlicmsDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtpedidoDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn materialdifDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn descitemdifDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn qtddifDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn centrodifDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn custodifDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn codimpdifDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn vldifDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn ivadifDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn emissaoDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn nfeDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn errDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn errcolDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtmiroDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn datavencDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn mesrefDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn codunicoDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn pedidoDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn migoDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn miroDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecoepDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn valormiroDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn statusDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn empresaDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn mesduplDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn vlduplDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn nowdateDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn idDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn materialDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn descitemDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn qtdDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn centroDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn custoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn codimpDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn basecalculoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn vlicmsDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtpedidoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn materialdifDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn descitemdifDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn qtddifDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn centrodifDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn custodifDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn codimpdifDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn vldifDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn ivadifDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn emissaoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn nfeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn errDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn errcolDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtmiroDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn datavencDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mesrefDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn codunicoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn pedidoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn migoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn miroDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecoepDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn valormiroDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn statusDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn empresaDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mesduplDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn vlduplDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn nowdateDataGridViewTextBoxColumn;
        private System.Windows.Forms.TextBox MesRef;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Button btnSAP;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}