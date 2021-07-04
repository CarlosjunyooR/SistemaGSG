
namespace SistemaGSG
{
    partial class FormDesconto
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormDesconto));
            this.dtDebitos = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CÓDIGO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NOME = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DATA = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TIPO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DESCRICAO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.QUANTI = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.UNID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.VALORDIV = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.txtSafra = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dataDesconto = new System.Windows.Forms.DateTimePicker();
            this.txtUnidade = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.txtQuant = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.cmbDesc = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtCodFornecedor = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.dtDetalhe = new System.Windows.Forms.DataGridView();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column11 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column12 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column13 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rdSafra2021 = new System.Windows.Forms.RadioButton();
            this.rdSafra2020 = new System.Windows.Forms.RadioButton();
            this.btnVoltar = new System.Windows.Forms.Button();
            this.textBoxCodigo = new System.Windows.Forms.TextBox();
            this.textBoxDesc = new System.Windows.Forms.TextBox();
            this.txtSafraOu = new System.Windows.Forms.TextBox();
            this.txtID = new System.Windows.Forms.TextBox();
            this.txtDescontado = new System.Windows.Forms.TextBox();
            this.txtSaldo = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtValorDTdesconto = new System.Windows.Forms.TextBox();
            this.txtDivida = new SistemaGSG.textValor();
            ((System.ComponentModel.ISupportInitialize)(this.dtDebitos)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtDetalhe)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dtDebitos
            // 
            this.dtDebitos.AllowUserToAddRows = false;
            this.dtDebitos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dtDebitos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtDebitos.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.CÓDIGO,
            this.NOME,
            this.DATA,
            this.TIPO,
            this.DESCRICAO,
            this.QUANTI,
            this.UNID,
            this.VALORDIV,
            this.Column8,
            this.Column9});
            this.dtDebitos.Location = new System.Drawing.Point(12, 63);
            this.dtDebitos.Name = "dtDebitos";
            this.dtDebitos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dtDebitos.Size = new System.Drawing.Size(1375, 216);
            this.dtDebitos.TabIndex = 0;
            this.dtDebitos.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtDebitos_CellClick);
            // 
            // Column1
            // 
            this.Column1.DataPropertyName = "id";
            this.Column1.HeaderText = "Column1";
            this.Column1.Name = "Column1";
            this.Column1.Visible = false;
            // 
            // CÓDIGO
            // 
            this.CÓDIGO.DataPropertyName = "col_Fornecedor";
            this.CÓDIGO.HeaderText = "CÓDIGO";
            this.CÓDIGO.Name = "CÓDIGO";
            // 
            // NOME
            // 
            this.NOME.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.NOME.DataPropertyName = "col_Nome";
            this.NOME.HeaderText = "NOME FORNECEDOR";
            this.NOME.Name = "NOME";
            this.NOME.Width = 130;
            // 
            // DATA
            // 
            this.DATA.DataPropertyName = "col_Data";
            this.DATA.HeaderText = "DATA";
            this.DATA.Name = "DATA";
            // 
            // TIPO
            // 
            this.TIPO.DataPropertyName = "col_Desconto";
            this.TIPO.HeaderText = "TIPO";
            this.TIPO.Name = "TIPO";
            this.TIPO.Visible = false;
            // 
            // DESCRICAO
            // 
            this.DESCRICAO.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.DESCRICAO.DataPropertyName = "col_Descricao";
            this.DESCRICAO.HeaderText = "DESCONTO";
            this.DESCRICAO.Name = "DESCRICAO";
            this.DESCRICAO.Width = 92;
            // 
            // QUANTI
            // 
            this.QUANTI.DataPropertyName = "col_Quantidade";
            this.QUANTI.HeaderText = "QUANTI";
            this.QUANTI.Name = "QUANTI";
            // 
            // UNID
            // 
            this.UNID.DataPropertyName = "col_Unidade";
            this.UNID.HeaderText = "UNID";
            this.UNID.Name = "UNID";
            // 
            // VALORDIV
            // 
            this.VALORDIV.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.VALORDIV.DataPropertyName = "col_ValorDiv";
            dataGridViewCellStyle1.Format = "N2";
            dataGridViewCellStyle1.NullValue = "0";
            this.VALORDIV.DefaultCellStyle = dataGridViewCellStyle1;
            this.VALORDIV.HeaderText = "VALOR DIVIDA";
            this.VALORDIV.Name = "VALORDIV";
            this.VALORDIV.Width = 98;
            // 
            // Column8
            // 
            this.Column8.DataPropertyName = "col_dataImport";
            this.Column8.HeaderText = "ADICIONADO EM";
            this.Column8.Name = "Column8";
            this.Column8.Width = 150;
            // 
            // Column9
            // 
            this.Column9.DataPropertyName = "safra";
            this.Column9.HeaderText = "SAFRA";
            this.Column9.Name = "Column9";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.txtDivida);
            this.groupBox1.Controls.Add(this.txtSafra);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.dataDesconto);
            this.groupBox1.Controls.Add(this.txtUnidade);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnSave);
            this.groupBox1.Controls.Add(this.txtQuant);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.cmbDesc);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.txtCodFornecedor);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Location = new System.Drawing.Point(12, 516);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(396, 168);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Apontamento de Descontos";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(230, 139);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "&Excluir";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // txtSafra
            // 
            this.txtSafra.Location = new System.Drawing.Point(286, 112);
            this.txtSafra.Name = "txtSafra";
            this.txtSafra.Size = new System.Drawing.Size(100, 20);
            this.txtSafra.TabIndex = 24;
            this.txtSafra.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtSafra.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtSafra_KeyUp);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(210, 115);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(32, 13);
            this.label2.TabIndex = 23;
            this.label2.Text = "Safra";
            // 
            // dataDesconto
            // 
            this.dataDesconto.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dataDesconto.Location = new System.Drawing.Point(80, 71);
            this.dataDesconto.Name = "dataDesconto";
            this.dataDesconto.Size = new System.Drawing.Size(100, 20);
            this.dataDesconto.TabIndex = 5;
            this.dataDesconto.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dataDesconto_KeyUp);
            // 
            // txtUnidade
            // 
            this.txtUnidade.Location = new System.Drawing.Point(286, 60);
            this.txtUnidade.Name = "txtUnidade";
            this.txtUnidade.Size = new System.Drawing.Size(100, 20);
            this.txtUnidade.TabIndex = 22;
            this.txtUnidade.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtUnidade_KeyUp);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(210, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 13);
            this.label1.TabIndex = 21;
            this.label1.Text = "Unid.";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(311, 139);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 20;
            this.btnSave.Text = "&Gravar";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click_1);
            // 
            // txtQuant
            // 
            this.txtQuant.Location = new System.Drawing.Point(286, 34);
            this.txtQuant.Name = "txtQuant";
            this.txtQuant.Size = new System.Drawing.Size(100, 20);
            this.txtQuant.TabIndex = 18;
            this.txtQuant.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtQuant_KeyUp);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(210, 89);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(68, 13);
            this.label11.TabIndex = 17;
            this.label11.Text = "Valor da Div.";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(210, 37);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(62, 13);
            this.label12.TabIndex = 16;
            this.label12.Text = "Quantidade";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(4, 116);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(74, 13);
            this.label9.TabIndex = 14;
            this.label9.Text = "Tipo de Desc.";
            // 
            // cmbDesc
            // 
            this.cmbDesc.FormattingEnabled = true;
            this.cmbDesc.Items.AddRange(new object[] {
            "13",
            "14",
            "16",
            "36"});
            this.cmbDesc.Location = new System.Drawing.Point(80, 113);
            this.cmbDesc.Name = "cmbDesc";
            this.cmbDesc.Size = new System.Drawing.Size(100, 21);
            this.cmbDesc.TabIndex = 13;
            this.cmbDesc.KeyUp += new System.Windows.Forms.KeyEventHandler(this.cmbDesc_KeyUp);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(4, 78);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(30, 13);
            this.label8.TabIndex = 12;
            this.label8.Text = "Data";
            // 
            // txtCodFornecedor
            // 
            this.txtCodFornecedor.Location = new System.Drawing.Point(80, 35);
            this.txtCodFornecedor.Name = "txtCodFornecedor";
            this.txtCodFornecedor.Size = new System.Drawing.Size(100, 20);
            this.txtCodFornecedor.TabIndex = 2;
            this.txtCodFornecedor.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtCodFornecedor_KeyUp);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(4, 38);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(40, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "Código";
            // 
            // dtDetalhe
            // 
            this.dtDetalhe.AllowUserToAddRows = false;
            this.dtDetalhe.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtDetalhe.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column2,
            this.Column3,
            this.Column4,
            this.Column5,
            this.Column6,
            this.Column7,
            this.Column10,
            this.Column11,
            this.Column12,
            this.Column13});
            this.dtDetalhe.Location = new System.Drawing.Point(12, 294);
            this.dtDetalhe.Name = "dtDetalhe";
            this.dtDetalhe.Size = new System.Drawing.Size(1375, 216);
            this.dtDetalhe.TabIndex = 4;
            this.dtDetalhe.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtDetalhe_CellClick);
            this.dtDetalhe.DoubleClick += new System.EventHandler(this.dtDetalhe_DoubleClick);
            // 
            // Column2
            // 
            this.Column2.DataPropertyName = "id";
            this.Column2.HeaderText = "Column2";
            this.Column2.Name = "Column2";
            this.Column2.Visible = false;
            // 
            // Column3
            // 
            this.Column3.DataPropertyName = "col_Fornec";
            this.Column3.HeaderText = "Column3";
            this.Column3.Name = "Column3";
            this.Column3.Visible = false;
            // 
            // Column4
            // 
            this.Column4.DataPropertyName = "col_Contrato";
            this.Column4.HeaderText = "Column4";
            this.Column4.Name = "Column4";
            this.Column4.Visible = false;
            // 
            // Column5
            // 
            this.Column5.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.Column5.DataPropertyName = "col_Data";
            this.Column5.HeaderText = "DATA";
            this.Column5.Name = "Column5";
            this.Column5.Width = 61;
            // 
            // Column6
            // 
            this.Column6.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.Column6.DataPropertyName = "col_ValorFixo";
            this.Column6.HeaderText = "VALOR FIXO";
            this.Column6.Name = "Column6";
            this.Column6.Width = 95;
            // 
            // Column7
            // 
            this.Column7.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.Column7.DataPropertyName = "col_TextoDesconto";
            this.Column7.HeaderText = "TEXTO DESC.";
            this.Column7.Name = "Column7";
            this.Column7.Width = 95;
            // 
            // Column10
            // 
            this.Column10.DataPropertyName = "col_TipoDesc";
            this.Column10.HeaderText = "Column10";
            this.Column10.Name = "Column10";
            this.Column10.Visible = false;
            // 
            // Column11
            // 
            this.Column11.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.Column11.DataPropertyName = "col_dataImport";
            this.Column11.HeaderText = "ADICIONADO EM";
            this.Column11.Name = "Column11";
            this.Column11.Width = 108;
            // 
            // Column12
            // 
            this.Column12.DataPropertyName = "col_Safra";
            this.Column12.HeaderText = "SAFRA";
            this.Column12.Name = "Column12";
            // 
            // Column13
            // 
            this.Column13.DataPropertyName = "id_Debito";
            this.Column13.HeaderText = "Column13";
            this.Column13.Name = "Column13";
            this.Column13.Visible = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rdSafra2021);
            this.groupBox2.Controls.Add(this.rdSafra2020);
            this.groupBox2.Location = new System.Drawing.Point(414, 516);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(154, 168);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Filtro";
            // 
            // rdSafra2021
            // 
            this.rdSafra2021.AutoSize = true;
            this.rdSafra2021.Location = new System.Drawing.Point(34, 89);
            this.rdSafra2021.Name = "rdSafra2021";
            this.rdSafra2021.Size = new System.Drawing.Size(78, 17);
            this.rdSafra2021.TabIndex = 1;
            this.rdSafra2021.Text = "2021/2022";
            this.rdSafra2021.UseVisualStyleBackColor = true;
            this.rdSafra2021.Click += new System.EventHandler(this.rdSafra2021_Click);
            // 
            // rdSafra2020
            // 
            this.rdSafra2020.AutoSize = true;
            this.rdSafra2020.Checked = true;
            this.rdSafra2020.Location = new System.Drawing.Point(34, 63);
            this.rdSafra2020.Name = "rdSafra2020";
            this.rdSafra2020.Size = new System.Drawing.Size(78, 17);
            this.rdSafra2020.TabIndex = 0;
            this.rdSafra2020.TabStop = true;
            this.rdSafra2020.Text = "2020/2021";
            this.rdSafra2020.UseVisualStyleBackColor = true;
            this.rdSafra2020.CheckedChanged += new System.EventHandler(this.rdSafra2020_CheckedChanged);
            this.rdSafra2020.Click += new System.EventHandler(this.rdSafra2020_Click);
            // 
            // btnVoltar
            // 
            this.btnVoltar.Location = new System.Drawing.Point(1301, 661);
            this.btnVoltar.Name = "btnVoltar";
            this.btnVoltar.Size = new System.Drawing.Size(75, 23);
            this.btnVoltar.TabIndex = 6;
            this.btnVoltar.Text = "&Voltar";
            this.btnVoltar.UseVisualStyleBackColor = true;
            this.btnVoltar.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBoxCodigo
            // 
            this.textBoxCodigo.Location = new System.Drawing.Point(717, 562);
            this.textBoxCodigo.Name = "textBoxCodigo";
            this.textBoxCodigo.Size = new System.Drawing.Size(100, 20);
            this.textBoxCodigo.TabIndex = 7;
            this.textBoxCodigo.Visible = false;
            // 
            // textBoxDesc
            // 
            this.textBoxDesc.Location = new System.Drawing.Point(717, 587);
            this.textBoxDesc.Name = "textBoxDesc";
            this.textBoxDesc.Size = new System.Drawing.Size(100, 20);
            this.textBoxDesc.TabIndex = 8;
            this.textBoxDesc.Visible = false;
            // 
            // txtSafraOu
            // 
            this.txtSafraOu.Location = new System.Drawing.Point(717, 613);
            this.txtSafraOu.Name = "txtSafraOu";
            this.txtSafraOu.Size = new System.Drawing.Size(100, 20);
            this.txtSafraOu.TabIndex = 10;
            this.txtSafraOu.Visible = false;
            // 
            // txtID
            // 
            this.txtID.Location = new System.Drawing.Point(717, 639);
            this.txtID.Name = "txtID";
            this.txtID.Size = new System.Drawing.Size(100, 20);
            this.txtID.TabIndex = 11;
            this.txtID.Visible = false;
            // 
            // txtDescontado
            // 
            this.txtDescontado.Enabled = false;
            this.txtDescontado.Location = new System.Drawing.Point(1287, 516);
            this.txtDescontado.Name = "txtDescontado";
            this.txtDescontado.Size = new System.Drawing.Size(100, 20);
            this.txtDescontado.TabIndex = 12;
            this.txtDescontado.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtSaldo
            // 
            this.txtSaldo.Enabled = false;
            this.txtSaldo.Location = new System.Drawing.Point(1287, 542);
            this.txtSaldo.Name = "txtSaldo";
            this.txtSaldo.Size = new System.Drawing.Size(100, 20);
            this.txtSaldo.TabIndex = 13;
            this.txtSaldo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(1202, 519);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(79, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "Já Descontado";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(1201, 545);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(80, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Saldo Restante";
            // 
            // txtValorDTdesconto
            // 
            this.txtValorDTdesconto.Location = new System.Drawing.Point(823, 562);
            this.txtValorDTdesconto.Name = "txtValorDTdesconto";
            this.txtValorDTdesconto.Size = new System.Drawing.Size(100, 20);
            this.txtValorDTdesconto.TabIndex = 16;
            this.txtValorDTdesconto.Visible = false;
            // 
            // txtDivida
            // 
            this.txtDivida.Location = new System.Drawing.Point(286, 86);
            this.txtDivida.Name = "txtDivida";
            this.txtDivida.Size = new System.Drawing.Size(100, 20);
            this.txtDivida.TabIndex = 11;
            this.txtDivida.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // FormDesconto
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1399, 696);
            this.Controls.Add(this.txtValorDTdesconto);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtSaldo);
            this.Controls.Add(this.txtDescontado);
            this.Controls.Add(this.txtID);
            this.Controls.Add(this.txtSafraOu);
            this.Controls.Add(this.textBoxDesc);
            this.Controls.Add(this.textBoxCodigo);
            this.Controls.Add(this.btnVoltar);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.dtDetalhe);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dtDebitos);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(1399, 696);
            this.MinimumSize = new System.Drawing.Size(1399, 696);
            this.Name = "FormDesconto";
            this.Style = MetroFramework.MetroColorStyle.Orange;
            this.Text = "Débitos de Fornecedores";
            this.Load += new System.EventHandler(this.DescontoForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dtDebitos)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtDetalhe)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dtDebitos;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cmbDesc;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtCodFornecedor;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtQuant;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DataGridView dtDetalhe;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox txtUnidade;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dataDesconto;
        private System.Windows.Forms.TextBox txtSafra;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton rdSafra2021;
        private System.Windows.Forms.RadioButton rdSafra2020;
        private System.Windows.Forms.Button btnVoltar;
        private System.Windows.Forms.TextBox textBoxCodigo;
        private System.Windows.Forms.TextBox textBoxDesc;
        private System.Windows.Forms.TextBox txtSafraOu;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn CÓDIGO;
        private System.Windows.Forms.DataGridViewTextBoxColumn NOME;
        private System.Windows.Forms.DataGridViewTextBoxColumn DATA;
        private System.Windows.Forms.DataGridViewTextBoxColumn TIPO;
        private System.Windows.Forms.DataGridViewTextBoxColumn DESCRICAO;
        private System.Windows.Forms.DataGridViewTextBoxColumn QUANTI;
        private System.Windows.Forms.DataGridViewTextBoxColumn UNID;
        private System.Windows.Forms.DataGridViewTextBoxColumn VALORDIV;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column8;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column9;
        private textValor txtDivida;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtID;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column7;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column10;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column11;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column12;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column13;
        private System.Windows.Forms.TextBox txtDescontado;
        private System.Windows.Forms.TextBox txtSaldo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtValorDTdesconto;
    }
}