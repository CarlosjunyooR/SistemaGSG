namespace SistemaGSG
{
    partial class frmProtocolo
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmProtocolo));
            this.webBrowser = new System.Windows.Forms.WebBrowser();
            this.txtNSU = new System.Windows.Forms.TextBox();
            this.dataGridViewSefaz = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBox = new System.Windows.Forms.CheckBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.txtChave = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.TempoPesquisa = new System.Windows.Forms.Timer(this.components);
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.txtResultNSU = new System.Windows.Forms.TextBox();
            this.dtXML = new System.Windows.Forms.DataGridView();
            this.txtUrl = new System.Windows.Forms.TextBox();
            this.txtEmpresa = new System.Windows.Forms.TextBox();
            this.txtCNPJ = new System.Windows.Forms.TextBox();
            this.txtURLxml = new System.Windows.Forms.TextBox();
            this.NFEGRID = new System.Windows.Forms.DataGridView();
            this.txtNFE = new System.Windows.Forms.TextBox();
            this.txtserie = new System.Windows.Forms.TextBox();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.txtdate = new System.Windows.Forms.TextBox();
            this.DADOSGRID = new System.Windows.Forms.DataGridView();
            this.txtDados = new System.Windows.Forms.TextBox();
            this.CHAVEGRID = new System.Windows.Forms.DataGridView();
            this.txtChavedeAcesso = new System.Windows.Forms.TextBox();
            this.txtProtocolo = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.ProgBar = new System.Windows.Forms.ProgressBar();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.lblPorcentagem = new System.Windows.Forms.Label();
            this.dataGridViewRestante = new System.Windows.Forms.DataGridView();
            this.label11 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.TempoEspera = new System.Windows.Forms.Timer(this.components);
            this.label12 = new System.Windows.Forms.Label();
            this.lblChaveDuplicidade = new System.Windows.Forms.Label();
            this.tpNF = new System.Windows.Forms.TextBox();
            this.dataGridViewProdutos = new System.Windows.Forms.DataGridView();
            this.maskFiltro = new System.Windows.Forms.MaskedTextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.textValor1 = new SistemaGSG.textValor();
            this.vNF = new SistemaGSG.textValor();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column13 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column14 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column11 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column12 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column15 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column16 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button5 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSefaz)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtXML)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.NFEGRID)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DADOSGRID)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CHAVEGRID)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewRestante)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewProdutos)).BeginInit();
            this.SuspendLayout();
            // 
            // webBrowser
            // 
            this.webBrowser.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.webBrowser.Location = new System.Drawing.Point(15, 760);
            this.webBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.Size = new System.Drawing.Size(879, 384);
            this.webBrowser.TabIndex = 0;
            this.webBrowser.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser_DocumentCompleted_1);
            // 
            // txtNSU
            // 
            this.txtNSU.Location = new System.Drawing.Point(2325, 423);
            this.txtNSU.Name = "txtNSU";
            this.txtNSU.Size = new System.Drawing.Size(152, 20);
            this.txtNSU.TabIndex = 1;
            this.txtNSU.Visible = false;
            // 
            // dataGridViewSefaz
            // 
            this.dataGridViewSefaz.AllowUserToAddRows = false;
            this.dataGridViewSefaz.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewSefaz.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewSefaz.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column2,
            this.Column3,
            this.Column4,
            this.Column5,
            this.Column6,
            this.Column13,
            this.Column14,
            this.Column7,
            this.Column8,
            this.Column9,
            this.Column10,
            this.Column11,
            this.Column12,
            this.Column15,
            this.Column16,
            this.Column1});
            this.dataGridViewSefaz.Location = new System.Drawing.Point(15, 82);
            this.dataGridViewSefaz.Name = "dataGridViewSefaz";
            this.dataGridViewSefaz.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewSefaz.Size = new System.Drawing.Size(1604, 564);
            this.dataGridViewSefaz.TabIndex = 3;
            this.dataGridViewSefaz.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewSefaz_CellClick);
            this.dataGridViewSefaz.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridViewSefaz_CellFormatting);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.maskFiltro);
            this.groupBox1.Controls.Add(this.checkBox);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 20);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1607, 56);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Informações do Destinatário";
            // 
            // checkBox
            // 
            this.checkBox.AutoSize = true;
            this.checkBox.Location = new System.Drawing.Point(228, 29);
            this.checkBox.Name = "checkBox";
            this.checkBox.Size = new System.Drawing.Size(82, 17);
            this.checkBox.TabIndex = 12;
            this.checkBox.Text = "Ativar Timer";
            this.checkBox.UseVisualStyleBackColor = true;
            this.checkBox.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // textBox2
            // 
            this.textBox2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.textBox2.Enabled = false;
            this.textBox2.Location = new System.Drawing.Point(1488, 23);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(113, 20);
            this.textBox2.TabIndex = 3;
            this.textBox2.Text = "12.706.289/0001-48";
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(1431, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(37, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "CNPJ:";
            // 
            // textBox1
            // 
            this.textBox1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(63, 23);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(159, 20);
            this.textBox1.TabIndex = 1;
            this.textBox1.Text = "USINA SERRA GRANDE SA";
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Empresa:";
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(1521, 699);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(98, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "&Baixar XML";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtChave
            // 
            this.txtChave.Location = new System.Drawing.Point(2289, 438);
            this.txtChave.Name = "txtChave";
            this.txtChave.Size = new System.Drawing.Size(373, 20);
            this.txtChave.TabIndex = 6;
            this.txtChave.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(2276, 426);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(30, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "NSU";
            this.label3.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(2240, 441);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(43, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "CHAVE";
            this.label4.Visible = false;
            // 
            // TempoPesquisa
            // 
            this.TempoPesquisa.Interval = 7000;
            this.TempoPesquisa.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.Location = new System.Drawing.Point(1424, 699);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(91, 23);
            this.button2.TabIndex = 9;
            this.button2.Text = "&Carregar XML";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Location = new System.Drawing.Point(1164, 699);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 10;
            this.button3.Text = "&Voltar";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // txtResultNSU
            // 
            this.txtResultNSU.Location = new System.Drawing.Point(2270, 423);
            this.txtResultNSU.Name = "txtResultNSU";
            this.txtResultNSU.Size = new System.Drawing.Size(100, 20);
            this.txtResultNSU.TabIndex = 11;
            this.txtResultNSU.Visible = false;
            // 
            // dtXML
            // 
            this.dtXML.AllowUserToAddRows = false;
            this.dtXML.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtXML.Location = new System.Drawing.Point(1943, 651);
            this.dtXML.Name = "dtXML";
            this.dtXML.Size = new System.Drawing.Size(240, 105);
            this.dtXML.TabIndex = 12;
            this.dtXML.Visible = false;
            // 
            // txtUrl
            // 
            this.txtUrl.Location = new System.Drawing.Point(2368, 503);
            this.txtUrl.Name = "txtUrl";
            this.txtUrl.Size = new System.Drawing.Size(138, 20);
            this.txtUrl.TabIndex = 13;
            this.txtUrl.Text = "C:\\ArquivosSAP\\XML\\";
            this.txtUrl.Visible = false;
            // 
            // txtEmpresa
            // 
            this.txtEmpresa.Location = new System.Drawing.Point(2227, 503);
            this.txtEmpresa.Name = "txtEmpresa";
            this.txtEmpresa.Size = new System.Drawing.Size(59, 20);
            this.txtEmpresa.TabIndex = 14;
            this.txtEmpresa.Visible = false;
            // 
            // txtCNPJ
            // 
            this.txtCNPJ.Location = new System.Drawing.Point(2227, 529);
            this.txtCNPJ.Name = "txtCNPJ";
            this.txtCNPJ.Size = new System.Drawing.Size(59, 20);
            this.txtCNPJ.TabIndex = 15;
            this.txtCNPJ.Visible = false;
            // 
            // txtURLxml
            // 
            this.txtURLxml.Location = new System.Drawing.Point(2368, 581);
            this.txtURLxml.Name = "txtURLxml";
            this.txtURLxml.Size = new System.Drawing.Size(702, 20);
            this.txtURLxml.TabIndex = 16;
            this.txtURLxml.Text = "NOTAFISCAL.xml";
            this.txtURLxml.Visible = false;
            // 
            // NFEGRID
            // 
            this.NFEGRID.AllowUserToAddRows = false;
            this.NFEGRID.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.NFEGRID.Location = new System.Drawing.Point(2364, 372);
            this.NFEGRID.Name = "NFEGRID";
            this.NFEGRID.Size = new System.Drawing.Size(240, 150);
            this.NFEGRID.TabIndex = 17;
            this.NFEGRID.Visible = false;
            // 
            // txtNFE
            // 
            this.txtNFE.Location = new System.Drawing.Point(2292, 503);
            this.txtNFE.Name = "txtNFE";
            this.txtNFE.Size = new System.Drawing.Size(100, 20);
            this.txtNFE.TabIndex = 18;
            this.txtNFE.Visible = false;
            // 
            // txtserie
            // 
            this.txtserie.Location = new System.Drawing.Point(2292, 529);
            this.txtserie.Name = "txtserie";
            this.txtserie.Size = new System.Drawing.Size(100, 20);
            this.txtserie.TabIndex = 19;
            this.txtserie.Visible = false;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(2398, 503);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker1.TabIndex = 20;
            this.dateTimePicker1.Visible = false;
            // 
            // txtdate
            // 
            this.txtdate.Location = new System.Drawing.Point(2398, 529);
            this.txtdate.Name = "txtdate";
            this.txtdate.Size = new System.Drawing.Size(100, 20);
            this.txtdate.TabIndex = 21;
            this.txtdate.Visible = false;
            // 
            // DADOSGRID
            // 
            this.DADOSGRID.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DADOSGRID.Location = new System.Drawing.Point(1871, 80);
            this.DADOSGRID.Name = "DADOSGRID";
            this.DADOSGRID.Size = new System.Drawing.Size(240, 150);
            this.DADOSGRID.TabIndex = 22;
            this.DADOSGRID.Visible = false;
            // 
            // txtDados
            // 
            this.txtDados.Location = new System.Drawing.Point(2242, 625);
            this.txtDados.Name = "txtDados";
            this.txtDados.Size = new System.Drawing.Size(94, 20);
            this.txtDados.TabIndex = 23;
            this.txtDados.Visible = false;
            // 
            // CHAVEGRID
            // 
            this.CHAVEGRID.AllowUserToAddRows = false;
            this.CHAVEGRID.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.CHAVEGRID.Location = new System.Drawing.Point(2368, 528);
            this.CHAVEGRID.Name = "CHAVEGRID";
            this.CHAVEGRID.Size = new System.Drawing.Size(379, 150);
            this.CHAVEGRID.TabIndex = 24;
            this.CHAVEGRID.Visible = false;
            // 
            // txtChavedeAcesso
            // 
            this.txtChavedeAcesso.Location = new System.Drawing.Point(2368, 555);
            this.txtChavedeAcesso.Name = "txtChavedeAcesso";
            this.txtChavedeAcesso.Size = new System.Drawing.Size(100, 20);
            this.txtChavedeAcesso.TabIndex = 25;
            this.txtChavedeAcesso.Visible = false;
            // 
            // txtProtocolo
            // 
            this.txtProtocolo.Location = new System.Drawing.Point(2261, 555);
            this.txtProtocolo.Name = "txtProtocolo";
            this.txtProtocolo.Size = new System.Drawing.Size(97, 20);
            this.txtProtocolo.TabIndex = 26;
            this.txtProtocolo.Visible = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(23, 680);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(68, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Quant. NF-e:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(97, 680);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(0, 13);
            this.label6.TabIndex = 14;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(484, 680);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(0, 13);
            this.label7.TabIndex = 27;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(377, 680);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(101, 13);
            this.label8.TabIndex = 28;
            this.label8.Text = "Pasta dos Arquivos:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(151, 680);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(134, 13);
            this.label9.TabIndex = 29;
            this.label9.Text = "Notas Fiscais Consultadas:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(284, 680);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(0, 13);
            this.label10.TabIndex = 30;
            // 
            // button4
            // 
            this.button4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button4.Location = new System.Drawing.Point(1326, 699);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(92, 23);
            this.button4.TabIndex = 31;
            this.button4.Text = "&Consultar SAP";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // ProgBar
            // 
            this.ProgBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ProgBar.Location = new System.Drawing.Point(15, 652);
            this.ProgBar.Maximum = 0;
            this.ProgBar.Name = "ProgBar";
            this.ProgBar.Size = new System.Drawing.Size(1566, 19);
            this.ProgBar.TabIndex = 32;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker2.Location = new System.Drawing.Point(1662, 567);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker2.TabIndex = 33;
            // 
            // lblPorcentagem
            // 
            this.lblPorcentagem.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblPorcentagem.AutoSize = true;
            this.lblPorcentagem.Location = new System.Drawing.Point(1587, 655);
            this.lblPorcentagem.Name = "lblPorcentagem";
            this.lblPorcentagem.Size = new System.Drawing.Size(24, 13);
            this.lblPorcentagem.TabIndex = 34;
            this.lblPorcentagem.Text = "0 %";
            // 
            // dataGridViewRestante
            // 
            this.dataGridViewRestante.AllowUserToAddRows = false;
            this.dataGridViewRestante.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewRestante.Location = new System.Drawing.Point(1943, 339);
            this.dataGridViewRestante.Name = "dataGridViewRestante";
            this.dataGridViewRestante.Size = new System.Drawing.Size(240, 150);
            this.dataGridViewRestante.TabIndex = 35;
            this.dataGridViewRestante.Visible = false;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(2406, 601);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(41, 13);
            this.label11.TabIndex = 37;
            this.label11.Text = "label11";
            this.label11.Visible = false;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(2437, 529);
            this.textBox3.MaxLength = 3;
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(100, 20);
            this.textBox3.TabIndex = 38;
            this.textBox3.Visible = false;
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(2406, 684);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(402, 20);
            this.textBox4.TabIndex = 39;
            this.textBox4.Visible = false;
            // 
            // TempoEspera
            // 
            this.TempoEspera.Interval = 1800000;
            this.TempoEspera.Tick += new System.EventHandler(this.TempoEspera_Tick);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(21, 704);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(66, 13);
            this.label12.TabIndex = 40;
            this.label12.Text = "Duplicidade:";
            // 
            // lblChaveDuplicidade
            // 
            this.lblChaveDuplicidade.AutoSize = true;
            this.lblChaveDuplicidade.Location = new System.Drawing.Point(93, 704);
            this.lblChaveDuplicidade.Name = "lblChaveDuplicidade";
            this.lblChaveDuplicidade.Size = new System.Drawing.Size(0, 13);
            this.lblChaveDuplicidade.TabIndex = 41;
            // 
            // tpNF
            // 
            this.tpNF.Location = new System.Drawing.Point(2239, 395);
            this.tpNF.Name = "tpNF";
            this.tpNF.Size = new System.Drawing.Size(100, 20);
            this.tpNF.TabIndex = 42;
            // 
            // dataGridViewProdutos
            // 
            this.dataGridViewProdutos.AllowUserToAddRows = false;
            this.dataGridViewProdutos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewProdutos.Location = new System.Drawing.Point(1943, 495);
            this.dataGridViewProdutos.Name = "dataGridViewProdutos";
            this.dataGridViewProdutos.Size = new System.Drawing.Size(240, 150);
            this.dataGridViewProdutos.TabIndex = 43;
            // 
            // maskFiltro
            // 
            this.maskFiltro.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.maskFiltro.Location = new System.Drawing.Point(719, 28);
            this.maskFiltro.Mask = "00/0000";
            this.maskFiltro.Name = "maskFiltro";
            this.maskFiltro.Size = new System.Drawing.Size(56, 20);
            this.maskFiltro.TabIndex = 13;
            this.maskFiltro.ValidatingType = typeof(System.DateTime);
            this.maskFiltro.MaskChanged += new System.EventHandler(this.maskFiltro_MaskChanged);
            this.maskFiltro.KeyUp += new System.Windows.Forms.KeyEventHandler(this.maskFiltro_KeyUp);
            // 
            // label13
            // 
            this.label13.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(730, 12);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(32, 13);
            this.label13.TabIndex = 14;
            this.label13.Text = "Filtrar";
            // 
            // textValor1
            // 
            this.textValor1.Location = new System.Drawing.Point(1851, 590);
            this.textValor1.Name = "textValor1";
            this.textValor1.Size = new System.Drawing.Size(100, 20);
            this.textValor1.TabIndex = 47;
            this.textValor1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // vNF
            // 
            this.vNF.Location = new System.Drawing.Point(2239, 421);
            this.vNF.Name = "vNF";
            this.vNF.Size = new System.Drawing.Size(100, 20);
            this.vNF.TabIndex = 44;
            this.vNF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // Column2
            // 
            this.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column2.DataPropertyName = "col_chave";
            this.Column2.HeaderText = "CHAVE";
            this.Column2.Name = "Column2";
            this.Column2.Width = 68;
            // 
            // Column3
            // 
            this.Column3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column3.DataPropertyName = "col_nsu";
            this.Column3.HeaderText = "NSU";
            this.Column3.Name = "Column3";
            this.Column3.Width = 55;
            // 
            // Column4
            // 
            this.Column4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column4.DataPropertyName = "empresa";
            this.Column4.HeaderText = "EMPRESA";
            this.Column4.Name = "Column4";
            this.Column4.Width = 84;
            // 
            // Column5
            // 
            this.Column5.DataPropertyName = "n_nfe";
            this.Column5.HeaderText = "NFE";
            this.Column5.Name = "Column5";
            // 
            // Column6
            // 
            this.Column6.DataPropertyName = "emisao";
            this.Column6.HeaderText = "EMISSAO";
            this.Column6.Name = "Column6";
            // 
            // Column13
            // 
            this.Column13.DataPropertyName = "col_desc_NFe";
            this.Column13.HeaderText = "TIPO";
            this.Column13.Name = "Column13";
            // 
            // Column14
            // 
            this.Column14.DataPropertyName = "vNF";
            this.Column14.HeaderText = "VALOR";
            this.Column14.Name = "Column14";
            // 
            // Column7
            // 
            this.Column7.DataPropertyName = "lancamento_sap";
            this.Column7.HeaderText = "LANÇAMENTO";
            this.Column7.Name = "Column7";
            // 
            // Column8
            // 
            this.Column8.DataPropertyName = "protocolo";
            this.Column8.HeaderText = "PROTOCOLO";
            this.Column8.Name = "Column8";
            // 
            // Column9
            // 
            this.Column9.DataPropertyName = "user_sap";
            this.Column9.HeaderText = "USER SAP";
            this.Column9.Name = "Column9";
            // 
            // Column10
            // 
            this.Column10.DataPropertyName = "status";
            this.Column10.HeaderText = "STATUS";
            this.Column10.Name = "Column10";
            // 
            // Column11
            // 
            this.Column11.DataPropertyName = "col_Downl";
            this.Column11.HeaderText = "Column11";
            this.Column11.Name = "Column11";
            this.Column11.Visible = false;
            // 
            // Column12
            // 
            this.Column12.DataPropertyName = "col_link";
            this.Column12.HeaderText = "Column12";
            this.Column12.Name = "Column12";
            this.Column12.Visible = false;
            // 
            // Column15
            // 
            this.Column15.DataPropertyName = "tpNF";
            this.Column15.HeaderText = "Column15";
            this.Column15.Name = "Column15";
            this.Column15.Visible = false;
            // 
            // Column16
            // 
            this.Column16.DataPropertyName = "ACTION_REQU";
            this.Column16.HeaderText = "Column16";
            this.Column16.Name = "Column16";
            this.Column16.Visible = false;
            // 
            // Column1
            // 
            this.Column1.DataPropertyName = "ID";
            this.Column1.HeaderText = "Column1";
            this.Column1.Name = "Column1";
            this.Column1.Visible = false;
            // 
            // button5
            // 
            this.button5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button5.Location = new System.Drawing.Point(1245, 699);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 48;
            this.button5.Text = "&Geral";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // frmProtocolo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1623, 733);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.textValor1);
            this.Controls.Add(this.vNF);
            this.Controls.Add(this.dataGridViewProdutos);
            this.Controls.Add(this.tpNF);
            this.Controls.Add(this.lblChaveDuplicidade);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.dataGridViewRestante);
            this.Controls.Add(this.lblPorcentagem);
            this.Controls.Add(this.dateTimePicker2);
            this.Controls.Add(this.ProgBar);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtProtocolo);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtChavedeAcesso);
            this.Controls.Add(this.CHAVEGRID);
            this.Controls.Add(this.txtDados);
            this.Controls.Add(this.DADOSGRID);
            this.Controls.Add(this.txtdate);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.txtserie);
            this.Controls.Add(this.txtNFE);
            this.Controls.Add(this.NFEGRID);
            this.Controls.Add(this.txtURLxml);
            this.Controls.Add(this.txtCNPJ);
            this.Controls.Add(this.txtEmpresa);
            this.Controls.Add(this.txtUrl);
            this.Controls.Add(this.dtXML);
            this.Controls.Add(this.txtResultNSU);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtChave);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dataGridViewSefaz);
            this.Controls.Add(this.txtNSU);
            this.Controls.Add(this.webBrowser);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmProtocolo";
            this.Style = MetroFramework.MetroColorStyle.Orange;
            this.Text = "Consulta DF-e";
            this.Load += new System.EventHandler(this.frmProtocolo_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSefaz)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtXML)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.NFEGRID)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DADOSGRID)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CHAVEGRID)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewRestante)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewProdutos)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.WebBrowser webBrowser;
        private System.Windows.Forms.TextBox txtNSU;
        private System.Windows.Forms.DataGridView dataGridViewSefaz;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtChave;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Timer TempoPesquisa;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.CheckBox checkBox;
        private System.Windows.Forms.TextBox txtResultNSU;
        private System.Windows.Forms.DataGridView dtXML;
        private System.Windows.Forms.TextBox txtUrl;
        private System.Windows.Forms.TextBox txtEmpresa;
        private System.Windows.Forms.TextBox txtCNPJ;
        private System.Windows.Forms.TextBox txtURLxml;
        private System.Windows.Forms.DataGridView NFEGRID;
        private System.Windows.Forms.TextBox txtNFE;
        private System.Windows.Forms.TextBox txtserie;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.TextBox txtdate;
        private System.Windows.Forms.DataGridView DADOSGRID;
        private System.Windows.Forms.TextBox txtDados;
        private System.Windows.Forms.DataGridView CHAVEGRID;
        private System.Windows.Forms.TextBox txtChavedeAcesso;
        private System.Windows.Forms.TextBox txtProtocolo;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.ProgressBar ProgBar;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Label lblPorcentagem;
        private System.Windows.Forms.DataGridView dataGridViewRestante;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Timer TempoEspera;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label lblChaveDuplicidade;
        private System.Windows.Forms.TextBox tpNF;
        private System.Windows.Forms.DataGridView dataGridViewProdutos;
        private textValor vNF;
        private textValor textValor1;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.MaskedTextBox maskFiltro;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column13;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column14;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column7;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column8;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column9;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column10;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column11;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column12;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column15;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column16;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.Button button5;
    }
}