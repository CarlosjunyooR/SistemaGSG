namespace SistemaGSG
{
    partial class ViewNf
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ViewNf));
            this.label2 = new System.Windows.Forms.Label();
            this.dataemissao = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.nfe = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.mes_nf = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.datavencimento = new System.Windows.Forms.DateTimePicker();
            this.label5 = new System.Windows.Forms.Label();
            this.cod_unico = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtFaz = new System.Windows.Forms.TextBox();
            this.vl_boleto = new SistemaGSG.textValor();
            this.vl_fecoep = new SistemaGSG.textValor();
            this.axAcroPDF1 = new AxAcroPDFLib.AxAcroPDF();
            this.btnClose = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF1)).BeginInit();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(59, 120);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Código Único";
            // 
            // dataemissao
            // 
            this.dataemissao.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.dataemissao.CustomFormat = "yyyy-MM-dd";
            this.dataemissao.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dataemissao.Location = new System.Drawing.Point(166, 217);
            this.dataemissao.Name = "dataemissao";
            this.dataemissao.Size = new System.Drawing.Size(101, 20);
            this.dataemissao.TabIndex = 23;
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(59, 297);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(31, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Valor";
            // 
            // label8
            // 
            this.label8.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(58, 223);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(87, 13);
            this.label8.TabIndex = 13;
            this.label8.Text = "Data de Emissão";
            // 
            // nfe
            // 
            this.nfe.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.nfe.Location = new System.Drawing.Point(167, 165);
            this.nfe.Name = "nfe";
            this.nfe.Size = new System.Drawing.Size(100, 20);
            this.nfe.TabIndex = 20;
            // 
            // label9
            // 
            this.label9.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(59, 168);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(42, 13);
            this.label9.TabIndex = 14;
            this.label9.Text = "Nº Nf-e";
            // 
            // mes_nf
            // 
            this.mes_nf.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.mes_nf.Location = new System.Drawing.Point(167, 191);
            this.mes_nf.Name = "mes_nf";
            this.mes_nf.Size = new System.Drawing.Size(50, 20);
            this.mes_nf.TabIndex = 21;
            // 
            // label14
            // 
            this.label14.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(59, 271);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(43, 13);
            this.label14.TabIndex = 24;
            this.label14.Text = "Fecoep";
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(58, 248);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(104, 13);
            this.label3.TabIndex = 15;
            this.label3.Text = "Data de Vencimento";
            // 
            // datavencimento
            // 
            this.datavencimento.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.datavencimento.CustomFormat = "yyyy-MM-dd";
            this.datavencimento.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.datavencimento.Location = new System.Drawing.Point(167, 242);
            this.datavencimento.Name = "datavencimento";
            this.datavencimento.Size = new System.Drawing.Size(100, 20);
            this.datavencimento.TabIndex = 22;
            // 
            // label5
            // 
            this.label5.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(59, 194);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(72, 13);
            this.label5.TabIndex = 16;
            this.label5.Text = "Mês Faturado";
            // 
            // cod_unico
            // 
            this.cod_unico.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.cod_unico.AutoCompleteCustomSource.AddRange(new string[] {
            "0147906-7",
            "0147907-5",
            "0147908-3",
            "0147910-5",
            "0148056-1",
            "0384326-2",
            "0386058-2",
            "0386074-4",
            "0386075-2",
            "0386079-5",
            "0386083-3",
            "0386087-6",
            "0386090-6",
            "0387602-0",
            "0387603-9",
            "0387604-7",
            "0387611-0",
            "0400485-0",
            "0412233-0",
            "0412235-6",
            "0412237-2",
            "0412238-0",
            "0412239-9",
            "0412242-9",
            "0412243-7",
            "0412244-5",
            "0412245-3",
            "0412246-1",
            "0412247-0",
            "0412248-8",
            "0412249-6",
            "0412250-0",
            "0412252-6",
            "0412253-4",
            "0412254-2",
            "0412255-0",
            "0412256-9",
            "0412257-7",
            "0412258-5",
            "0412259-3",
            "0412260-7",
            "0412261-5",
            "0412263-1",
            "0412264-0",
            "0412265-8",
            "0412266-6",
            "0412268-2",
            "0412269-0",
            "0412270-4",
            "0412271-2",
            "0412272-0",
            "0412273-9",
            "0412274-7",
            "0412275-5",
            "0412276-3",
            "0412277-1",
            "0412278-0",
            "0412279-8",
            "0412280-1",
            "0412281-0",
            "0421011-5",
            "0421012-3",
            "0428979-0",
            "0428981-1",
            "0473505-6",
            "0502213-4",
            "0511133-1",
            "0531222-1",
            "0531233-7",
            "0559927-0",
            "0559933-4",
            "0559940-7",
            "0559942-3",
            "0559943-1",
            "0559944-0",
            "0559945-8",
            "0559946-6",
            "0559947-4",
            "0559949-0",
            "0559950-4",
            "0559951-2",
            "0559953-9",
            "0559954-7",
            "0559955-5",
            "0559959-8",
            "0559962-8",
            "0559964-4",
            "0559966-0",
            "0559970-9",
            "0559971-7",
            "0559972-5",
            "0559973-3",
            "0559974-1",
            "0559975-0",
            "0559976-8",
            "0559977-6",
            "0559978-4",
            "0559979-2",
            "0559980-6",
            "0559981-4",
            "0559982-2",
            "0559983-0",
            "0568045-0",
            "0568048-4",
            "0571315-3",
            "0575921-8",
            "0577915-4",
            "0598098-4",
            "0615456-5",
            "0621541-6",
            "0647393-8",
            "0767723-5",
            "0767731-6",
            "0767733-2",
            "0772672-4",
            "0774160-0",
            "0774161-8",
            "0774162-6",
            "0774163-4",
            "0774164-2",
            "0774165-0",
            "0774166-9",
            "0774167-7",
            "0779218-2",
            "0801715-8",
            "0926005-6",
            "0926015-3",
            "0926017-0",
            "0926020-0",
            "0926033-1",
            "0926037-4",
            "0926043-9",
            "0926057-9",
            "0926069-2",
            "0972290-4",
            "1103579-0",
            "1136474-2",
            "1136478-5",
            "1136480-7",
            "1136482-3",
            "1136485-8",
            "1136486-6",
            "1136488-2",
            "1136490-4",
            "1136493-9",
            "1136495-5",
            "1136497-1",
            "1136498-0",
            "1136503-0",
            "1136504-8",
            "1136513-7",
            "1136516-1",
            "1136519-6",
            "1136521-8",
            "1137038-6",
            "1137041-6",
            "1175395-1",
            "1175398-6",
            "1175400-1",
            "1175402-8",
            "1175405-2",
            "1175406-0",
            "1175407-9",
            "1175408-7",
            "1175411-7",
            "1175412-5",
            "1175413-3",
            "1175415-0",
            "1175416-8",
            "1175418-4",
            "1175420-6",
            "1175422-2",
            "1175423-0",
            "1175424-9",
            "1175425-7",
            "1175426-5",
            "1175450-8",
            "1175453-2",
            "1175458-3",
            "1175470-2",
            "1175473-7",
            "1175476-1",
            "1175478-8",
            "1175479-6",
            "1175480-0",
            "1175481-8",
            "1175482-6",
            "1175483-4",
            "1175485-0",
            "1175491-5",
            "1175494-0",
            "1175498-2",
            "1175500-8",
            "1175504-0",
            "1175507-5",
            "1175510-5",
            "1304836-8",
            "1304839-2",
            "1304840-6",
            "1305227-6",
            "1412487-4",
            "1436520-0",
            "1436535-9",
            "1436540-5",
            "1436544-8",
            "1436550-2",
            "1436554-5",
            "1436557-0",
            "1436562-6",
            "1436564-2",
            "1436618-5",
            "1436619-3",
            "1436620-7",
            "1436622-3",
            "1436623-1",
            "1436624-0",
            "1436625-8",
            "1438603-8",
            "1438619-4",
            "1438656-9",
            "1438700-0",
            "1438714-0",
            "1438768-9",
            "1438783-2",
            "1438798-0",
            "1439351-4",
            "1439374-3",
            "1457087-4",
            "1457153-6",
            "1457166-8",
            "1457180-3",
            "1457201-0",
            "1457248-6",
            "1457306-7",
            "1457417-9",
            "1457428-4",
            "1457580-9",
            "1457583-3",
            "1457604-0",
            "1457614-7",
            "1457649-0",
            "1457696-1",
            "1457720-8",
            "1457733-0",
            "1457734-8",
            "1457735-6",
            "1457746-1",
            "1457821-2",
            "1457831-0",
            "1457839-5",
            "1457843-3",
            "1459987-2",
            "1459994-5",
            "1459998-8",
            "1460003-0",
            "1460086-2",
            "1460101-0",
            "1460104-4",
            "1460253-9",
            "1460279-2",
            "1460281-4",
            "1460283-0",
            "1460285-7",
            "1473031-6",
            "1473036-7",
            "1473041-3",
            "1473042-1",
            "1473103-7",
            "1473117-7",
            "1473125-8",
            "1473138-0",
            "1473262-9",
            "1473263-7",
            "1473271-8",
            "1482135-4",
            "1482174-5",
            "1482179-6",
            "1482185-0",
            "1482188-5",
            "1482204-0",
            "1482215-6",
            "1482219-9",
            "1482291-1",
            "1482318-7",
            "1482368-3",
            "1482384-5",
            "1482387-0",
            "1482793-0",
            "1482804-9",
            "1482811-1",
            "1482815-4",
            "1482819-7",
            "1482823-5",
            "1482910-0",
            "1482926-6",
            "1482930-4",
            "1482934-7",
            "1482938-0",
            "1482943-6",
            "1489889-6",
            "1489898-5",
            "1489903-5",
            "1489913-2",
            "1489920-5",
            "1489925-6",
            "1489932-9",
            "1489938-8",
            "1489942-6",
            "1489948-5",
            "1489958-2",
            "1489973-6",
            "1489978-7",
            "1489984-1",
            "1489995-7",
            "1489997-3",
            "1490003-3",
            "1490006-8",
            "1490011-4",
            "1490038-6",
            "1490058-0",
            "1490069-6",
            "1490072-6"});
            this.cod_unico.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cod_unico.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.cod_unico.FormattingEnabled = true;
            this.cod_unico.ItemHeight = 13;
            this.cod_unico.Location = new System.Drawing.Point(167, 112);
            this.cod_unico.Name = "cod_unico";
            this.cod_unico.Size = new System.Drawing.Size(100, 21);
            this.cod_unico.TabIndex = 18;
            // 
            // label6
            // 
            this.label6.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(59, 142);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(48, 13);
            this.label6.TabIndex = 17;
            this.label6.Text = "Fazenda";
            // 
            // txtFaz
            // 
            this.txtFaz.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtFaz.AutoCompleteCustomSource.AddRange(new string[] {
            "Fz Aliança",
            "Fz Alto do Guerra",
            "Fz Barro Branco",
            "Fz Belo Horizonte",
            "Fz Bom Retiro",
            "Fz Cajueiro",
            "Fz Campo Alegre",
            "Fz Campos",
            "Fz Campos II",
            "Fz Fundao",
            "Fz Ideal",
            "Fz Mulungu",
            "Fz Petropolis",
            "Fz Retiro",
            "Fz Riacho Seco",
            "Fz Santa Rita",
            "Fz Triunfo",
            "Fz Valparaizo",
            "Fz Valparaizo I",
            "Fz Valparaizo III",
            "Lt Rocha Cavalcante",
            "Pc Cel Carlos Lyra",
            "Pv Apolinario",
            "Pv Bambus",
            "Pv Brejo",
            "Pv Cajueiro",
            "Pv Cajueiro de Baixo",
            "Pv Camaratuba",
            "Pv Cambui",
            "Pv Campo Novo e Ideal",
            "Pv Campo Novo Horizonte",
            "Pv Canivete",
            "Pv Catole",
            "Pv Cuscuz",
            "Pv Novo Horizonte",
            "Pv Olho Dagua",
            "Pv Oriental",
            "Pv Pitomba",
            "Pv Retiro",
            "Pv Riachinho",
            "Pv Rocadinho",
            "Pv Sitio Retiro",
            "Pv Valparaizo",
            "Pv Varzea Bonita",
            "R Cel Carlos Lyra",
            "R do Campo",
            "R Nova",
            "St Alto do Guerra",
            "St Baixo Guzera",
            "St Coimbra",
            "St Portugues",
            "St Potugues"});
            this.txtFaz.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.txtFaz.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.txtFaz.Location = new System.Drawing.Point(167, 139);
            this.txtFaz.Name = "txtFaz";
            this.txtFaz.Size = new System.Drawing.Size(100, 20);
            this.txtFaz.TabIndex = 19;
            // 
            // vl_boleto
            // 
            this.vl_boleto.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.vl_boleto.Location = new System.Drawing.Point(167, 294);
            this.vl_boleto.Name = "vl_boleto";
            this.vl_boleto.Size = new System.Drawing.Size(100, 20);
            this.vl_boleto.TabIndex = 26;
            this.vl_boleto.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // vl_fecoep
            // 
            this.vl_fecoep.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.vl_fecoep.Location = new System.Drawing.Point(167, 268);
            this.vl_fecoep.Name = "vl_fecoep";
            this.vl_fecoep.Size = new System.Drawing.Size(100, 20);
            this.vl_fecoep.TabIndex = 25;
            this.vl_fecoep.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // axAcroPDF1
            // 
            this.axAcroPDF1.Enabled = true;
            this.axAcroPDF1.Location = new System.Drawing.Point(346, 24);
            this.axAcroPDF1.Name = "axAcroPDF1";
            this.axAcroPDF1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axAcroPDF1.OcxState")));
            this.axAcroPDF1.Size = new System.Drawing.Size(326, 403);
            this.axAcroPDF1.TabIndex = 27;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(265, 404);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 28;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // ViewNf
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(699, 450);
            this.ControlBox = false;
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.axAcroPDF1);
            this.Controls.Add(this.vl_boleto);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataemissao);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.nfe);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.vl_fecoep);
            this.Controls.Add(this.mes_nf);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.datavencimento);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cod_unico);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtFaz);
            this.Name = "ViewNf";
            this.Text = "Visualizar NF-e";
            this.Load += new System.EventHandler(this.ViewNf_Load);
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private textValor vl_boleto;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dataemissao;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox nfe;
        private System.Windows.Forms.Label label9;
        private textValor vl_fecoep;
        private System.Windows.Forms.TextBox mes_nf;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker datavencimento;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cod_unico;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtFaz;
        private AxAcroPDFLib.AxAcroPDF axAcroPDF1;
        private System.Windows.Forms.Button btnClose;
    }
}