using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class FormRelat : Form
    {
        public FormRelat()
        {
            InitializeComponent();
        }

        private void FormRelat_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dBSGSGDataSetSemanaSaida.DBSGSG_SaidaSemana' table. You can move, or remove it, as needed.
            this.dBSGSG_SaidaSemanaTableAdapter.Fill(this.dBSGSGDataSetSemanaSaida.DBSGSG_SaidaSemana);
            //this.reportViewer1.RefreshReport();
            //using(DBSGSGDataSetSemanaSaida db = new DBSGSGDataSetSemanaSaida())
            //{
            //    dBSGSGDataSetSemanaSaidaBindingSource.DataSource = db.;
            //}
            ReportParameterCollection reportParameters = new ReportParameterCollection();
            reportParameters.Add(new ReportParameter("fromDate", dtFrom.Text));
            reportParameters.Add(new ReportParameter("toDate", dtToDate.Text));
            this.reportViewer1.LocalReport.SetParameters(reportParameters);
            this.reportViewer1.RefreshReport();
        }
    }
}
