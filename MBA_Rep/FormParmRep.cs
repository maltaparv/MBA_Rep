using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MBA_Rep
{
    public partial class FormParmRep : Form
    {
        public FormParmRep(string sRepName)
        {
            InitializeComponent();
            //this.LblRepName.Text = sRepName;
            this.RtbRepName.Text = sRepName;

            DtpDat1.Text = ParmRep.Dat1.ToString("dd.MM.yyyy");
            DtpDat2.Text = ParmRep.Dat2.ToString("dd.MM.yyyy");
            Lbl_Period1.Visible = ParmRep.IsPeriod;
            Lbl_Period2.Visible = ParmRep.IsPeriod;
            DtpDat1.Visible = ParmRep.IsPeriod;
            DtpDat2.Visible = ParmRep.IsPeriod;
            //TxtCond0.Visible = ParmRep.IsCond[0];
            TxtCond1.Visible = ParmRep.IsCond[1];
            TxtCond2.Visible = ParmRep.IsCond[2];
            TxtCond3.Visible = ParmRep.IsCond[3];

            //LblCond0.Visible = ParmRep.IsCond[0];
            LblCond1.Visible = ParmRep.IsCond[1];
            LblCond2.Visible = ParmRep.IsCond[2];
             LblCond3.Visible = ParmRep.IsCond[3];
            //LblCond0.Text = ParmRep.NamCond[0];
            LblCond1.Text = ParmRep.NamCond[1];
            LblCond2.Text = ParmRep.NamCond[2];
            LblCond3.Text = ParmRep.NamCond[3];
        }
        //string sRepName;

        private void BtnRun_Click(object sender, EventArgs e)
        {
            //заполнение введённых параметров отчёта
            ParmRep.Dat1 = Convert.ToDateTime(DtpDat1.Text);
            ParmRep.Dat2 = Convert.ToDateTime(DtpDat2.Text);        //Convert.ToDateTime(TxtDat2.Text);
            //ParmRep.Cond[0] = TxtCond0.Text;
            ParmRep.Cond[1] = TxtCond1.Text;
            ParmRep.Cond[2] = TxtCond2.Text;
            ParmRep.Cond[3] = TxtCond3.Text;
            int ii = 1;
            ii++;
            this.Close();
        }

        private void FormParmRep_KeyUp(object sender, KeyEventArgs e)  
            // обработка перехода по Enter к следующему полю ввода (+на форме установить KeyPreview=True) //2020-02-19
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                SelectNextControl(ActiveControl, true, true, true, true);
            }
        }
    }
}
