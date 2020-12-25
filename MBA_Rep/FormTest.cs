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
    public partial class FormTest : Form
    {
        public FormTest()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
             MessageBox.Show(" Нажата кнопка\n в тестовом окне.", 
                 "  Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
        }

        private void FormTest_Load(object sender, EventArgs e)
        {
            /*
            MessageBox.Show($" Событие от {sender} \n {e}",
                "  тест ...", MessageBoxButtons.OK, MessageBoxIcon.Information);
            */
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            MessageBox.Show($" Событие от {sender} \n {e}",
               "  тест ...", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void FormTest_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                SelectNextControl(ActiveControl, true, true, true, true);
            }
        }
    }
}
