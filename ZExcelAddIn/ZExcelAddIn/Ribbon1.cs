using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using ZExcelAddIn.Properties;

namespace ZExcelAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            zoom1_5.Text = Properties.Settings.Default.ズーム率;
            cursor1_5.Text = Properties.Settings.Default.カーソル位置;
            editBox1_7.Text = Properties.Settings.Default.上記に続けて;
            editBox1_8.Text = Properties.Settings.Default.列を記憶;
        }

        private void button1_1_Click(object sender, RibbonControlEventArgs e)
        {
            ZMethodClass.delete_customviews();
        }

        private void button1_2_Click(object sender, RibbonControlEventArgs e)
        {
            ZMethodClass.delete_autofilter();
        }

        private void button1_3_Click(object sender, RibbonControlEventArgs e)
        {
            ZMethodClass.delete_freezepanes();
        }

        private void button1_4_Click(object sender, RibbonControlEventArgs e)
        {
            ZMethodClass.delete_displaygridlines();
        }

        private void button1_5_Click(object sender, RibbonControlEventArgs e)
        {
            ZMethodClass.delete_group();
        }

        private void button2_1_Click(object sender, RibbonControlEventArgs e)
        {
            ZMethodClass.reset_zoom(int.Parse(zoom1_5.Text), cursor1_5.Text);
        }

        private void button1_6_Click(object sender, RibbonControlEventArgs e)
        {
            ZMethodClass.add_lf();
        }

        private void button1_7_Click(object sender, RibbonControlEventArgs e)
        {
            ZMethodClass.renumber(editBox1_7.Text);
        }

        private void editBox1_7_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.上記に続けて = editBox1_7.Text;
            Properties.Settings.Default.Save();
        }

        private void zoom1_5_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.ズーム率 = zoom1_5.Text;
            Properties.Settings.Default.Save();
        }

        private void cursor1_5_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.カーソル位置 = cursor1_5.Text;
            Properties.Settings.Default.Save();
        }

        private void button1_8_Click(object sender, RibbonControlEventArgs e)
        {
            var sel = ZExcelAddIn.Globals.ThisAddIn.Application.Selection;
            editBox1_8.Text = sel[1].Address(true, false).Split('$')[0];
            Properties.Settings.Default.列を記憶 = editBox1_8.Text;
            Properties.Settings.Default.Save();
        }

        private void button1_9_Click(object sender, RibbonControlEventArgs e)
        {
            ZMethodClass.jump_column(editBox1_8.Text);
        }

    }
}
