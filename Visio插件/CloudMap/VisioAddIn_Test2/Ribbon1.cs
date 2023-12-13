using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using VisioAddIn_Test2;


namespace VisioAddIn_Test2
{
    public partial class Ribbon1 : RibbonBase
    {

        public event EventHandler MyCustomEvent;
        public event EventHandler MyCustomEvent1;
        // 触发 MyCustomEvent 事件
        private void triggerMyCustomEvent()
        {
            MyCustomEvent?.Invoke(this, EventArgs.Empty);
        }
        private void triggerMyCustomEvent1()
        {
            MyCustomEvent1?.Invoke(this, EventArgs.Empty);
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
   

        }



        private bool toggleButton_Checked = false;
        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            //RibbonToggleButton toggleButton = (RibbonToggleButton)sender;
            toggleButton_Checked = !toggleButton_Checked;
            if (toggleButton_Checked)
            {

                toggleButton1.Image =Properties.Resources.ON2;
                triggerMyCustomEvent();

            }
            else
            {
               
                toggleButton1.Image = Properties.Resources.OFF2;
                triggerMyCustomEvent1();
            }
        }
       

    }
}
