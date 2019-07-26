using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TUFO公式复制工具
{
    public partial class main : Form
    {
        private int childFormNumber = 0;

        public main(Form childf)
        {
            InitializeComponent();

            main.ShowNewForm(this,childf);
        }

        public static void ShowNewForm(Form main,Form childForm)
        {
            for (int i = 0; i < main.MdiChildren.Length; i++)
            {
                if (main.MdiChildren[i].Name == childForm.Name)
                {
                    main.MdiChildren[i].Activate();
                }
            }
            if (main.MdiChildren.Length <= 0 || main.ActiveMdiChild.Name != childForm.Name)
            {
                childForm.MdiParent = main;
                childForm.WindowState = FormWindowState.Normal;
                childForm.MaximizeBox = false;
                childForm.FormBorderStyle = FormBorderStyle.None;
                childForm.Dock = DockStyle.Fill;
                childForm.Show();
            }
        }
    }
}
