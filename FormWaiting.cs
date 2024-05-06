using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelQuery
{
    public partial class FormWaiting : Form
    {
        public FormWaiting()
        {
            InitializeComponent();
            this.lblWating.Location = new System.Drawing.Point((this.ClientSize.Width - lblWating.Size.Width) / 2,
                                           (this.ClientSize.Height - lblWating.Size.Height) / 2);
        }
    }
}
