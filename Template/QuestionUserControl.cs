using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Template
{
    public partial class QuestionUserControl : UserControl
    {
        public event EventHandler OnUserControlButtonClicked;

        public string Question { get => txtQuestion.Text.Trim(); set => txtQuestion.Text = value; }

        public QuestionUserControl()
        {
            InitializeComponent();
            btnAdd.Click += (s, e) => this.OnUserControlButtonClicked?.Invoke(this, e);
        }
    }
}
