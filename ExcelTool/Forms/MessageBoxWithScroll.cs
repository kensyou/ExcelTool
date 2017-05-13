using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTool.Forms
{
    public partial class MessageBoxWithScroll : Form
    {
        private bool _IsError;
        public MessageBoxWithScroll(string message, string title, string instruction, bool isError)
        {
            InitializeComponent();
            _IsError = isError;
            SetDefault(message, title, instruction);
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            if (_IsError)
            {
                e.Graphics.DrawIcon(SystemIcons.Error, 10, 10);
            }
            base.OnPaint(e);
        }
        private void SetDefault(string message, string title, string instruction)
        {
            this.Text = title;
            this.lblInstruction.Text = instruction;
            this.txtMessage.Text = message;
        }

    }
}
