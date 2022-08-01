using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MateInvet
{
    internal class InputDialog
    {
        public static DialogResult Show(String Titlename ,out string strText)
        {
            string strTemp = string.Empty;

            FrmInputDialog inputDialog = new FrmInputDialog();
            inputDialog.TextHandler = (str) => { strTemp = str; };
            inputDialog.Text = Titlename;

            DialogResult result = inputDialog.ShowDialog();
            strText = strTemp;

            return result;
        }
    }
}
