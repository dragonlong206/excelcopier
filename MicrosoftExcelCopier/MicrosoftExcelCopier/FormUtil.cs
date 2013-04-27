using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MicrosoftExcelCopier
{
    public class FormUtil
    {
        public static DialogResult ShowMessageBoxLocalize(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            // Localize messagebox button text
            MessageBoxManager.OK = vi_VN.buttonOKText;
            MessageBoxManager.Yes = vi_VN.buttonYesText;
            MessageBoxManager.No = vi_VN.buttonNoText;
            MessageBoxManager.Cancel = vi_VN.buttonCancelText;
            MessageBoxManager.Register();

            DialogResult result = MessageBox.Show(owner, text, caption, buttons, icon);

            MessageBoxManager.Unregister();

            return result;
        }
    }
}
