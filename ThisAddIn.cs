
using System.Runtime.InteropServices;

namespace BankingQRCodeForExcel
{
    public partial class ThisAddIn
    {
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            // nothing to do
        }

        #endregion

        private AddInUtilities utilities;

        protected override object RequestComAddInAutomationService()
        {
            if (utilities == null)
            {
                utilities = new AddInUtilities();
            }

            return utilities;
        }
    }

    [ComVisible(true)]
    public interface IAddInUtilities
    {
        void Generate();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddInUtilities : IAddInUtilities
    {
        public void Generate()
        {
            QRCodeGeneratorWrapper.GenerateSwissQRCode();
        }
    }
}
