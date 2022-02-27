
using Microsoft.Office.Tools.Ribbon;

namespace BankingQRCodeForExcel
{

    public partial class RibbonButton
    {
        private void RibbonButton_Load(object sender, RibbonUIEventArgs e)
        {
            // nothing to do
        }

        private void buttonSwiss_Click(object sender, RibbonControlEventArgs e)
        {
            QRCodeGeneratorWrapper.GenerateSwissQRCode();
        }

        private void buttonSEPA_Click(object sender, RibbonControlEventArgs e)
        {
            QRCodeGeneratorWrapper.GenerateSEPAQRCode();
        }
    }
}
