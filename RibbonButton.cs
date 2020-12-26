using Microsoft.Office.Tools.Ribbon;

namespace SwissQRCode
{

    public partial class RibbonButton
    {
        private void RibbonButton_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void buttonGenerate_Click(object sender, RibbonControlEventArgs e)
        {
            QRCodeGeneratorWrapper.Generate();
        }
    }
}
