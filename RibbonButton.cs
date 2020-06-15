using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using QRCoder;
using SwissQRCode.Properties;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using static QRCoder.PayloadGenerator;

namespace SwissQRCode
{
    public partial class RibbonButton
    {
        private void RibbonButton_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void buttonGenerate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            
                string contactIBAN = sheet.get_Range("A4").Value2.ToString();
                SwissQrCode.Iban iban = new SwissQrCode.Iban(contactIBAN, SwissQrCode.Iban.IbanType.Iban);

                string contactName = sheet.get_Range("A5").Value2.ToString();
                string contactStreet = sheet.get_Range("A6").Value2.ToString();
                string contactPlace = sheet.get_Range("A7").Value2.ToString();
                SwissQrCode.Contact contact = new SwissQrCode.Contact(contactName, "CH", contactStreet, contactPlace);

                string debitorName = sheet.get_Range("A10").Value2.ToString();
                string debitorStreet = sheet.get_Range("A11").Value2.ToString();
                string debitorPlace = sheet.get_Range("A12").Value2.ToString();
                SwissQrCode.Contact debitor = new SwissQrCode.Contact(debitorName, "CH", debitorStreet, debitorPlace);

                string additionalInfo1 = sheet.get_Range("F9").Value2.ToString();
                string additionalInfo2 = sheet.get_Range("F10").Value2.ToString();
                SwissQrCode.AdditionalInformation additionalInformation = new SwissQrCode.AdditionalInformation(additionalInfo1, additionalInfo2);

                SwissQrCode.Reference reference = new SwissQrCode.Reference(SwissQrCode.Reference.ReferenceType.NON);

                SwissQrCode.Currency currency = SwissQrCode.Currency.CHF;

                decimal amount = (decimal) sheet.get_Range("B16").Value2;

                SwissQrCode generator = new SwissQrCode(iban, currency, contact, reference, additionalInformation, debitor, amount);

                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(generator.ToString(), QRCodeGenerator.ECCLevel.M);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeAsBitmap = qrCode.GetGraphic(20, Color.Black, Color.White, Resources.CH_Kreuz_7mm, 14, 1);

                string picturePath = Path.GetTempPath() + "qrcode.bmp";
                if (File.Exists(picturePath))
                {
                    File.Delete(picturePath);
                }
                qrCodeAsBitmap.Save(picturePath, ImageFormat.Bmp);
                sheet.Shapes.AddPicture(picturePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, 180, 40, 140, 140);
                File.Delete(picturePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Swiss QR Code Generator", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
