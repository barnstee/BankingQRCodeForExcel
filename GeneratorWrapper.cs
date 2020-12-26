using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using QRCoder;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

namespace SwissQRCode
{
    public static class QRCodeGeneratorWrapper
    {
        public static void Generate()
        {
            try
            {
                Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

                string contactIBAN = sheet.get_Range("A4").Value2.ToString();
                PayloadGenerator.SwissQrCode.Iban iban = new PayloadGenerator.SwissQrCode.Iban(contactIBAN, PayloadGenerator.SwissQrCode.Iban.IbanType.Iban);

                string contactName = sheet.get_Range("A5").Value2.ToString();
                string contactStreet = sheet.get_Range("A6").Value2.ToString();
                string contactPlace = sheet.get_Range("A7").Value2.ToString();
                PayloadGenerator.SwissQrCode.Contact contact = new PayloadGenerator.SwissQrCode.Contact(contactName, "CH", contactStreet, contactPlace);

                string debitorName = sheet.get_Range("A10").Value2.ToString();
                string debitorStreet = sheet.get_Range("A11").Value2.ToString();
                string debitorPlace = sheet.get_Range("A12").Value2.ToString();
                PayloadGenerator.SwissQrCode.Contact debitor = new PayloadGenerator.SwissQrCode.Contact(debitorName, "CH", debitorStreet, debitorPlace);

                string additionalInfo1 = sheet.get_Range("F9").Value2.ToString();
                string additionalInfo2 = sheet.get_Range("F10").Value2.ToString();
                PayloadGenerator.SwissQrCode.AdditionalInformation additionalInformation = new PayloadGenerator.SwissQrCode.AdditionalInformation(additionalInfo1, additionalInfo2);

                PayloadGenerator.SwissQrCode.Reference reference = new PayloadGenerator.SwissQrCode.Reference(PayloadGenerator.SwissQrCode.Reference.ReferenceType.NON);

                PayloadGenerator.SwissQrCode.Currency currency = PayloadGenerator.SwissQrCode.Currency.CHF;

                decimal amount = (decimal)sheet.get_Range("B16").Value2;

                PayloadGenerator.SwissQrCode generator = new PayloadGenerator.SwissQrCode(iban, currency, contact, reference, additionalInformation, debitor, amount);

                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(generator.ToString(), QRCodeGenerator.ECCLevel.M);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeAsBitmap = qrCode.GetGraphic(20, Color.Black, Color.White, Properties.Resources.CH_Kreuz_7mm, 14, 1);

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
