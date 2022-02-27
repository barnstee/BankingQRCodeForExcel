using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using QRCoder;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using static QRCoder.PayloadGenerator;

namespace BankingQRCodeForExcel
{
    public static class QRCodeGeneratorWrapper
    {
        public static void GenerateSwissQRCode()
        {
            try
            {
                Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

                string contactIBAN = sheet.get_Range("A4").Value2.ToString();
                SwissQrCode.Iban iban = new SwissQrCode.Iban(contactIBAN, SwissQrCode.Iban.IbanType.Iban);

                string contactName = sheet.get_Range("A5").Value2.ToString();
                string contactStreet = sheet.get_Range("A6").Value2.ToString();
                string contactPlace = sheet.get_Range("A7").Value2.ToString();
                SwissQrCode.Contact contact = SwissQrCode.Contact.WithCombinedAddress(contactName, "CH", contactStreet, contactPlace);

                string debitorName = sheet.get_Range("A10").Value2.ToString();
                string debitorStreet = sheet.get_Range("A11").Value2.ToString();
                string debitorPlace = sheet.get_Range("A12").Value2.ToString();
                SwissQrCode.Contact debitor = SwissQrCode.Contact.WithCombinedAddress(debitorName, "CH", debitorStreet, debitorPlace);

                string additionalInfo1 = sheet.get_Range("F9").Value2.ToString();
                string additionalInfo2 = sheet.get_Range("F10").Value2.ToString();
                SwissQrCode.AdditionalInformation additionalInformation = new SwissQrCode.AdditionalInformation(additionalInfo1, additionalInfo2);

                SwissQrCode.Reference reference = new SwissQrCode.Reference(SwissQrCode.Reference.ReferenceType.NON);

                SwissQrCode.Currency currency = SwissQrCode.Currency.CHF;

                decimal amount = (decimal)sheet.get_Range("B16").Value2;

                SwissQrCode payload = new SwissQrCode(iban, currency, contact, reference, additionalInformation, debitor, amount);

                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(payload.ToString(), QRCodeGenerator.ECCLevel.M);
                Bitmap qrCodeAsBitmap = new QRCode(qrCodeData).GetGraphic(20, Color.Black, Color.White, Properties.Resources.CH_Kreuz_7mm, 14, 1);

                string picturePath = Path.GetTempPath() + "swissqrcode.bmp";
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
                MessageBox.Show(ex.Message, "Banking QRCode for Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void GenerateSEPAQRCode()
        {
            try
            {
                Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

                string iban = sheet.get_Range("A4").Value2.ToString();
                string bic = sheet.get_Range("A5").Value2.ToString();
                string name = sheet.get_Range("A6").Value2.ToString();
                string remittanceInformation = sheet.get_Range("A7").Value2.ToString();

                decimal amount = (decimal)sheet.get_Range("A10").Value2;

                Girocode payload = new Girocode(iban, bic, name, amount, remittanceInformation);

                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(payload.ToString(), QRCodeGenerator.ECCLevel.M);
                Bitmap qrCodeAsBitmap = new QRCode(qrCodeData).GetGraphic(20, Color.Black, Color.White, true);

                string picturePath = Path.GetTempPath() + "sepaqrcode.bmp";
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
                MessageBox.Show(ex.Message, "Banking QRCode for Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

}
