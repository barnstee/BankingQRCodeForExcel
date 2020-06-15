
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
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
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.WorkbookBeforeSave += new AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
        }

        void Application_WorkbookBeforeSave(Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Worksheet sheet = (Worksheet) Application.ActiveSheet;

            try
            {
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

                string picturePath = Application.StartupPath + "\\qrcode.bmp";
                if (File.Exists(picturePath))
                {
                    File.Delete(picturePath);
                }
                qrCodeAsBitmap.Save(picturePath, ImageFormat.Bmp);
                sheet.Shapes.AddPicture(picturePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, 10, 10, 130, 130);
                File.Delete(picturePath);
            }
            catch (Exception ex)
            {
                sheet.get_Range("A24").Value2 = ex.Message;
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }
        
        #endregion
    }
}
