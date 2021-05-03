# SwissQRCodeExcel
QR code generator for Microsoft Excel to be used in the Swiss banking sector. The requirements for the new Swiss QR bill can be read [here](https://www.moneytoday.ch/lexikon/qr-rechnung/).

The QR code generator is an Excel ribbon button that, when clicked, will read the data from the supplied SwissQRBill.xlsx Excel template, generate the QR code and then place it in the correct position in the SwissQRBill.xlsx.

Installation:
1. Publish: Load the solution in Visual Studio and publish it. This generates an installer in the project's "publish" directory.
2. Deploy: Deploy the installation files as well as the "SwissQRBill.xlsx" Excel file to the target machine and run setup. The Excel add-on will be installed.
3. Open: Open the "SwissQRBill.xlsx" Excel file on the target machine.

Usage:
1. Update: Update the "Empfangsschein" section of the template with the information you need (sender, receiver, payment amount).
2. Generate: Click on the "Generate" button in the "Swiss QR Code" section from the "Home" Excel ribbon-menu.
3. Print: Print the Excel page and attach it to your paper bill.

