# Banking QR Code Generator For Excel

A QR code generator for Microsoft Excel to be used in the Swiss and SEPA/EU banking sector.

## Background

The requirements for the Swiss QR bill can be read [here](https://www.moneytoday.ch/lexikon/qr-rechnung).

The requirements for the SEPA QR code can be read [here](https://en.wikipedia.org/wiki/EPC_QR_code).

The QR code generator is an Excel ribbon button that, when clicked, will read the data from the supplied BankingQRBill.xlsx Excel template, generate the QR code and then place it in the correct position in the BankingQRBill.xlsx.

## Installation

1. Publish: Load the solution in Visual Studio and publish it. This generates an installer in the project's "publish" directory.
2. Deploy: Deploy the installation files as well as the "BankingQRBill.xlsx" Excel file to the target machine and run setup. The Excel add-on will be installed.
3. Open: Open the "BankingQRBill.xlsx" Excel file on the target machine.

## Usage

1. Update: Update the Swiss "Empfangsschein" section of the template with the information you need (sender, receiver, payment amount).
2. Generate: Click on the "Generate" button in the "Banking QRCode" section from the "Home" Excel ribbon-menu.
3. Print: Print the Excel page and attach it to your paper bill.

