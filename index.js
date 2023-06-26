const PDFServicesSdk = require('@adobe/pdfservices-node-sdk');
const fs = require('fs');
const AdmZip = require('adm-zip');
const { log } = require('console');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const filepath = './ExtractedData.csv'
if (!fs.existsSync(filepath)) {
    fs.writeFileSync(filepath, '', (err) => {
        if (err) {
            console.error(err)
        }
        else {
            console.log('lol')
        }
    })
}
else {
    fs.unlinkSync(filepath);
    fs.writeFileSync(filepath, '', (err) => {
        if (err) {
            console.error(err)
        }
    })
}

const csvWriter = createCsvWriter({
    path: filepath,
    header: [
        { id: 'city', title: 'Bussiness__City' },
        { id: 'coun', title: 'Bussiness__Country' },
        { id: 'desc', title: 'Bussiness__Description' },
        { id: 'bname', title: 'Bussiness__Name' },
        { id: 'baddress', title: 'Bussiness__StreetAddress' },
        { id: 'bzipcode', title: 'Bussiness__Zipcode' },
        { id: 'caddress1', title: 'Customer__Address__line1' },
        { id: 'caddress2', title: 'Customer__Address__line2' },
        { id: 'cemail', title: 'Customer__Email' },
        { id: 'cname', title: 'Customer__Name' },
        { id: 'cnumber', title: 'Customer__PhoneNumber' },
        { id: 'iname', title: 'Invoice__BillDetails__Name' },
        { id: 'iquan', title: 'Invoice__BillDetails__Quantity' },
        { id: 'irate', title: 'Invoice__BillDetails__Rate' },
        { id: 'idesc', title: 'Invoice__Description' },
        { id: 'idate', title: 'Invoice__DueDate' },
        { id: 'iidate', title: 'Invoice__IssueDate' },
        { id: 'inum', title: 'Invoice__Number' },
        { id: 'itax', title: 'Invoice__Tax' },
    ]
});

const credentials = PDFServicesSdk.Credentials
    .serviceAccountCredentialsBuilder()
    .fromFile('pdfservices-api-credentials.json')
    .build();

const executionContext = PDFServicesSdk.ExecutionContext.create(credentials);

const options = new PDFServicesSdk.ExtractPDF.options.ExtractPdfOptions.Builder()
    .addElementsToExtract(PDFServicesSdk.ExtractPDF.options.ExtractElementType.TEXT, PDFServicesSdk.ExtractPDF.options.ExtractElementType.TABLES)
    .addTableStructureFormat(PDFServicesSdk.ExtractPDF.options.TableStructureType.CSV)
    .build();

for (let i = 0; i < 100; i++) {
    const extractPDFOperation = PDFServicesSdk.ExtractPDF.Operation.createNew(),
        input = PDFServicesSdk.FileRef.createFromLocalFile(
            './output' + i + '.pdf',
            PDFServicesSdk.ExtractPDF.SupportedSourceFormat.pdf
        );
    extractPDFOperation.setOptions(options);

    extractPDFOperation.setInput(input);
    const OUTPUT_ZIP = './ExtractTextInfoFromPDF' + i + '.zip'
    if (fs.existsSync(OUTPUT_ZIP)) fs.unlinkSync(OUTPUT_ZIP);
    extractPDFOperation.execute(executionContext)
        .then(result => result.saveAsFile(OUTPUT_ZIP))
        .then(() => {
            console.log('Successfully extracted information from PDF.\n');
            let zip = new AdmZip(OUTPUT_ZIP);
            let jsondata = zip.readAsText('structuredData.json');
            let data = JSON.parse(jsondata);
            const records = []
            let city = 'Jamestown'
            let coun = 'Tennessee, USA'
            let desc = 'We are here to serve you better. Reach out to us in case of any concern or feedbacks.'
            let bname = 'NearBy Electronics'
            let baddress = '3741 Glory Road'
            let bzipcode = 38556
            let itax = 10
            let s = ''
            let trnum = 1
            let cemail
            let cname
            let inum
            let caddress1
            let cnumber
            console.log('\n')
            let caddress2
            let index = 0
            let dates = []
            let x
            let y
            const arrayy = data.elements.filter((l) => l.Text !== undefined)
            arrayy.forEach(element => {
                console.log(element.Text)
                if (element.Text.includes('AMOUNT')) {
                    x = index + 1;
                }
                if (element.Text.includes('Subtotal')) {
                    y = index;
                }
                if ((element.Text).match(/\b\d{2}-\d{2}-\d{4}\b/)) {
                    let matched = (element.Text).match(/\b\d{2}-\d{2}-\d{4}\b/)[0]
                    dates.push(matched)
                }
                if (/^\d{3}-\d{3}-\d{4}.*/.test(element.Text)) {
                    cnumber = element.Text
                }
                if ((/^\d{3}\s+\w+/).test(element.Text)) {
                    caddress1 = element.Text
                    caddress2 = arrayy[index + 1].Text
                }
                if (element.Text.includes('@')) {
                    cemail = (element.Text)
                    cname = arrayy[index - 1].Text
                }
                if (element.Text.includes("#")) {
                    inum = element.Text.split(" ")[1]
                }
                index++
            });
            let a, b, c
            while (x < y) {
                a = arrayy[x].Text;
                x++;
                b = arrayy[x].Text
                x++
                c = arrayy[x].Text
                x = x + 2
                let o = ''
                records.push({
                    city: city, coun: coun, desc: desc, bname: bname,
                    baddress: baddress, bzipcode: bzipcode, caddress1: caddress1,
                    caddress2: o, cemail: cemail, cname: cname,
                    cnumber: cnumber, iname: a, iquan: b,
                    irate: c, idesc: o, idate: dates[0],
                    iidate: dates[1], inum: inum, itax: itax
                })
            }

            csvWriter.writeRecords(records)
                .then(() => {
                    console.log('...Done');
                });
        })
        .catch(err => console.log(err));
}
