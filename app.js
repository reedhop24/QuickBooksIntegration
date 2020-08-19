const express = require('express');
const app = express();
const bodyParser = require('body-parser');

const addDate = (date) => {
    date = date.split('T').shift();
    let arr = date.split('-');
    if(arr[1] === '12') {
        arr[0] = (parseInt(arr[0])+1).toString();
        arr[1] = '01';
    } else {
        if(arr[1] < 9) {
            arr[1] = '0'+(parseInt(arr[1])+1);
        } else {
            arr[1] = (parseInt(arr[1])+1).toString();
        }
    }
    return arr.join('-')
}

const removeTime = (date) => {
    date = date.split('T').shift();
    return date
}

app.use(bodyParser({limit: '50mb'}));

app.post('/analytics', (req, res) => {
    var xl = require('excel4node');
    
    // Create a new instance of a Workbook class
    var wb = new xl.Workbook();
    
    // Add Worksheets to the workbook
    var ws = wb.addWorksheet('Sheet 1');
    
    // Create a reusable style
    var style = wb.createStyle({
        font: {
            color: '#000000',
            size: 12,
        },
    });
    
    // Set value of cell A1 to 100 as a number type styled with paramaters of style
    ws.cell(1, 1)
    .string('id')
    
    // Set value of cell B1 to 200 as a number type styled with paramaters of style
    ws.cell(1, 2)
    .string('Company Name')
    
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(1, 3)
    .string('Create Time')
    
    ws.cell(1, 4)
    .string('Update Time')

    const customers = req.body.QueryResponse.Customer
    
    for(let i = 0; i < customers.length; i++) {
        ws.cell(i+2, 1)
        .string(customers[i].Id)

        ws.cell(i+2, 2)
        .string(customers[i].FullyQualifiedName)

        ws.cell(i+2, 3)
        .string(addDate(customers[i].MetaData.CreateTime))

        ws.cell(i+2, 4)
        .string(removeTime(customers[i].MetaData.LastUpdatedTime))
    }
    
    wb.write('Excel.xlsx');
    res.sendFile(__dirname + '/Excel.xlsx');
});

app.listen(500, () => {
    console.log('server running 500');
})