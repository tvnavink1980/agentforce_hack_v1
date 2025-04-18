const express = require('express');
const cors = require('cors');
const morgan = require('morgan');
const fs = require('fs');
const bodyParser = require('body-parser');
require('dotenv').config();
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(morgan('dev'));
app.use(bodyParser.json());

app.get('/order', (req, res) => {
  const workbook = xlsx.readFile(path.join(__dirname, 'Dummy_Data.xlsx'));
  // List all sheet names
  const sheetNames = workbook.SheetNames;
  // Access a specific sheet
  if (sheetNames[0] == 'Order') {
    const sheet = workbook.Sheets[sheetNames[0]];
    console.log('Sheetnames ', sheetNames[0])
    const jsonData = xlsx.utils.sheet_to_json(sheet);
    fs.writeFileSync('Order.json', JSON.stringify(jsonData, null, 2));
    return res.send(jsonData);
  }else{
    return res.send('Invalid data')
  }
});

app.get('/vendor', (req, res) => {
  const workbook = xlsx.readFile(path.join(__dirname, 'Dummy_Data.xlsx'));
  // List all sheet names
  const sheetNames = workbook.SheetNames;
  // Access a specific sheet
  if (sheetNames[1] == 'Vendor') {
    const sheet = workbook.Sheets[sheetNames[1]];
    console.log('Sheetnames ', sheetNames[1])
    const jsonData = xlsx.utils.sheet_to_json(sheet);
    fs.writeFileSync('Vendor.json', JSON.stringify(jsonData, null, 2));
    return res.send(jsonData);
  }else{
    return res.send('Invalid data')
  }
});

app.get('/vendor-fullfilment', (req, res) => {
  const workbook = xlsx.readFile(path.join(__dirname, 'Dummy_Data.xlsx'));
  // List all sheet names
  const sheetNames = workbook.SheetNames;
  // Access a specific sheet
  if (sheetNames[2] == 'Vendor Fullfilment') {
    const sheet = workbook.Sheets[sheetNames[2]];
    console.log('Sheetnames ', sheetNames[2])
    const jsonData = xlsx.utils.sheet_to_json(sheet);
    fs.writeFileSync('Vendor-Fullfilment.json', JSON.stringify(jsonData, null, 2));
    return res.send(jsonData);
  }else{
    return res.send('Invalid data')
  }
});

app.get('/deliever-status', async (req, res) => {
  const orderData = await readJsonFile('Order.json');
  if(orderData){
    const allStatus = await orderData.map(order => ({
      "id": order["Order ID"],
      "status": order["Order Status"]
    
    }));
    // return res.json(allStatus)
    const vendorData = await readJsonFile('Vendor-Fullfilment.json');
    // Update the vendor data based on the order status
    vendorData.forEach(vendor => {
      const orderId = vendor["Order ID"];
      let obj = allStatus.find((o) => o.id === orderId);
      // return res.send(obj)
      if(obj){
        if (obj.status === "Closed") {
          vendor["Order_Delivered"] = "Yes";
        }else{
          vendor["Order_Delivered"] = "No";
        }
      }
    });
    fs.writeFileSync('Updated-Vendor-Fullfilment.json', JSON.stringify(vendorData, null, 2));
    return res.send(vendorData)
    // return res.send(allStatus);
  }else{
    return res.send('Invalid data')
  }
});

app.get('/update-vendor-fullfilment', async (req, res) => {
  const filePath = path.join(__dirname, 'Dummy_Data.xlsx');
  const workbook = xlsx.readFile(filePath);
  // List all sheet names
  const sheetNames = workbook.SheetNames;
  // Access a specific sheet
  if (sheetNames[2] == 'Vendor Fullfilment') {
    const updatedData = await readJsonFile('Updated-Vendor-Fullfilment.json');
    const updatedSheet = xlsx.utils.json_to_sheet(updatedData);
    workbook.Sheets[sheetNames[2]] = updatedSheet;
    xlsx.writeFile(workbook, filePath);
    console.log('Excel file updated successfully.');
    return res.send('Excel file updated successfully.');
  }else{
    return res.send('Invalid data')
  }
});

app.get('/create-product', async (req, res) => {
  const filePath = path.join(__dirname, 'Dummy_Data.xlsx');
  const workbook = xlsx.readFile(filePath);
  const productData = await readJsonFile('product.json');
  const productSheet = xlsx.utils.json_to_sheet(productData);
  xlsx.utils.book_append_sheet(workbook, productSheet, 'Products');
  xlsx.writeFile(workbook, 'Dummy_Data.xlsx');
  console.log('New sheet "Products" added to Dummy_Data.xlsx');
  return res.send('New sheet "Products" added to Dummy_Data.xlsx.');
});

app.get('/order-product', async (req, res) => {
  const orders = await readJsonFile('Order.json');
  const products = await readJsonFile('product.json');
  const orderIds = await orders.map((order, index) => {
    return order["Order ID"];
  })
  const productIds = await products.map((product, index) => {
    return product.id
  })
  const orderProduct = await orderProducts(orderIds, productIds);
  fs.writeFileSync('Order-Product.json', JSON.stringify(orderProduct, null, 2));
  return res.send(orderProduct);
})

app.get('/create-order-product', async (req, res) => {
  const filePath = path.join(__dirname, 'Dummy_Data.xlsx');
  const workbook = xlsx.readFile(filePath);
  const orderProductData = await readJsonFile('Order-Product.json');
  const orderProductSheet = xlsx.utils.json_to_sheet(orderProductData);
  xlsx.utils.book_append_sheet(workbook, orderProductSheet, 'Orders-Products');
  xlsx.writeFile(workbook, 'Dummy_Data.xlsx');
  console.log('New sheet "Orders-Products" added to Dummy_Data.xlsx');
  return res.send('New sheet "Orders-Products" added to Dummy_Data.xlsx.');
});

app.get('/location', async (req, res) => {
  res.send('location')
})


async function orderProducts(orders, productIds){
  let  arr = [];
  let id = 0;
  for (let index = 0; index <= orders.length; index++) {
    let orderId = orders[index] 
    let randomNum = Math.floor(Math.random() * 4);
    if (randomNum > 0) {
      for (let i = 0; i < randomNum; i++) {
        const randomProduct = Math.floor(Math.random() * 50001);
        let obj = {
          id: ++id,
          "Order ID": orderId,
          "Product ID": productIds[randomProduct]
        };
        arr.push(obj);
      }
    }
  }
  return arr;
}

async function readJsonFile(filePath) {
  const data = await fs.readFileSync(filePath, 'utf-8');
  return JSON.parse(data);
}


app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
