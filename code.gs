const wbook = SpreadsheetApp.getActive();

function doGet() {
  const sheet = wbook.getSheetByName('BCS-UMR STCK');
  const data_stock = [];
  const product_expired = [];
  const product_kurang_dari_tiga_bulan = [];
  const product_lebih_dari_tiga_bulan = [];

  const startRow = 9;  // Mulai dari baris ke-9
  const startColumn = 3;  // Mulai dari kolom C
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const numRows = lastRow - startRow + 1;
  const numColumns = lastColumn - startColumn + 1;

  const range = sheet.getRange(startRow, startColumn, numRows, numColumns);
  const values = range.getValues();

    // Stock
  for (let i = 0; i < numRows; i++) {
    const no = values[i][0];
    const code_item = values[i][1];
    const description = values[i][2];
    const category = values[i][3];
    const plant = values[i][4];
    const tgl_masuk = values[i][5];
    const expired_date = values[i][6];
    const aging_days = values[i][7];
    const shelf_life_days = values[i][8];
    const shelf_life_month = values[i][9];
    const self_life = values[i][10];
    const beginning_colly = values[i][11];
    const beginning_kg = values[i][12];
    const beginning_pallet = values[i][13];
    const mutasi_in_colly = values[i][14];
    const mutasi_in_kg = values[i][15];
    const mutasi_in_pallet = values[i][16];
    const mutasi_out_colly = values[i][17];
    const mutasi_out_kg = values[i][18];
    const mutasi_out_pallet = values[i][19];
    const ending_stock_colly = values[i][20];
    const ending_stock_kg = values[i][21];
    const ending_stock_pallet = values[i][22];
    const hasil_sampling = values[i][23];

    if (!description == "") {
    data_stock.push({
      no, code_item, description, category, plant, tgl_masuk, expired_date, aging_days,
      shelf_life_days, shelf_life_month, self_life, beginning_colly, beginning_kg,
      beginning_pallet, mutasi_in_colly, mutasi_in_kg, mutasi_in_pallet,
      mutasi_out_colly, mutasi_out_kg, mutasi_out_pallet,
      ending_stock_colly, ending_stock_kg, ending_stock_pallet, hasil_sampling
    });
    } 
  }

    // Product Expired
  for (let i = 0; i < numRows; i++) {
    const code_item = values[i][25];
    const description = values[i][26];
    const category = values[i][27];
    const total_per_item = values[i][28];
  
    if (code_item == "") {
      break
    } 
     product_expired.push({code_item, description,category,total_per_item})
  }

    // Produk < 3 Month
  for (let i = 0; i < numRows; i++) {
    const code_item = values[i][30];
    const description = values[i][31];
    const category = values[i][32];
    const total_per_item = values[i][33];
  
    if (code_item == "") {
      break
    } 
     product_kurang_dari_tiga_bulan.push({code_item, description,category,total_per_item})
  }

    // Produk > 3 Month
  for (let i = 0; i < numRows; i++) {
    const code_item = values[i][35];
    const description = values[i][36];
    const category = values[i][37];
    const total_per_item = values[i][38];
  
    if (code_item == "") {
      break
    } 
     product_lebih_dari_tiga_bulan.push({code_item, description,category,total_per_item})
  }
 
  const response = {
    "stock": data_stock,
    "product_expired": product_expired,
    "product_kurang_dari_tiga_bulan" : product_kurang_dari_tiga_bulan,
    "product_lebih_dari_tiga_bulan" : product_lebih_dari_tiga_bulan
  };

  console.log(response);
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}
