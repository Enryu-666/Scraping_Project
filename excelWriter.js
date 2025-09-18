const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

//function untuk menyimpan data ke file excel
function saveToExcel(data, filename, directory = '') {
  if (!Array.isArray(data) || data.length === 0) {
    console.error('Error: Data tidak ditemukan atau tidak sesuai format array.');
    return;
  }

  
}