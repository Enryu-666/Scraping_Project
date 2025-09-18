const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

//function untuk menyimpan data ke file excel
function saveToExcel(data, filename, directory = '') {
    if (!Array.isArray(data) || data.length === 0) {
        console.error('Error: Data tidak ditemukan atau tidak sesuai format array.');
        return;
    }
    //membuat path lengkap untuk file excel
    const fullpath = path.join(directory, filename);
    
    try {
        //membuat direktori jika belum ada
        if (directory) {
            fs.mkdirSync(directory, { recursive: true });
        }
        //menulis data ke file excel
        const workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.json_to_sheet(data);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        xlsx.writeFile(workbook, fullpath);
        console.log(`Data berhasil disimpan ke ${fullpath}`);
    } catch (error) {
        console.error('Error: Gagal menyimpan data ke file Excel.', error);
    }
}