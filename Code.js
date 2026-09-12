/**
 * Pixelin - Google Apps Script Backend System
 * Web App & Google Sheets Database Controller
 */

// Konfigurasi Nama Sheet Database
const SHEETS = {
  UNITS: 'Data_Unit',
  TRANSACTIONS: 'Barang_Keluar_Penjualan',
  CASHFLOW: 'Cashflow_Harian',
  CUSTOMERS: 'Galeri_Customers',
  SETTINGS: 'Pengaturan'
};

/**
 * Render halaman web HTML utama
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Pixelin - Jual Beli HP Google Pixel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fungsi Otomatis Setup Database Spreadsheet
 * Jalankan fungsi ini sekali di Google Apps Script editor untuk membuat struktur tabel secara otomatis.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Sheet Data Unit (Inventaris & Stok)
  let unitSheet = ss.getSheetByName(SHEETS.UNITS);
  if (!unitSheet) {
    unitSheet = ss.insertSheet(SHEETS.UNITS);
    unitSheet.appendRow([
      'Kode Unit', 'Nama Unit', 'Nomor Seri (Privat)', 'Harga Beli (Rp)', 
      'Harga Jual (Rp)', 'Stok', 'Deskripsi Unit', 'Foto 1', 'Foto 2', 
      'Foto 3', 'Foto 4', 'Foto 5', 'Status', 'Tanggal Masuk'
    ]);
    unitSheet.getRange('A1:N1').setFontWeight('bold').setBackground('#4a5568').setFontColor('#ffffff');
  }

  // 2. Sheet Transaksi / Barang Keluar
  let txSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
  if (!txSheet) {
    txSheet = ss.insertSheet(SHEETS.TRANSACTIONS);
    txSheet.appendRow([
      'ID Transaksi', 'Tanggal', 'Kode Unit', 'Nama Unit', 
      'Nomor Seri', 'Harga Beli (Rp)', 'Harga Jual (Rp)', 
      'Keuntungan (Rp)', 'Nama Pembeli', 'Catatan'
    ]);
    txSheet.getRange('A1:J1').setFontWeight('bold').setBackground('#2b6cb0').setFontColor('#ffffff');
  }

  // 3. Sheet Cashflow Harian
  let cfSheet = ss.getSheetByName(SHEETS.CASHFLOW);
  if (!cfSheet) {
    cfSheet = ss.insertSheet(SHEETS.CASHFLOW);
    cfSheet.appendRow([
      'ID Cashflow', 'Tanggal', 'Tipe', 'Kategori', 'Keterangan', 'Jumlah (Rp)', 'Ref Kode Unit'
    ]);
    cfSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#2f855a').setFontColor('#ffffff');
  }

  // 4. Sheet Galeri Customers
  let custSheet = ss.getSheetByName(SHEETS.CUSTOMERS);
  if (!custSheet) {
    custSheet = ss.insertSheet(SHEETS.CUSTOMERS);
    custSheet.appendRow([
      'ID Customer', 'Nama Customer', 'Unit Dibeli', 'Foto URL', 'Testimoni', 'Tanggal'
    ]);
    custSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#d69e2e').setFontColor('#ffffff');
  }

  // 5. Sheet Pengaturan System
  let settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEETS.SETTINGS);
    settingsSheet.appendRow(['Key', 'Value']);
    settingsSheet.appendRow(['ADMIN_PASSWORD', 'tikbraownernya']);
    settingsSheet.appendRow(['WA_NUMBER', '085601197498']);
  }

  return { success: true, message: 'Database Spreadsheet Pixelin Berhasil Dibuat!' };
}

/**
 * Verifikasi Kode Akses Admin
 */
function verifyAdminPassword(inputPassword) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
    
    let storedPassword = 'tikbraownernya'; // Default fallback
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === 'ADMIN_PASSWORD') {
          storedPassword = data[i][1];
          break;
        }
      }
    }
    
    if (inputPassword === storedPassword) {
      return { success: true, token: 'AUTH_VALID_PIXELIN_' + new Date().getTime() };
    } else {
      return { success: false, message: 'Kode akses password admin salah!' };
    }
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Mengambil data Katalog untuk Publik (Nomor Seri disembunyikan)
 */
function getPublicCatalogData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const unitSheet = ss.getSheetByName(SHEETS.UNITS);
    if (!unitSheet) return { units: [], customers: [] };

    const data = unitSheet.getDataRange().getValues();
    const headers = data[0];
    const units = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Hanya tampilkan jika stok > 0 atau status Aktif
      if (row[5] > 0 || row[12] === 'Tersedia') {
        units.push({
          kodeUnit: row[0],
          namaUnit: row[1],
          // Nomor Seri (row[2]) DISEMBUNYIKAN UNTUK PUBLIK
          hargaJual: row[4],
          stok: row[5],
          deskripsi: row[6],
          foto1: row[7] || '',
          foto2: row[8] || '',
          foto3: row[9] || '',
          foto4: row[10] || '',
          foto5: row[11] || '',
          status: row[12]
        });
      }
    }

    // Ambil galeri customer
    const custSheet = ss.getSheetByName(SHEETS.CUSTOMERS);
    const customers = [];
    if (custSheet) {
      const custData = custSheet.getDataRange().getValues();
      for (let j = 1; j < custData.length; j++) {
        customers.push({
          id: custData[j][0],
          nama: custData[j][1],
          unit: custData[j][2],
          fotoUrl: custData[j][3],
          testimoni: custData[j][4],
          tanggal: custData[j][5]
        });
      }
    }

    return { success: true, units: units, customers: customers };
  } catch (err) {
    return { success: false, message: err.toString(), units: [], customers: [] };
  }
}

/**
 * Mengambil Seluruh Data untuk Dashboard Admin
 */
function getAdminDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Data Inventaris Lengkap (Termasuk Nomor Seri & Harga Beli)
    const unitSheet = ss.getSheetByName(SHEETS.UNITS);
    const units = [];
    if (unitSheet) {
      const uData = unitSheet.getDataRange().getValues();
      for (let i = 1; i < uData.length; i++) {
        units.push({
          rowIndex: i + 1,
          kodeUnit: uData[i][0],
          namaUnit: uData[i][1],
          nomorSeri: uData[i][2], // Ditampilkan untuk Admin
          hargaBeli: uData[i][3],
          hargaJual: uData[i][4],
          stok: uData[i][5],
          deskripsi: uData[i][6],
          foto1: uData[i][7],
          foto2: uData[i][8],
          foto3: uData[i][9],
          foto4: uData[i][10],
          foto5: uData[i][11],
          status: uData[i][12],
          tanggalMasuk: uData[i][13]
        });
      }
    }

    // 2. Data Barang Keluar / Penjualan
    const txSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
    const transactions = [];
    if (txSheet) {
      const tData = txSheet.getDataRange().getValues();
      for (let i = 1; i < tData.length; i++) {
        transactions.push({
          rowIndex: i + 1,
          id: tData[i][0],
          tanggal: tData[i][1],
          kodeUnit: tData[i][2],
          namaUnit: tData[i][3],
          nomorSeri: tData[i][4],
          hargaBeli: tData[i][5],
          hargaJual: tData[i][6],
          keuntungan: tData[i][7],
          namaPembeli: tData[i][8],
          catatan: tData[i][9]
        });
      }
    }

    // 3. Data Cashflow Harian
    const cfSheet = ss.getSheetByName(SHEETS.CASHFLOW);
    const cashflows = [];
    if (cfSheet) {
      const cData = cfSheet.getDataRange().getValues();
      for (let i = 1; i < cData.length; i++) {
        cashflows.push({
          rowIndex: i + 1,
          id: cData[i][0],
          tanggal: cData[i][1],
          tipe: cData[i][2], // 'Pemasukan' atau 'Pengeluaran'
          kategori: cData[i][3],
          keterangan: cData[i][4],
          jumlah: cData[i][5],
          refKodeUnit: cData[i][6]
        });
      }
    }

    // 4. Data Customers
    const custSheet = ss.getSheetByName(SHEETS.CUSTOMERS);
    const customers = [];
    if (custSheet) {
      const custData = custSheet.getDataRange().getValues();
      for (let i = 1; i < custData.length; i++) {
        customers.push({
          rowIndex: i + 1,
          id: custData[i][0],
          nama: custData[i][1],
          unit: custData[i][2],
          fotoUrl: custData[i][3],
          testimoni: custData[i][4],
          tanggal: custData[i][5]
        });
      }
    }

    return {
      success: true,
      units: units,
      transactions: transactions,
      cashflows: cashflows,
      customers: customers
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Menyimpan / Update Data Unit Baru (Satu Input Terintegrasi)
 * Otomatis mencatat pengeluaran stok & cashflow harian jika unit baru ditambah!
 */
function saveUnitIntegrated(unitData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const unitSheet = ss.getSheetByName(SHEETS.UNITS);
    const cfSheet = ss.getSheetByName(SHEETS.CASHFLOW);
    
    const isEdit = unitData.isEdit;
    const today = new Date().toISOString().split('T')[0];

    if (isEdit && unitData.rowIndex) {
      // Update data baris yang ada
      unitSheet.getRange(unitData.rowIndex, 1, 1, 14).setValues([[
        unitData.kodeUnit,
        unitData.namaUnit,
        unitData.nomorSeri,
        unitData.hargaBeli,
        unitData.hargaJual,
        unitData.stok,
        unitData.deskripsi,
        unitData.foto1 || '',
        unitData.foto2 || '',
        unitData.foto3 || '',
        unitData.foto4 || '',
        unitData.foto5 || '',
        unitData.status || 'Tersedia',
        unitData.tanggalMasuk || today
      ]]);
    } else {
      // Tambah Unit Baru (Otomatis log Barang Masuk & Pengeluaran Modal di Cashflow)
      unitSheet.appendRow([
        unitData.kodeUnit,
        unitData.namaUnit,
        unitData.nomorSeri,
        unitData.hargaBeli,
        unitData.hargaJual,
        unitData.stok,
        unitData.deskripsi,
        unitData.foto1 || '',
        unitData.foto2 || '',
        unitData.foto3 || '',
        unitData.foto4 || '',
        unitData.foto5 || '',
        'Tersedia',
        today
      ]);

      // Integrasi Otomatis ke Cashflow Harian (Pengeluaran Pembelian Stok)
      if (cfSheet && Number(unitData.hargaBeli) > 0) {
        const totalModal = Number(unitData.hargaBeli) * Number(unitData.stok || 1);
        const cfId = 'CF-IN-' + new Date().getTime();
        cfSheet.appendRow([
          cfId,
          today,
          'Pengeluaran',
          'Pembelian Stok Unit',
          `Restok Unit Baru: ${unitData.namaUnit} (${unitData.kodeUnit}) x${unitData.stok}`,
          totalModal,
          unitData.kodeUnit
        ]);
      }
    }

    return { success: true, message: 'Data Unit & Integrasi Keuangan Berhasil Disimpan!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Mencatat Transaksi Penjualan (Barang Keluar)
 * Otomatis mengurangi stok unit, mencatat Pemasukan Kas, dan menghitung Laba/Rugi.
 */
function processSaleTransaction(saleData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const unitSheet = ss.getSheetByName(SHEETS.UNITS);
    const txSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
    const cfSheet = ss.getSheetByName(SHEETS.CASHFLOW);

    const today = new Date().toISOString().split('T')[0];
    const txId = 'TRX-' + new Date().getTime();
    
    // 1. Ambil data unit untuk mendapatkan harga beli & kurangi stok
    let foundUnitRow = -1;
    let currentStock = 0;
    let hargaBeli = saleData.hargaBeli || 0;
    let namaUnit = saleData.namaUnit;
    let nomorSeri = saleData.nomorSeri || '-';

    if (unitSheet) {
      const data = unitSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === saleData.kodeUnit) {
          foundUnitRow = i + 1;
          namaUnit = data[i][1];
          nomorSeri = saleData.nomorSeri || data[i][2];
          hargaBeli = data[i][3];
          currentStock = Number(data[i][5]);
          break;
        }
      }
    }

    const hargaJual = Number(saleData.hargaJual);
    const keuntungan = hargaJual - Number(hargaBeli);

    // 2. Tambah Catatan Penjualan / Barang Keluar
    if (txSheet) {
      txSheet.appendRow([
        txId,
        today,
        saleData.kodeUnit,
        namaUnit,
        nomorSeri,
        hargaBeli,
        hargaJual,
        keuntungan,
        saleData.namaPembeli || 'Pelanggan Direct',
        saleData.catatan || 'Penjualan Lunas'
      ]);
    }

    // 3. Integrasi Otomatis ke Cashflow (Pemasukan Kas Penjualan)
    if (cfSheet) {
      cfSheet.appendRow([
        'CF-OUT-' + new Date().getTime(),
        today,
        'Pemasukan',
        'Penjualan HP',
        `Penjualan: ${namaUnit} (${saleData.kodeUnit}) - Pembeli: ${saleData.namaPembeli || 'Pelanggan'}`,
        hargaJual,
        saleData.kodeUnit
      ]);
    }

    // 4. Update Stok Unit di Sheet Inventaris
    if (foundUnitRow > 0 && unitSheet) {
      const newStock = Math.max(0, currentStock - 1);
      unitSheet.getRange(foundUnitRow, 6).setValue(newStock);
      if (newStock === 0) {
        unitSheet.getRange(foundUnitRow, 13).setValue('Habis');
      }
    }

    return { success: true, message: 'Transaksi Penjualan Berhasil Terintegrasi!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Hapus / Koreksi Data Typo dari Baris Sheet Manapun
 */
function deleteRecordByRowIndex(sheetName, rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) return { success: false, message: 'Sheet tidak ditemukan!' };

    targetSheet.deleteRow(rowIndex);
    return { success: true, message: 'Data berhasil dihapus/dikoreksi!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Tambah/Simpan Galeri Customer
 */
function saveCustomerGallery(custData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const custSheet = ss.getSheetByName(SHEETS.CUSTOMERS);
    const today = new Date().toISOString().split('T')[0];

    if (custSheet) {
      custSheet.appendRow([
        'CUST-' + new Date().getTime(),
        custData.nama,
        custData.unit,
        custData.fotoUrl,
        custData.testimoni || 'Customer puas dengan unit Pixelin!',
        today
      ]);
    }

    return { success: true, message: 'Galeri customer berhasil ditambahkan!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}