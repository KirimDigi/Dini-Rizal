/**
 * Google Apps Script untuk RSVP Wedding Dini & Rizal
 * Spreadsheet ID: 1zseacxItA7wyUVfgwTgCFb6qhN8Gs3wUCqXK2L80wJU
 * Sheet Name: dini&rizal
 * Spreadsheet Name: nama_tamu
 */

// ID Spreadsheet
const SPREADSHEET_ID = '1zseacxItA7wyUVfgwTgCFb6qhN8Gs3wUCqXK2L80wJU';
const SHEET_NAME = 'dini&rizal';

/**
 * Fungsi untuk menangani POST request dari form RSVP
 */
function doPost(e) {
  try {
    // Set CORS headers
    const output = ContentService.createTextOutput();
    output.setMimeType(ContentService.MimeType.JSON);
    
    // Parse data dari request
    let data;
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (jsonError) {
        // If JSON parsing fails, try form data
        data = e.parameter;
      }
    } else {
      data = e.parameter;
    }
    
    // Validasi data
    if (!data.name || !data.message || !data.attendance) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          message: 'Data tidak lengkap'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Simpan data ke spreadsheet
    const result = saveToSpreadsheet(data);
    
    if (result.success) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          message: 'Data berhasil disimpan',
          rowNumber: result.rowNumber
        }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          message: result.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'Terjadi kesalahan server: ' + error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Fungsi untuk menangani GET request (untuk testing dan mengambil data)
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === 'getAllRSVPData') {
      return ContentService
        .createTextOutput(JSON.stringify(getAllRSVPData()))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'getRSVPStats') {
      return ContentService
        .createTextOutput(JSON.stringify(getRSVPStats()))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Default response
    return ContentService
      .createTextOutput(JSON.stringify({
        message: 'RSVP API untuk Wedding Dini & Rizal',
        status: 'active',
        timestamp: new Date().toISOString(),
        availableActions: ['getAllRSVPData', 'getRSVPStats']
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'Error: ' + error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Fungsi untuk menyimpan data ke spreadsheet
 */
function saveToSpreadsheet(data) {
  try {
    // Buka spreadsheet
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Cari atau buat sheet
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      
      // Tambahkan header jika sheet baru
      const headers = [
        'Timestamp',
        'Nama Lengkap',
        'Ucapan',
        'Konfirmasi Kehadiran',
        'Status'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#053a53');
      headerRange.setFontColor('white');
    }
    
    // Siapkan data untuk disimpan
    const timestamp = new Date();
    const rowData = [
      timestamp,
      data.name,
      data.message,
      data.attendance,
      'Tampilkan' // Status default untuk ucapan
    ];
    
    // Tambahkan data ke baris baru
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Format baris data
    const dataRange = sheet.getRange(newRow, 1, 1, rowData.length);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // Auto resize kolom
    sheet.autoResizeColumns(1, rowData.length);
    
    Logger.log('Data berhasil disimpan di baris: ' + newRow);
    
    return {
      success: true,
      rowNumber: newRow,
      timestamp: timestamp
    };
    
  } catch (error) {
    Logger.log('Error in saveToSpreadsheet: ' + error.toString());
    return {
      success: false,
      message: 'Gagal menyimpan ke spreadsheet: ' + error.toString()
    };
  }
}

/**
 * Fungsi untuk testing - bisa dipanggil manual
 */
function testSaveData() {
  const testData = {
    name: 'Test Dini & Rizal',
    message: 'Selamat menempuh hidup baru Dini & Rizal! Semoga samawa selalu.',
    attendance: 'Hadir',
    timestamp: new Date().toISOString()
  };
  
  const result = saveToSpreadsheet(testData);
  Logger.log('Test result: ' + JSON.stringify(result));
  
  return result;
}

/**
 * Fungsi untuk mendapatkan semua data RSVP yang memiliki status Tampilkan
 */
function getAllRSVPData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return {
        success: false,
        message: 'Sheet tidak ditemukan'
      };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return {
        success: true,
        data: [],
        message: 'Belum ada data RSVP'
      };
    }
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    const values = dataRange.getValues();
    
    // Mapping data dan memfilter yang statusnya BUKAN 'Sembunyikan'
    const filteredData = [];
    values.forEach((row, index) => {
      const status = row[4] ? row[4].toString().trim().toLowerCase() : 'tampilkan';
      
      // Jika status diset 'sembunyikan', lewati (hide ucapan)
      if (status === 'sembunyikan' || status === 'hide' || status === 'no') {
        return;
      }
      
      filteredData.push({
        id: index + 2,
        timestamp: row[0],
        name: row[1] || 'Anonim',
        message: row[2] || 'Tidak ada pesan',
        attendance: row[3] || 'TIDAK HADIR'
      });
    });
    
    // Urutkan berdasarkan waktu kirim terbaru di paling atas (Newest First)
    filteredData.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    return {
      success: true,
      data: filteredData,
      total: filteredData.length
    };
    
  } catch (error) {
    Logger.log('Error in getAllRSVPData: ' + error.toString());
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

/**
 * Fungsi untuk menghitung statistik RSVP
 */
function getRSVPStats() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return {
        success: false,
        message: 'Sheet tidak ditemukan'
      };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return {
        success: true,
        stats: {
          totalResponses: 0,
          attending: 0,
          notAttending: 0,
          attendanceRate: 0
        }
      };
    }
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    const values = dataRange.getValues();
    
    let totalResponses = values.length;
    let attendingCount = 0;
    let notAttendingCount = 0;
    
    values.forEach(row => {
      const attendance = row[3] ? row[3].toString().trim().toLowerCase() : '';
      if (attendance.includes('hadir') && !attendance.includes('tidak')) {
        attendingCount++;
      } else {
        notAttendingCount++;
      }
    });
    
    return {
      success: true,
      stats: {
        totalResponses: totalResponses,
        attending: attendingCount,
        notAttending: notAttendingCount,
        attendanceRate: totalResponses > 0 ? Math.round((attendingCount / totalResponses) * 100) : 0
      }
    };
    
  } catch (error) {
    Logger.log('Error in getRSVPStats: ' + error.toString());
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}