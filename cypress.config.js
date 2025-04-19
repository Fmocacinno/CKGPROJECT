const { defineConfig } = require('cypress');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

module.exports = defineConfig({
  e2e: {
    setupNodeEvents(on, config) {
      on('task', {
        // Read Excel
        readExcel(filePath) {
          const workbook = xlsx.readFile(path.resolve(filePath));
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: '' });
          return jsonData;
        },

        // Write remark to original file
        writeExcelRemark({ filePath, nik, remark }) {
          const workbook = xlsx.readFile(path.resolve(filePath));
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

          const updatedData = jsonData.map((row) => {
            if (row.NIK === nik) {
              return { ...row, remark };
            }
            return row;
          });

          const newWorksheet = xlsx.utils.json_to_sheet(updatedData);
          workbook.Sheets[sheetName] = newWorksheet;
          xlsx.writeFile(workbook, path.resolve(filePath));

          return true;
        },

        // 🆕 Export summary report
        exportReportingData({ filePath, rows }) {
          const reportingData = rows.map((row, index) => ({
            No: index + 1,
            TanggalPemeriksaan: row.TanggalPemeriksaan,
            Nama: row.Nama,
            NIK: row.NIK,
            remark: row.remark || '',
          }));

          const worksheet = xlsx.utils.json_to_sheet(reportingData);
          const workbook = xlsx.utils.book_new();
          xlsx.utils.book_append_sheet(workbook, worksheet, 'Report');

          const outputPath = path.resolve(filePath);
          xlsx.writeFile(workbook, outputPath);

          return true;
        }
      });
    },

    viewportWidth: 1280,
    viewportHeight: 800,
  }
});
