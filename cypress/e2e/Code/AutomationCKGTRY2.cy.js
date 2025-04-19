const XLSX = require('xlsx');

describe('Loops through each row from Excel', () => {
 let formDataArray;

 before(() => {
  // 🔹 Load Excel file before tests run
  const workbook = XLSX.readFile('cypress/fixtures/Data/2025/Apr/Data.xlsx'); // adjust path as needed
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  formDataArray = XLSX.utils.sheet_to_json(worksheet);
 });

 it('Fills form with each row from Excel', function () {
  cy.wrap(formDataArray).each((formData) => {
   cy.visit('https://pkg.kemkes.go.id/ckg?token=e90277b68bb4b5e7');

   cy.get('input#Nama\\ Lengkap').clear().type(formData['Nama']);
   cy.get('input#NIK').clear().type(formData['NIK']);

   // 🌟 Open date picker
   cy.contains('div', 'Pilih tanggal lahir').click();

   // 🌟 Pick date from Excel (format: YYYY-MM-DD)
   const dob = formData['DateBirthday'];
   cy.get(`.mx-table-date td.cell[title="${dob}"]`).click();

   cy.get('input#No\\.\\ WhatsApp\\ Aktif').clear().type(formData['WA']);

   cy.contains('button', 'Selanjutnya').click();

   // 🌍 Select Provinsi
   cy.contains('div', 'Provinsi').scrollIntoView().click();
   cy.contains('div', formData['Provinsi']).click();

   cy.wait(2000);
  });
 });
});
