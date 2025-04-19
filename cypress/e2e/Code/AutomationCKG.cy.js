describe('Form Automation from Excel (Multiple Rows)', () => {
 let formDataArray = [];

 before(() => {
  cy.task('readExcel', 'cypress/fixtures/Data/2025/Apr/Data.xlsx').then((data) => {
   formDataArray = data;
  });
 });

 it('Loops through each row from Excel', function () {
  cy.wrap(formDataArray).each((formData, index) => {
   cy.visit('https://pkg.kemkes.go.id/ckg?token=e90277b68bb4b5e7');

   cy.get('input#Nama\\ Lengkap').clear().type(formData['Nama']);
   cy.get('input#NIK').clear().type(formData['NIK']);
   // 🌟 Open date picker
   cy.contains('div', 'Pilih tanggal lahir').click();

   // 🌟 Pick date from Excel
   const dob = formData['DateBirthday']; // format: YYYY-MM-DD
   const dobYear = dob.split('-')[0]; // Extract year from DOB
   const currentYear = new Date().getFullYear(); // Get current year
   if (dobYear !== currentYear.toString()) {
    // Click the button for the current year
    cy.get('button.mx-btn.mx-btn-text.mx-btn-current-year').click();
   }
   cy.get(`.mx-table-date td.cell[title="${dob}"]`).click();
   // input Whats App
   cy.get('input#No\\.\\ WhatsApp\\ Aktif').clear().type(formData['WA']);

   cy.contains('button', 'Selanjutnya').click();
   cy.contains('div', 'Provinsi').scrollIntoView().click();
   cy.contains('div', formData['Provinsi']).click();

   cy.wait(2000); // optional pause for UI
  });
 });
});
