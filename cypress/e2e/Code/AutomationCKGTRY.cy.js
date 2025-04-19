describe('Form Automation from Excel (Multiple Rows)', () => {
 let formDataArray = [];

 // Utility function to select year, month, and date dynamically
 function selectYearAndMonthFromPicker(dob) {
  const [year, month, day] = dob.split('-');
  const monthIndex = parseInt(month, 10) - 1;

  // Click year picker
  cy.get('button.mx-btn.mx-btn-text.mx-btn-current-year').click();

  function pickYear(targetYear) {
   cy.get('div.mx-calendar-content table.mx-table-year td.cell').then(($cells) => {
    const yearFound = [...$cells].some(cell => cell.dataset.year === targetYear);

    if (yearFound) {
     cy.get(`td.cell[data-year="${targetYear}"]`).click();
    } else {
     cy.get('button.mx-btn-icon-double-left').click();
     pickYear(targetYear);
    }
   });
  }

  pickYear(year);

  // After year is selected, select the month
  cy.get(`table.mx-table-month td.cell[data-month="${monthIndex}"]`).click();

  // Then select the full date
  cy.get(`.mx-table-date td.cell[title="${dob}"]`).click();
 }

 before(function () {
  cy.task('readExcel', 'cypress/fixtures/Data/2025/Apr/Data.xlsx').then((data) => {
   this.formDataArray = data;
  });
 });

 it('Loops through each row from Excel', function () {
  const submittedNIKs = new Set();

  cy.wrap(this.formDataArray).each((formData) => {
   const nik = formData['NIK'];

   if (submittedNIKs.has(nik)) {
    cy.log(`NIK ${nik} has already been submitted. Skipping this entry.`);
    return;
   }

   cy.visit('https://pkg.kemkes.go.id/ckg?token=e90277b68bb4b5e7');

   cy.get('input#Nama\\ Lengkap').should('be.visible').clear().type(formData['Nama']);
   cy.get('input#NIK').clear().type(nik);

   // 🌟 Open date picker
   cy.contains('div', 'Pilih tanggal lahir').click();

   const dob = formData['DateBirthday']; // Expected format: YYYY-MM-DD

   if (!dob || !/^\d{4}-\d{2}-\d{2}$/.test(dob)) {
    cy.log(`Invalid DOB format for NIK ${nik}: ${dob}`);
    return;
   }

   // 📅 Select year, month, and date
   selectYearAndMonthFromPicker(dob);

   const gender = formData['JenisKelamin']; // Expected "Laki-laki" or "Perempuan"
   if (gender === 'Laki-laki' || gender === 'Perempuan') {
    cy.contains('label', gender).find('input[type="radio"]').check({ force: true });
   } else {
    cy.log(`Invalid gender value for NIK ${nik}: ${gender}`);
    return;
   }
   // WhatsApp input
   cy.get('input#No\\.\\ WhatsApp\\ Aktif').clear().type(formData['WA']);

   // Click Selanjutnya
   cy.contains('button', 'Selanjutnya').click();

   // Select Provinsi
   cy.contains('div', 'Provinsi').scrollIntoView().click();
   cy.contains('div', formData['Provinsi']).click();

   submittedNIKs.add(nik);

   cy.log(`✅ Submitted form for NIK ${nik}`);
   cy.wait(2000);
  });
 });
});
