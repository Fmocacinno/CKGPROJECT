describe('Form Automation from Excel (Multiple Rows)', () => {
  let formDataArray = [];
  let reportSummary = []; // ✅ This is the fix
  // Utility function to select year, month, and date dynamically
  function selectYearAndMonthFromPicker(dob) {
    const [year, month, day] = dob.split('-');
    const monthIndex = parseInt(month, 10) - 1;
    let processedRows = [];
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
    cy.task('readExcel', 'cypress/e2e/Code/Data/2025/Apr/DataCating.xlsx').then((data) => {
      this.formDataArray = data;
    });
  });
  beforeEach(() => {
    cy.visit('https://sehatindonesiaku.kemkes.go.id/', { failOnStatusCode: false });

    cy.url().then((url) => {
      if (url.includes('/profile')) {
        cy.log('✅ Already logged in. Session restored.');
      } else {
        cy.session('userSession', () => {
          cy.clearCookies();
          cy.clearLocalStorage();
          cy.visit('https://sehatindonesiaku.kemkes.go.id/login');

          if (!url.includes('/login')) {
            cy.visit('https://sehatindonesiaku.kemkes.go.id/login');
          }

          cy.wait(2000);

          cy.get('input[name="Email"]').type('puskesmas.mondokan@gmail.com');
          cy.wait(2000);
          cy.get('input[name="Kata sandi"]').type('Pkm12345@');

          cy.log('⏳ Please complete CAPTCHA manually...');
          cy.wait(10000);

          cy.get('button[type="submit"]').click();
          cy.wait(10000);
          cy.url().should('include', '/profile');

          // Scroll to the specific div containing the text
          cy.contains('div', 'Pengguna dapat menyampaikan pertanyaan, keluhan dan/atau pengaduan sehubungan dengan penggunaan Data Pribadi atau Informasi Kesehatan Pasien pada ASIK.')
            .scrollIntoView();
          cy.wait(5000);
        });
      }
    });
  });


  it('Loops through each row from Excel', function () {
    const submittedNIKs = new Set();

    cy.wrap(this.formDataArray).each((formData) => {
      const nik = formData['NIK'];
      const dob = formData['DateBirthday']; // Expected format: YYYY-MM-DD
      const gender = formData['JenisKelamin']; // "Laki-laki" or "Perempuan"
      const dateStr = formData['TanggalPemeriksaan']; // 👈 Make sure this is defined here

      if (!dateStr || isNaN(new Date(dateStr))) {
        cy.log(`Invalid or missing TanggalPemeriksaan for NIK ${nik}`);
        return;
      }

      const tanggalPemeriksaan = new Date(dateStr);
      const dayOfMonth = tanggalPemeriksaan.getDate().toString(); // "14"





      if (submittedNIKs.has(nik)) {
        cy.log(`NIK ${nik} has already been submitted. Skipping this entry.`);
        return;
      }
      cy.wait(2000);
      cy.visit('https://sehatindonesiaku.kemkes.go.id/ckg-pendaftaran-individu');
      cy.wait(1000);

      // cy.contains('button', 'Daftar Baru').click();
      cy.contains('button', 'Daftar Baru').click();

      cy.wait(2000);
      cy.get('input#nik') // or: cy.get('input[name="NIK"]')
        .should('be.visible')
        .clear()
        .type(formData['NIK']); // Replace `formData['NIK']` with your NIK value

      cy.wait(2000);
      cy.get('input#Nama\\ Lengkap').should('be.visible').clear().type(formData['Nama']);
      // 🌟 Open date picker
      cy.contains('div', 'Pilih tanggal lahir').click();


      if (!dob || !/^\d{4}-\d{2}-\d{2}$/.test(dob)) {
        cy.log(`Invalid DOB format for NIK ${nik}: ${dob}`);
        return;
      }

      // 📅 Select year, month, and date
      selectYearAndMonthFromPicker(dob);

      if (gender === 'Laki-laki' || gender === 'Perempuan') {
        // Open dropdown
        cy.contains('div', 'Pilih jenis kelamin')
          .should('be.visible')
          .click();

        // Click the matching dropdown option
        cy.get('.py-2.px-4.cursor-pointer')
          .contains(gender)
          .click();
      } else {
        cy.log(`Invalid gender value for NIK ${nik}: ${gender}`);
        return;
      }
      // WhatsApp input
      cy.get('input#No\\ Whatsapp') // Escape the space in the ID
        .should('be.visible')
        .clear()
        .type(formData['WA']); // Assumes the WhatsApp number is in the Excel column "WA"

      cy.wait(2000);
      cy.get('button').each(($btn) => {
        const text = $btn.text().trim();
        if (text.startsWith(dayOfMonth)) {
          cy.wrap($btn).click();
          return false; // Break out of the loop
        }
      });
      // Click lanjut if date same
      cy.contains('button', 'Selanjutnya').click();
      // Check if the warning message is visible
      cy.wait(2000);
      // Check if the warning message is visible
      cy.get('.p-2').then(($warning) => {
        if ($warning.is(':visible')) {
          cy.get('button', { timeout: 0 }).then(($buttons) => {
            const lanjutBtn = $buttons.filter((i, el) => el.innerText.includes('Lanjut'))[0];
            if (lanjutBtn) {
              cy.wrap(lanjutBtn).click();
            } else {
              cy.log('"Lanjut" button not found. Skipping...');
            }
          });
        } else {
          cy.log('No warning visible. Skipping...');
        }
      });

      // Click Selanjutnya

      cy.wait(2000);
      // Select Provinsi
      // 2️⃣ After selecting the day, click the "Pilih" button
      cy.contains('button', 'Pilih')
        .should('be.visible')
        .click();
      cy.wait(1000);
      // 3️⃣ Click the "Daftarkan dengan NIK" button
      cy.contains('button', 'Daftarkan dengan NIK', { timeout: 30000 })
        .should('be.visible')
        .click();

      cy.wait(2000);
      // Check for specific warning "Data belum sesuai KTP"

      cy.get('.p-2').then(($popup) => {
        // Grab any text inside .pb-2 or fallback to a default message
        const resultMessage = $popup.find('div.pb-2').text().trim() || 'No status message';
        cy.log(`📋 Result for ${nik}: ${resultMessage}`);

        // Always write to Excel with whatever result message is shown
        cy.task('writeExcelRemark', {
          filePath: 'cypress/e2e/Code/Data/2025/Apr/DataCatingRport.xlsx',
          nik,
          remark: resultMessage
        });

        // If it's an error that requires going back
        if (resultMessage.includes('Data belum sesuai KTP')) {
          cy.contains('button', 'Perbaiki data').click();
          return;
        }

        // Save to summary array for final export
        reportSummary.push({
          TanggalPemeriksaan: formData['TanggalPemeriksaan'],
          Nama: formData['Nama'],
          NIK: formData['NIK'],
          remark: resultMessage
        });

      }).then(() => {
        // Export to Excel (overwrite or append logic based on how you're using it)
        cy.task('exportReportingData', {
          filePath: 'cypress/e2e/Code/Data/2025/Apr/DataCatingRport.xlsx',
          rows: reportSummary
        });


        cy.wait(2000);
        // Click the "Perbaiki data" button to reset or return
        // cy.contains('button', 'Kembali').click();
        // return;

        //   cy.wrap(this.formDataArray).each((formData) => {
        //     // Your form handling logic...
        //     // Then inside the loop, push to summary:
        //     reportSummary.push({
        //       TanggalPemeriksaan: formData['TanggalPemeriksaan'],
        //       Nama: formData['Nama'],
        //       NIK: formData['NIK'],
        //       remark: 'Success' // Or based on logic
        //     });
        //   }).then(() => {
        //     // ✅ Export after all rows processed
        //     cy.task('exportReportingData', {
        //       filePath: 'cypress/e2e/Code/Data/2025/Apr/DataCatingRport.xlsx',
        //       rows: reportSummary
        //     });
        //   });
      });

      //   reportSummary.push({
      //     TanggalPemeriksaan: formData['TanggalPemeriksaan'],
      //     Nama: formData['Nama'],
      //     NIK: formData['NIK'],
      //     remark: 'Some condition result'
      //   });

      //   // ... later, after the loop ...
      // }).then(() => {
      //   cy.task('exportReportingData', {

      //     filePath: 'cypress/e2e/Code/Data/2025/Apr/DataCatingRport.xlsx',
      //     rows: reportSummary
      //   });


      // If Kuota warning popup is shown, click "Lanjut"
      // if ($popup.find('div.pb-2:contains("Kuota Pemeriksaan Habis")').length > 0) {
      //   cy.contains('button', 'Lanjut').should('be.visible').click();
      // }
    });
  });
});
