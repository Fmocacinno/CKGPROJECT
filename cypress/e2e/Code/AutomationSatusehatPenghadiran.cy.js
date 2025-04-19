const checkedData = [];

Cypress.on('fail', (error, runnable) => {
  if (checkedData.length > 0) {
    cy.task('writeToExcel', {
      filePath: 'cypress/results/ReportingOnFail.xlsx',
      data: checkedData
    });
  }
  throw error;
});

describe('Form Automation from Excel (Multiple Rows)', () => {
  let formDataArray = [];
  const checkedData = []; // Array to store checked data for export

  function selectYearAndMonthFromPicker(dob) {
    const [year, month, day] = dob.split('-');
    const monthIndex = parseInt(month, 10) - 1;

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
    cy.get(`table.mx-table-month td.cell[data-month="${monthIndex}"]`).click();
    cy.get(`.mx-table-date td.cell[title="${dob}"]`).click();
  }

  function processTableAndClick() {
    function handleNextPage() {
      cy.get('li.page-item a.page-link').contains('>').then(($next) => {
        if (!$next.parent().hasClass('disabled')) {
          cy.wrap($next).click();
          cy.wait(1000); // wait for rows to load
          processRows(); // call next page
        }
      });
    }

    function processRows() {
      cy.get('tbody tr').each(($row, index, $list) => {
        cy.get('tbody tr').eq(index).then($tr => {
          const statusText = $tr.find('td').eq(6).text().trim();
          if (statusText.includes('Belum Pemeriksaan')) {
            const no = $tr.find('td').eq(0).text().trim();
            const nama = $tr.find('td').eq(1).text().trim();
            const tanggalLahir = $tr.find('td').eq(2).text().trim();
            const nomorTiket = $tr.find('td').eq(3).text().trim();
            const pelayanan = statusText; // "Belum Pemeriksaan"

            // Add the collected data to the checkedData array
            checkedData.push({
              No: no,
              Nama: nama,
              TanggalLahir: tanggalLahir,
              NomorTiket: nomorTiket,
              Pelayanan: pelayanan
            });

            // Click "Mulai" and "Mulai Pemeriksaan"
            cy.wrap($tr).find('button').contains('Mulai').as('mulaiBtn');
            cy.get('@mulaiBtn').click();
            cy.wait(1000);
            cy.get('button')
              .contains('Mulai Pemeriksaan')
              .should('be.visible')
              .click();
            cy.wait(1000);
            // Wait then click back arrow image
            cy.get('img[src="/images/icons/icon-arrow-left.svg"]')
              .should('be.visible')
              .click();

            cy.wait(1000); // wait for return to table
          }
        });

        if ((index + 1) === $list.length && $list.length === 10) {
          handleNextPage();
        }
      });
    }

    processRows(); // first page
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
      const gender = formData['JenisKelamin'];
      const dateStr = formData['TanggalPemeriksaan'];

      if (!dateStr || isNaN(new Date(dateStr))) {
        cy.log(`Invalid or missing TanggalPemeriksaan for NIK ${nik}`);
        return;
      }

      if (submittedNIKs.has(nik)) {
        cy.log(`NIK ${nik} has already been submitted. Skipping this entry.`);
        return;
      }

      submittedNIKs.add(nik);
      cy.wait(2000);
      cy.visit('https://sehatindonesiaku.kemkes.go.id/ckg-pelayanan');
      cy.wait(1000);

      cy.contains('button', 'Simpan').click();
      cy.wait(1000);

      // ✅ Use the improved stable row-check function
      processTableAndClick();
    });

    // After all rows processed, export to Excel
    Cypress.on('fail', (error, runnable) => {
      if (checkedData.length > 0) {
        cy.task('writeToExcel', {
          filePath: 'cypress/results/ReportingOnFail.xlsx',
          data: checkedData
        });
      }
      throw error; // let the test fail as normal
    });
  });
});
