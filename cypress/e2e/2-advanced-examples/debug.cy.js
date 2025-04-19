describe('Read Excel File', () => {
 it('Logs data from Excel', () => {
  cy.task('readExcel', 'cypress/fixtures/Data/2025/Apr/Data.xlsx').then((data) => {
   cy.log(JSON.stringify(data));
  });
 });
});
