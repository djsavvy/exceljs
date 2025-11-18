const Enums = verquire('doc/enums');
const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('Data Tables (What-If Analysis)', () => {
    describe('Data Table - Read and Write', () => {
      it('should read data table with shareType and ref', () => {
        const wb = new ExcelJS.Workbook();
        return wb.xlsx
          .readFile('./spec/integration/data/data-table.xlsx')
          .then(() => {
            const ws = wb.getWorksheet('Sheet1');

            // Test data table master cell D5
            const cellD5 = ws.getCell('D5');
            expect(cellD5.type).to.equal(Enums.ValueType.Formula);
            expect(cellD5.value.shareType).to.equal('dataTable');
            expect(cellD5.value.ref).to.equal('D5:E6');

            // Verify data table attributes are present on master cell
            expect(cellD5.value.r1).to.be.ok();
            expect(cellD5.value.r2).to.be.ok();
            expect(cellD5.value.dt2D).to.be.ok();

            // Test ALL cells in the data table range have the attributes
            const cellE5 = ws.getCell('E5');
            expect(cellE5.type).to.equal(Enums.ValueType.Formula);
            expect(cellE5.value.shareType).to.equal('dataTable');
            expect(cellE5.value.ref).to.equal('D5:E6'); // All cells have ref
            expect(cellE5.value.r1).to.be.ok();
            expect(cellE5.value.r2).to.be.ok();
            expect(cellE5.value.dt2D).to.be.ok();
            expect(cellE5.value.result).to.equal('sbsb'); // Cached value preserved

            const cellD6 = ws.getCell('D6');
            expect(cellD6.type).to.equal(Enums.ValueType.Formula);
            expect(cellD6.value.shareType).to.equal('dataTable');
            expect(cellD6.value.ref).to.equal('D5:E6'); // All cells have ref
            expect(cellD6.value.r1).to.be.ok();
            expect(cellD6.value.r2).to.be.ok();
            expect(cellD6.value.dt2D).to.be.ok();
            expect(cellD6.value.result).to.equal('ablabl'); // Cached value preserved

            const cellE6 = ws.getCell('E6');
            expect(cellE6.type).to.equal(Enums.ValueType.Formula);
            expect(cellE6.value.shareType).to.equal('dataTable');
            expect(cellE6.value.ref).to.equal('D5:E6'); // All cells have ref
            expect(cellE6.value.r1).to.be.ok();
            expect(cellE6.value.r2).to.be.ok();
            expect(cellE6.value.dt2D).to.be.ok();
            expect(cellE6.value.result).to.equal('sblsbl'); // Cached value preserved

            // Check the formula cell C4
            const cellC4 = ws.getCell('C4');
            expect(cellC4.type).to.equal(Enums.ValueType.Formula);
            expect(cellC4.value.formula).to.equal('A1&B1&A1&B1');
          });
      });

      it('should write and re-read data table preserving attributes', () => {
        const wb = new ExcelJS.Workbook();
        return wb.xlsx
          .readFile('./spec/integration/data/data-table.xlsx')
          .then(() => {
            const ws = wb.getWorksheet('Sheet1');
            const originalCell = ws.getCell('D5');
            const originalValue = {...originalCell.value};

            // Write to a buffer
            return wb.xlsx
              .writeBuffer()
              .then(buffer => ({buffer, originalValue}));
          })
          .then(({buffer, originalValue}) => {
            // Read from buffer
            const wb2 = new ExcelJS.Workbook();
            return wb2.xlsx.load(buffer).then(() => ({wb2, originalValue}));
          })
          .then(({wb2, originalValue}) => {
            const ws2 = wb2.getWorksheet('Sheet1');
            const cell2 = ws2.getCell('D5');

            // Verify data table attributes are preserved
            expect(cell2.type).to.equal(Enums.ValueType.Formula);
            expect(cell2.value.shareType).to.equal('dataTable');
            expect(cell2.value.ref).to.equal('D5:E6');

            // Check that all attributes from original are preserved
            if (originalValue.r1) {
              expect(cell2.value.r1).to.equal(originalValue.r1);
            }
            if (originalValue.r2) {
              expect(cell2.value.r2).to.equal(originalValue.r2);
            }
            if (originalValue.dt2D) {
              expect(cell2.value.dt2D).to.equal(originalValue.dt2D);
            }
          });
      });
    });
  });
});
