const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

async function createNamespacedWorkbookBuffer() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Knowledge Base');

  worksheet.getCell('A1').value = 'Question';
  worksheet.getCell('B1').value = 'Short Answer';
  worksheet.getCell('A2').value = 'Describe your company.';
  worksheet.getCell('B2').value = 'Short company description';

  worksheet.addTable({
    name: 'SecurityQuestionnaireKB',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    columns: [{name: 'Question'}, {name: 'Short Answer'}],
    rows: [['Describe your company.', 'Short company description']],
  });

  const baseBuffer = await workbook.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(baseBuffer);

  await Promise.all(
    Object.keys(zip.files).map(async filename => {
      if (!filename.endsWith('.xml')) {
        return;
      }

      const file = zip.file(filename);
      const xml = await file.async('string');

      if (xml.includes('xmlns:x=')) {
        return;
      }

      const rootTagMatch = xml.match(/^<\?xml[^>]*\?><([A-Za-z][\w.-]*)/);
      if (!rootTagMatch) {
        return;
      }

      zip.file(
        filename,
        xml
          .replace(
            rootTagMatch[0],
            rootTagMatch[0].replace(
              `<${rootTagMatch[1]}`,
              `<x:${rootTagMatch[1]} xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"`
            )
          )
          .replace(
            /<(\/?)(?!\?)([A-Za-z][\w.-]*)(?=[\s/>])/g,
            (match, slash, tag) => `<${slash}x:${tag}`
          )
      );
    })
  );

  const sheetRelsEntry = zip.file('xl/worksheets/_rels/sheet1.xml.rels');
  const sheetRelsXml = await sheetRelsEntry.async('string');

  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    sheetRelsXml.replace('../tables/table1.xml', '/xl/tables/table1.xml')
  );

  return zip.generateAsync({type: 'nodebuffer'});
}

describe('Issue: namespaced xml elements', () => {
  it('reads xlsx files whose element names use a namespace prefix', async () => {
    const workbook = new ExcelJS.Workbook();
    const buffer = await createNamespacedWorkbookBuffer();

    await workbook.xlsx.load(buffer);

    const worksheet = workbook.getWorksheet('Knowledge Base');
    expect(worksheet).to.exist();
    expect(worksheet.getCell('A1').value).to.equal('Question');
    expect(worksheet.getCell('A2').value).to.equal('Describe your company.');
    expect(worksheet.getTable('SecurityQuestionnaireKB')).to.exist();
  });
});
