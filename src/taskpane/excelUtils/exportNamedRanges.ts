import { getNamedRanges } from './common';

const WORKSHEET_NAME = 'Existing Names';

const exportNamedRangesToWorksheet = async (): Promise<{
  success: boolean;
  errorCode: string;
  count: number;
}> => {
  const status = await Excel.run(async (context) => {
    const namesMap = await getNamedRanges();

    if (namesMap === null) {
      throw new Error('FailedToGetNames');
    }

    const rows = namesMap.items.map(({
      name, formula, type, scope,
    }, index) => [index + 1, name, formula, type, scope]);

    const values = [
      ['ID', 'Name', 'Formula', 'Type', 'Scope'],
      ...rows,
    ];

    const worksheetToCreate = context.workbook.worksheets.getItemOrNullObject(WORKSHEET_NAME);
    await context.sync();

    let sheet: Excel.Worksheet;

    if (!worksheetToCreate.isNullObject) {
      sheet = context.workbook.worksheets.getItem(WORKSHEET_NAME);
      sheet.activate();
      sheet.getRange().clear();
      await context.sync();
    } else {
      sheet = context.workbook.worksheets.add(WORKSHEET_NAME);
      sheet.activate();
      await context.sync();
    }

    const range = sheet.getRange('B2').getResizedRange(values.length - 1, values[0].length - 1);

    // Format cell to "Text"
    range.numberFormat = [['@']];

    // Write values
    range.values = values;

    // Cell borders
    range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
    range.format.borders.getItem('InsideVertical').style = 'Continuous';
    range.format.borders.getItem('EdgeBottom').style = 'Continuous';
    range.format.borders.getItem('EdgeLeft').style = 'Continuous';
    range.format.borders.getItem('EdgeRight').style = 'Continuous';
    range.format.borders.getItem('EdgeTop').style = 'Continuous';

    // Header styles
    const headerRange = range.getRow(0);
    headerRange.format.font.bold = true;
    headerRange.format.font.color = '#ffffff';
    headerRange.format.fill.color = '#201A3D';

    // Column width
    sheet.getRange('A1').format.columnWidth = 18;
    range.format.autofitColumns();

    // Error styles
    rows.filter((row) => row[3] === 'Error').forEach((row) => {
      const errorRange = range.getRow(+row[0]);
      errorRange.format.font.bold = true;
      errorRange.format.font.color = '#f0533d';
    });
    return { success: true, errorCode: '', count: values.length };
  }).catch((error) => {
    console.error(error);
    return { success: false, errorCode: error.message.split(':')[0], count: 0 };
  });

  return status;
};

export default exportNamedRangesToWorksheet;
