type NamedRangeType = {
  comment: string;
  formula: string;
  name: string;
  scope: string;
  type: string;
  value: string;
  visible: boolean;
};

interface NamedRanges {
  items: NamedRangeType[];
}

const getNamedRanges = async (): Promise<void> => {
  await Excel.run(async (context) => {
    const SHEET_NAME = 'Existing Names';
    const { names } = context.workbook;
    names.load();
    await context.sync();

    const namesMap: NamedRanges = JSON.parse(JSON.stringify(names));
    const rows = namesMap.items.map(({
      name, formula, type, scope,
    }, index) => [index + 1, name, `'${formula}`, type, scope]);
    const values = [['ID', 'Existing name', 'Existing Fomular', 'Type', 'Scope'], ...rows];

    let sheet;

    try {
      sheet = context.workbook.worksheets.getItem(SHEET_NAME);
      sheet.activate();
      sheet.getRange().clear();
      await context.sync();
    } catch (error) {
      // ItemNotFound
      sheet = context.workbook.worksheets.add(SHEET_NAME);
      sheet.activate();
      await context.sync();
    }

    const range = sheet.getRange('B2').getResizedRange(values.length - 1, values[0].length - 1);
    range.values = values;
    sheet.getRange('A1').format.columnWidth = 18;

    // Cell borders
    range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
    range.format.borders.getItem('InsideVertical').style = 'Continuous';
    range.format.borders.getItem('EdgeBottom').style = 'Continuous';
    range.format.borders.getItem('EdgeLeft').style = 'Continuous';
    range.format.borders.getItem('EdgeRight').style = 'Continuous';
    range.format.borders.getItem('EdgeTop').style = 'Continuous';

    // Header Styles
    const headerRange = range.getRow(0);
    headerRange.format.font.bold = true;
    headerRange.format.font.color = '#ffffff';
    headerRange.format.fill.color = '#201A3D';

    // Column Width
    [18, 300, 300, 100, 100].forEach((width, index) => {
      range.getColumn(index).format.columnWidth = width;
    });

    // Error styles
    rows.filter((row) => row[3] === 'Error').forEach((row) => {
      const errorRange = range.getRow(+row[0]);
      errorRange.format.font.bold = true;
      errorRange.format.font.color = '#f0533d';
    });
  }).catch((error) => {
    console.error(error);
  });
};

export default getNamedRanges;
