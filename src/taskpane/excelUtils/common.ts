export type NamedRangeType = {
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

export const getNamedRanges = async (): Promise<NamedRanges | null> => {
  const result = await Excel.run(async (context) => {
    const { names } = context.workbook;
    names.load();
    await context.sync();

    return JSON.parse(JSON.stringify(names));
  }).catch((error) => {
    console.error(error);
    return null;
  });

  return result;
};

export const validateNamedRangesName = (value: string): boolean => (
  value.length > 0
  && value.length <= 255
  && /[a-z_\\]/i.test(value[0])
  && (value.length === 1 || /[a-z0-9?._\\]+/i.test(value.slice(1)))
  && !value.includes(' ')
  && value !== 'R'
  && value !== 'C'
  && value !== '\\?'
  && value !== '\\\\'
);

export const addFormTemplate = async (
  worksheetName: string,
  formNamedRange: string,
): Promise<void> => {
  await Excel.run(async (context) => {
    const namesMap = await getNamedRanges();

    if (namesMap === null) {
      throw new Error('FailedToGetNames');
    }

    const rows = namesMap.items
      .filter(({ name }) => name !== formNamedRange)
      .map(({
        name, formula, type, scope,
      }) => [name, formula, type, scope, '', '']);

    const sheet = context.workbook.worksheets.getItem(worksheetName);
    sheet.activate();
    const range = sheet.getRange('$A1:$F9999');

    const values = [
      ['Current Name', 'Current Formula', 'Type', 'Scope', 'New Name', 'New Formula'],
      ...rows,
    ];
    const headerRange = range.getRow(0);

    // Format Cell to "Text"
    range.numberFormat = [['@']];

    sheet.getRange('A1').getResizedRange(values.length - 1, values[0].length - 1).values = values;

    // Header Styles
    headerRange.format.font.bold = true;
    headerRange.format.font.color = '#ffffff';
    headerRange.format.fill.color = '#3B8CFF';

    // Cell borders
    range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
    range.format.borders.getItem('InsideVertical').style = 'Continuous';
    range.format.borders.getItem('EdgeBottom').style = 'Continuous';
    range.format.borders.getItem('EdgeLeft').style = 'Continuous';
    range.format.borders.getItem('EdgeRight').style = 'Continuous';
    range.format.borders.getItem('EdgeTop').style = 'Continuous';

    // Column Width
    range.format.autofitColumns();
    range.getColumn(4).format.columnWidth = 300;
    range.getColumn(5).format.columnWidth = 300;
  });
};

export const insertForm = async (worksheetName: string, formNamedRange: string): Promise<{
  success: boolean;
  errorCode: string;
}> => {
  const result = await Excel.run(async (context) => {
    // Check if existing worksheet and named ranges exists
    const worksheetToCreate = context.workbook.worksheets.getItemOrNullObject(worksheetName);
    await context.sync();

    const rangeToCreate = context.workbook.names.getItemOrNullObject(formNamedRange);
    rangeToCreate.load();
    await context.sync();

    // create worksheet and ranges if one of it does not exist (consider as corrupted)
    let sheet;

    if (worksheetToCreate.isNullObject || rangeToCreate.isNullObject) {
      if (!worksheetToCreate.isNullObject) {
        sheet = context.workbook.worksheets.getItem(worksheetName);
        sheet.activate();
        sheet.getRange().clear();
        await context.sync();
      } else {
        sheet = context.workbook.worksheets.add(worksheetName);
        sheet.activate();
        await context.sync();
      }

      // delete range if exist and create
      if (!rangeToCreate.isNullObject) {
        context.workbook.names.getItem(formNamedRange).delete();
        await context.sync();
      }

      context.workbook.names.add(formNamedRange, sheet.getRange('$E2:$F9999'));
      await context.sync();

      await addFormTemplate(worksheetName, formNamedRange);
    }

    return { success: true, errorCode: '' };
  }).catch((error) => {
    console.error(error);
    return { success: false, errorCode: error.message.split(':')[0] };
  });

  return result;
};

export const deleteForm = async (worksheetName: string, formNamedRange: string): Promise<{
  success: boolean;
  errorCode: string;
}> => {
  const result = await Excel.run(async (context) => {
    // delete range if exist and create
    const rangeObject = context.workbook.names.getItemOrNullObject(formNamedRange);
    rangeObject.load();
    await context.sync();

    if (!rangeObject.isNullObject) {
      context.workbook.names.getItem(formNamedRange).delete();
      await context.sync();
    }

    // delete worksheet if exist
    const worksheetToDelete = context.workbook.worksheets.getItemOrNullObject(worksheetName);
    await context.sync();

    if (!worksheetToDelete.isNullObject) {
      worksheetToDelete.delete();
    }

    return { success: true, errorCode: '' };
  }).catch((error) => {
    console.error(error);
    return { success: false, errorCode: error.message.split(':')[0] };
  });

  return result;
};
