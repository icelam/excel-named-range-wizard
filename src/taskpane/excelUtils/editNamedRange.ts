import {
  insertForm, deleteForm, getNamedRanges,
} from './common';

const WORKSHEET_NAME = 'Edit Names Wizard';
const FORM_NAMED_RANGE = 'EDIT_NAMED_RANGE_FORM';
const FORM_NAMED_RANGE_ADDRESS = '$A2:$F9999';

export interface EditNamedRangesOperationResult {
  oldName: string;
  oldFormula: string;
  newName: string;
  newFormula: string;
  runtimeError: string;
}

export const insertEditNamedRangesForm = async (): Promise<{
  success: boolean;
  errorCode: string;
}> => {
  const namesMap = await getNamedRanges();
  if (!namesMap) {
    return { success: false, errorCode: 'FailedToGetNames' };
  }

  if (!namesMap.items.length) {
    return { success: false, errorCode: 'NoExistingNamedRanges' };
  }

  const result = await insertForm(WORKSHEET_NAME, FORM_NAMED_RANGE, FORM_NAMED_RANGE_ADDRESS);
  return result;
};

export const deleteEditNamedRangeForm = async (): Promise<{
  success: boolean;
  errorCode: string;
}> => {
  const result = await deleteForm(WORKSHEET_NAME, FORM_NAMED_RANGE);
  return result;
};

export const editNamedRange = async (): Promise<{
  success: boolean;
  errorCode: string;
  operationResult: EditNamedRangesOperationResult[];
}> => {
  const result = await Excel.run(async (context) => {
    const formRange = context.workbook.names.getItem(FORM_NAMED_RANGE).getRange();
    formRange.load('values');
    await context.sync();

    const rangesToEdit: EditNamedRangesOperationResult[] = JSON.parse(
      JSON.stringify(formRange.values),
    )
      .filter((row) => Boolean(row[0] && row[1] && (row[4] || row[5])))
      .map((row) => ({
        oldName: row[0],
        oldFormula: row[1],
        newName: row[4],
        newFormula: row[5],
        runtimeError: '',
      }));

    if (!rangesToEdit.length) {
      return { success: false, errorCode: 'NothingToEdit', operationResult: [] };
    }

    const failedOperation = [];

    // eslint-disable-next-line no-restricted-syntax
    for (const rangeItem of rangesToEdit) {
      try {
        context.workbook.names.getItem(rangeItem.oldName).delete();
        context.workbook.names.add(
          rangeItem.newName || rangeItem.oldName,
          rangeItem.newFormula || rangeItem.oldFormula,
        );
        // eslint-disable-next-line no-await-in-loop
        await context.sync();
      } catch (error) {
        // Rollback
        context.workbook.names.add(rangeItem.oldName, rangeItem.oldFormula);
        failedOperation.push({ ...rangeItem, runtimeError: error.message });
      }
    }

    if (failedOperation.length) {
      return { success: false, errorCode: 'FailedToEdit', operationResult: failedOperation };
    }

    return { success: true, errorCode: '', operationResult: [] };
  }).catch((error) => {
    console.error(error);
    return { success: false, errorCode: error.message.split(':')[0], operationResult: [] };
  });

  return result;
};
