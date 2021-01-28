import {
  insertForm, deleteForm, getNamedRanges,
} from './common';

const WORKSHEET_NAME = 'Edit Names Wizard';
const FORM_NAMED_RANGE = 'EDIT_NAMED_RANGE_FORM';
const FORM_NAMED_RANGE_ADDRESS = '$A2:$G9999';

export interface EditNamedRangesOperationResult {
  oldName: string;
  oldFormula: string;
  oldScope: string;
  newName: string;
  newFormula: string;
  newScope: string;
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
      .filter((row) => Boolean(row[0] && row[1] && (row[4] || row[5] || row[6])))
      .map((row) => ({
        oldName: row[0],
        oldFormula: row[1],
        oldScope: row[3],
        newName: row[4],
        newFormula: row[5],
        newScope: row[6],
        runtimeError: '',
      }));

    if (!rangesToEdit.length) {
      return { success: false, errorCode: 'NothingToEdit', operationResult: [] };
    }

    const failedOperation = [];

    for (const rangeItem of rangesToEdit) {
      try {
        const {
          oldName, oldFormula, oldScope, newName, newFormula, newScope,
        } = rangeItem;

        const oldNameContext = (oldScope === 'Workbook'
          ? context.workbook
          : context.workbook.worksheets.getItem(oldScope)).names;

        oldNameContext.getItem(oldName).delete();

        const scopeToAdd = newScope || oldScope;
        const nameContext = (scopeToAdd === 'Workbook'
          ? context.workbook
          : context.workbook.worksheets.getItem(scopeToAdd)).names;

        nameContext.add(
          newName || oldName,
          newFormula || oldFormula,
        );

        await context.sync();
      } catch (error) {
        // Rollback
        const oldNameContext = (rangeItem.oldScope === 'Workbook'
          ? context.workbook
          : context.workbook.worksheets.getItem(rangeItem.oldScope)).names;
        oldNameContext.add(rangeItem.oldName, rangeItem.oldFormula);
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
