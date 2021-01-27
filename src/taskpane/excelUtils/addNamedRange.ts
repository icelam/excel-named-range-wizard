import {
  insertForm, deleteForm, validateNamedRangesName, addFormTemplate,
} from './common';

export interface NamedRangeValidationResult {
  name: string;
  formula: string;
  validations: {
    isNameNonEmpty: boolean;
    isFormulaNonEmpty: boolean;
    isNameValid: boolean;
  };
  runtimeError: string;
}

const WORKSHEET_NAME = 'Add Names Wizard';
const FORM_NAMED_RANGE = 'ADD_NAMED_RANGE_FORM';

export const insertAddNamedRangesForm = async (): Promise<{
  success: boolean;
  errorCode: string;
}> => {
  const result = await insertForm(WORKSHEET_NAME, FORM_NAMED_RANGE);
  return result;
};

export const deleteAddNamedRangeForm = async (): Promise<{
  success: boolean;
  errorCode: string;
}> => {
  const result = await deleteForm(WORKSHEET_NAME, FORM_NAMED_RANGE);
  return result;
};

export const addNamedRange = async (): Promise<{
  success: boolean;
  errorCode: string;
  validationResult: NamedRangeValidationResult[];
}> => {
  const result = await Excel.run(async (context) => {
    const range = context.workbook.names.getItem(FORM_NAMED_RANGE).getRange();
    range.load('values');
    await context.sync();

    const rangesToAdd = JSON.parse(JSON.stringify(range.values))
      .filter((row) => !!row[0] || !!row[1])
      .map((row) => {
        const [name, formula] = row;
        return ({
          name,
          formula,
          validations: {
            isNameNonEmpty: !!name,
            isFormulaNonEmpty: !!formula,
            isNameValid: validateNamedRangesName(name),
            // isFormulaValid: true, // TODO: formula format validation
            // isUniqueName: true, // TODO: Name duplication check
          },
          runtimeError: '',
        });
      });

    if (!rangesToAdd.length) {
      return { success: false, errorCode: 'NothingToAdd', validationResult: [] };
    }

    const invalidRanges = [];
    const validRanges = [];

    rangesToAdd.forEach((rangeDefinition) => {
      const isRangeDefinationValid = Object.entries(rangeDefinition.validations)
        .every((validation) => validation[1]);

      if (isRangeDefinationValid) {
        validRanges.push(rangeDefinition);
      } else {
        invalidRanges.push(rangeDefinition);
      }
    });

    if (invalidRanges.length) {
      return { success: false, errorCode: 'InvalidInput', validationResult: invalidRanges };
    }

    const failedOperation = [];

    // eslint-disable-next-line no-restricted-syntax
    for (const validRange of validRanges) {
      try {
        context.workbook.names.add(validRange.name, validRange.formula);
        // eslint-disable-next-line no-await-in-loop
        await context.sync();
      } catch (error) {
        failedOperation.push({ ...validRange, runtimeError: error.message });
      }
    }

    if (failedOperation.length) {
      await addFormTemplate(WORKSHEET_NAME, FORM_NAMED_RANGE);

      const values = failedOperation.map(({ name, formula }) => [name, formula]);
      range.getCell(0, 0).getResizedRange(values.length - 1, values[0].length - 1).values = values;
      range.getCell(0, 0).select();
      return { success: false, errorCode: 'FailedToAdd', validationResult: failedOperation };
    }

    return { success: true, errorCode: '', validationResult: invalidRanges };
  }).catch((error) => {
    console.error(error);
    return { success: false, errorCode: error.message.split(':')[0], validationResult: [] };
  });

  return result;
};
