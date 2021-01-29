import React, {
  FC, useState, useMemo, useEffect,
} from 'react';
import styled from 'styled-components';
import { useIntl } from 'react-intl';
import {
  DefaultButton, DialogFooter,
} from 'office-ui-fabric-react';
import {
  Header, Modal, Progress, HeroList, HeroListItem,
} from '../components';
import {
  exportNamedRangesToWorksheet,
  validateNamedRanges,
  NamedRangeType,
  insertAddNamedRangesForm,
  deleteAddNamedRangeForm,
  addNamedRange,
  NamedRangeValidationResult,
  insertEditNamedRangesForm,
  deleteEditNamedRangeForm,
  editNamedRange,
  EditNamedRangesOperationResult,
} from '../excelUtils';

const LoadingModal = styled<{isLoading: boolean}>(Progress)`
  position: fixed;
  top: 0;
  left: 0;
  background-color: rgba(255, 255, 255, 0.5);
  ${(props) => !props.isLoading && 'display: none;'}
`;

const MainWrapper = styled.main`
  display: flex;
  flex-direction: column;
  flex-wrap: nowrap;
  align-items: center;
  flex: 1 0 0;
  padding: 0.625rem 1.25rem;
`;

const Title = styled.h2`
  width: 100%;
  text-align: center;
  font-size: 1.3125rem;
  font-weight: 300;
`;

const Description = styled.p`
  text-align: center;
  font-size: 0.75rem;
`;

const FullwidthButton = styled(DefaultButton)`
  width: 100%;
  margin: 0.5rem;
`;

const ModalDetailsWrapper = styled.div`
  max-height: 200px;
  overflow-y: auto;
  background-color: ${(props) => props.theme.color.background};
  padding: 0 1rem;
`;

const Home: FC = () => {
  const intl = useIntl();
  const [isLoading, setIsLoading] = useState(false);

  /**
   * Error Dialog
   */
  const [errorMessage, setErrorMessage] = useState('');
  const [isErrorDialogOpen, setIsErrorDialogOpen] = useState(false);

  useEffect(() => {
    setIsErrorDialogOpen(!!errorMessage);
  }, [errorMessage]);

  const hideErrorDialog = (): void => {
    setErrorMessage('');
  };

  const errorDialog = (
    <Modal
      onDismiss={hideErrorDialog}
      title={intl.formatMessage({ id: 'app.error.title' })}
      isOpen={isErrorDialogOpen}
      modalId="validateModal"
      theme="failure"
    >
      <p>{errorMessage}</p>

      <DialogFooter>
        <DefaultButton onClick={hideErrorDialog} text={intl.formatMessage({ id: 'app.modal.ok' })} />
      </DialogFooter>
    </Modal>
  );

  /**
   * Export Named Ranges to excel worksheet
   */
  const exportNamedRanges = async (): Promise<void> => {
    setIsLoading(true);
    const { success: isSuccess, errorCode } = await exportNamedRangesToWorksheet();
    if (!isSuccess) {
      setErrorMessage(intl.formatMessage({ id: 'app.function.export.error' }, { ERROR_CODE: errorCode }));
    }
    setIsLoading(false);
  };

  /**
   * Add Named Ranges Modal
   */
  const [isAddNamesModalOpen, setIsAddNamesModalOpen] = useState(false);
  const [
    addNamesErrorCode,
    setAddNamesErrorCode,
  ] = useState<'' | 'InvalidInput' | 'FailedToAdd' | 'NothingToAdd'>('');
  const [
    addNamesValidationResult,
    setAddNamesValidationResult,
  ] = useState<NamedRangeValidationResult[]>([]);
  const [isAddNamesSuccess, setIsAddNamesSuccess] = useState(false);

  const showAddNamesModal = async (): Promise<void> => {
    setIsLoading(true);
    const { success: isInsertSuccess, errorCode } = await insertAddNamedRangesForm();
    if (!isInsertSuccess) {
      setErrorMessage(intl.formatMessage({ id: 'app.function.add.error.insertForm' }, { ERROR_CODE: errorCode }));
      setIsLoading(false);
      return;
    }
    setIsAddNamesModalOpen(true);
    setIsLoading(false);
  };

  const hideAddNamesModal = async (shouldDeleteWorkSheet = true): Promise<void> => {
    setIsAddNamesModalOpen(false);
    setAddNamesErrorCode('');
    setAddNamesValidationResult([]);
    setIsAddNamesSuccess(false);
    setIsLoading(true);
    if (shouldDeleteWorkSheet) {
      await deleteAddNamedRangeForm();
    }
    setIsLoading(false);
  };

  const closeAddNameSuccessModal = (): void => {
    hideAddNamesModal(false);
  };

  const discardAddNameChanges = (): void => {
    hideAddNamesModal(true);
  };

  const addNewNames = async (): Promise<void> => {
    const { success: isSuccess, errorCode, validationResult } = await addNamedRange();
    if (!isSuccess) {
      if (errorCode === 'FailedToAdd' || errorCode === 'InvalidInput' || errorCode === 'NothingToAdd') {
        setAddNamesErrorCode(errorCode);
        setAddNamesValidationResult(validationResult);
        return;
      }

      setErrorMessage(intl.formatMessage({ id: 'app.function.add.error' }, { ERROR_CODE: errorCode }));
      await hideAddNamesModal();
      return;
    }

    await deleteAddNamedRangeForm();
    setIsAddNamesSuccess(true);
  };

  const addNamesHowTo = (
    <>
      <p>{intl.formatMessage({ id: 'app.function.add.howTo' })}</p>
      <ModalDetailsWrapper>
        <HeroList
          items={Array.from({ length: 7 }, (_, index) => ({
            primaryText: intl.formatMessage({ id: `app.function.namedRange.tips${index + 1}` }),
            icon: 'RadioBullet',
          }))}
        />
      </ModalDetailsWrapper>
      <DialogFooter>
        <DefaultButton onClick={discardAddNameChanges} text={intl.formatMessage({ id: 'app.modal.cancel' })} />
        <DefaultButton onClick={addNewNames} text={intl.formatMessage({ id: 'app.modal.add' })} />
      </DialogFooter>
    </>
  );

  const addNamesErrorDetailsList = useMemo(() => {
    if (addNamesErrorCode === 'InvalidInput') {
      const errorList: HeroListItem[] = [];

      if (!addNamesValidationResult.every(({ validations }) => validations.isNameNonEmpty)) {
        errorList.push({
          primaryText: intl.formatMessage({ id: 'app.function.add.error.missingName' }),
          icon: 'IncidentTriangle',
        });
      }

      addNamesValidationResult.forEach(({ name, validations }) => {
        if (name && !validations.isNameValid) {
          errorList.push({
            primaryText: intl.formatMessage({ id: 'app.function.add.error.invalidName' }, { NAME: name }),
            icon: 'IncidentTriangle',
          });
        }

        if (name && !validations.isFormulaNonEmpty) {
          errorList.push({
            primaryText: intl.formatMessage({ id: 'app.function.add.error.missingFormula' }, { NAME: name }),
            icon: 'IncidentTriangle',
          });
        }
      });

      return errorList;
    }

    if (addNamesErrorCode === 'FailedToAdd') {
      return addNamesValidationResult.map(({ name, runtimeError }) => ({
        primaryText: `${name}: ${runtimeError}`,
        icon: 'IncidentTriangle',
      }));
    }

    return [];
  }, [addNamesValidationResult, addNamesErrorCode]);

  const addNamesErrorDetails = (
    <>
      <p>
        {addNamesErrorCode && intl.formatMessage({
          id: `app.function.add.error.${addNamesErrorCode}`,
        })}
      </p>
      {
        addNamesValidationResult.length
          ? (
            <ModalDetailsWrapper>
              <HeroList items={addNamesErrorDetailsList} />
            </ModalDetailsWrapper>
          )
          : null
      }
      <DialogFooter>
        <DefaultButton onClick={discardAddNameChanges} text={intl.formatMessage({ id: 'app.modal.cancel' })} />
        <DefaultButton onClick={addNewNames} text={intl.formatMessage({ id: 'app.modal.retry' })} />
      </DialogFooter>
    </>
  );

  const addNamesSuccessMessage = (
    <>
      <p>{intl.formatMessage({ id: 'app.function.add.success' })}</p>
      <DialogFooter>
        <DefaultButton onClick={closeAddNameSuccessModal} text={intl.formatMessage({ id: 'app.modal.ok' })} />
      </DialogFooter>
    </>
  );

  const addNamesModal = (
    <Modal
      onDismiss={discardAddNameChanges}
      title={intl.formatMessage({ id: 'app.function.add' })}
      isOpen={isAddNamesModalOpen}
      modalId="addNameModal"
      isClosable={false}
      theme={isAddNamesSuccess ? 'success' : addNamesErrorCode ? 'failure' : undefined}
    >
      {
        isAddNamesSuccess
          ? addNamesSuccessMessage
          : addNamesErrorCode
            ? addNamesErrorDetails
            : addNamesHowTo
      }
    </Modal>
  );

  /**
   * Vaidate Name ranges
   */
  const [isValidateNamedRangesDialogOpen, setIsValidateNamedRangesDialogOpen] = useState(false);
  const [invalidNamedRanges, setInvalidNamedRanges] = useState<NamedRangeType[]>([]);

  const validateAndShowInvalidNamedRanges = async (): Promise<void> => {
    setIsLoading(true);
    const {
      success: isSuccess, errorCode, invalidNamedRanges: names,
    } = await validateNamedRanges();

    if (!isSuccess) {
      setErrorMessage(intl.formatMessage({ id: 'app.function.validate.error' }, { ERROR_CODE: errorCode }));
      setIsLoading(false);
      return;
    }

    setIsValidateNamedRangesDialogOpen(true);
    setInvalidNamedRanges(names);
    setIsLoading(false);
  };

  const hideValidateNamedRangesDialog = (): void => {
    setIsValidateNamedRangesDialogOpen(false);
  };

  const invalidNamedRangesList = useMemo(() => invalidNamedRanges.map(({ name }) => (
    { icon: 'IncidentTriangle', primaryText: name }
  )), [invalidNamedRanges]);

  const validateNamedRangesDialog = (
    <Modal
      onDismiss={hideValidateNamedRangesDialog}
      title={intl.formatMessage({ id: 'app.function.validate' })}
      isOpen={isValidateNamedRangesDialogOpen}
      modalId="validateModal"
      theme={invalidNamedRanges.length ? 'failure' : 'success'}
    >
      <p>
        {invalidNamedRanges.length
          ? intl.formatMessage(
            { id: `app.function.validate.failed.${invalidNamedRanges.length > 1 ? 'other' : 'one'}` },
            { COUNT: invalidNamedRanges.length },
          )
          : intl.formatMessage({ id: 'app.function.validate.success' })}
      </p>

      {
        invalidNamedRanges.length
          ? (
            <ModalDetailsWrapper>
              <HeroList items={invalidNamedRangesList} />
            </ModalDetailsWrapper>
          )
          : null
      }

      <DialogFooter>
        <DefaultButton onClick={hideValidateNamedRangesDialog} text={intl.formatMessage({ id: 'app.modal.ok' })} />
      </DialogFooter>
    </Modal>
  );

  /**
   * Edit Name ranges
   */
  const [isEditNamesModalOpen, setIsEditNamesModalOpen] = useState(false);
  const [
    editNamesErrorCode,
    setEditNamesErrorCode,
  ] = useState<'' | 'FailedToEdit' | 'NothingToEdit'>('');
  const [
    editNamesOperationResult,
    setEditNamesOperationResult,
  ] = useState<EditNamedRangesOperationResult[]>([]);
  const [isEditNamesSuccess, setIsEditNamesSuccess] = useState(false);

  const showEditNamesModal = async (): Promise<void> => {
    setIsLoading(true);
    const { success: isInsertSuccess, errorCode } = await insertEditNamedRangesForm();
    if (!isInsertSuccess) {
      if (errorCode === 'NoExistingNamedRanges') {
        setErrorMessage(intl.formatMessage({ id: 'app.function.edit.error.NoExistingNamedRanges' }));
      } else {
        setErrorMessage(intl.formatMessage({ id: 'app.function.edit.error.insertForm' }, { ERROR_CODE: errorCode }));
      }
      setIsLoading(false);
      return;
    }
    setIsEditNamesModalOpen(true);
    setIsLoading(false);
  };

  const hideEditNamesModal = async (shouldDeleteWorkSheet = true): Promise<void> => {
    setIsEditNamesModalOpen(false);
    setEditNamesErrorCode('');
    setEditNamesOperationResult([]);
    setIsEditNamesSuccess(false);
    setIsLoading(true);
    if (shouldDeleteWorkSheet) {
      await deleteEditNamedRangeForm();
    }
    setIsLoading(false);
  };

  const closeEditNameSuccessModal = (): void => {
    hideEditNamesModal(false);
  };

  const discardEditNameChanges = (): void => {
    hideEditNamesModal(true);
  };

  const editNames = async (): Promise<void> => {
    const { success: isSuccess, errorCode, operationResult } = await editNamedRange();
    if (!isSuccess) {
      if (errorCode === 'FailedToEdit' || errorCode === 'NothingToEdit') {
        setEditNamesErrorCode(errorCode);
        setEditNamesOperationResult(operationResult);
        return;
      }

      setErrorMessage(intl.formatMessage({ id: 'app.function.edit.error' }, { ERROR_CODE: errorCode }));
      await hideEditNamesModal();
      return;
    }

    await deleteEditNamedRangeForm();
    setIsEditNamesSuccess(true);
  };

  const editNamesHowTo = (
    <>
      <p>{intl.formatMessage({ id: 'app.function.edit.howTo' })}</p>
      <ModalDetailsWrapper>
        <HeroList
          items={Array.from({ length: 7 }, (_, index) => ({
            primaryText: intl.formatMessage({ id: `app.function.namedRange.tips${index + 1}` }),
            icon: 'RadioBullet',
          }))}
        />
      </ModalDetailsWrapper>
      <DialogFooter>
        <DefaultButton onClick={discardEditNameChanges} text={intl.formatMessage({ id: 'app.modal.cancel' })} />
        <DefaultButton onClick={editNames} text={intl.formatMessage({ id: 'app.modal.edit' })} />
      </DialogFooter>
    </>
  );

  const editNamesErrorDetailsList = useMemo(() => {
    if (editNamesErrorCode === 'FailedToEdit') {
      return editNamesOperationResult.map(({ oldName, runtimeError }) => ({
        primaryText: intl.formatMessage(
          { id: 'app.function.edit.error.runtimeError' },
          { NAME: oldName, ERROR: runtimeError },
        ),
        icon: 'IncidentTriangle',
      }));
    }

    return [];
  }, [editNamesOperationResult, editNamesErrorCode]);

  const editNamesErrorDetails = (
    <>
      <p>
        {editNamesErrorCode && intl.formatMessage({
          id: `app.function.edit.error.${editNamesErrorCode}`,
        })}
      </p>
      {
        editNamesOperationResult.length
          ? (
            <ModalDetailsWrapper>
              <HeroList items={editNamesErrorDetailsList} />
            </ModalDetailsWrapper>
          )
          : null
      }
      <DialogFooter>
        {
          editNamesErrorCode === 'NothingToEdit'
            ? (
              <>
                <DefaultButton onClick={discardEditNameChanges} text={intl.formatMessage({ id: 'app.modal.close' })} />
                <DefaultButton onClick={editNames} text={intl.formatMessage({ id: 'app.modal.retry' })} />
              </>
            )
            : <DefaultButton onClick={discardEditNameChanges} text={intl.formatMessage({ id: 'app.modal.terminate' })} />
        }
      </DialogFooter>
    </>
  );

  const editNamesSuccessMessage = (
    <>
      <p>{intl.formatMessage({ id: 'app.function.edit.success' })}</p>
      <DialogFooter>
        <DefaultButton onClick={closeEditNameSuccessModal} text={intl.formatMessage({ id: 'app.modal.ok' })} />
      </DialogFooter>
    </>
  );

  const editNamesModal = (
    <Modal
      onDismiss={discardEditNameChanges}
      title={intl.formatMessage({ id: 'app.function.edit' })}
      isOpen={isEditNamesModalOpen}
      modalId="editNameModal"
      isClosable={false}
      theme={isEditNamesSuccess ? 'success' : editNamesErrorCode ? 'failure' : undefined}
    >
      {
        isEditNamesSuccess
          ? editNamesSuccessMessage
          : editNamesErrorCode
            ? editNamesErrorDetails
            : editNamesHowTo
      }
    </Modal>
  );

  /**
   * Home UI
   */
  return (
    <div>
      <Header
        logo="assets/logo-filled.png"
        title={intl.formatMessage({ id: 'app.name' })}
        message={intl.formatMessage({ id: 'app.home.welcome' })}
      />
      <MainWrapper>
        <Title>{intl.formatMessage({ id: 'app.home.tagline' })}</Title>
        <Description>{intl.formatMessage({ id: 'app.home.description' })}</Description>
        <FullwidthButton
          iconProps={{ iconName: 'Download' }}
          onClick={exportNamedRanges}
        >
          {intl.formatMessage({ id: 'app.function.export' })}
        </FullwidthButton>
        <FullwidthButton
          iconProps={{ iconName: 'PageAdd' }}
          onClick={showAddNamesModal}
        >
          {intl.formatMessage({ id: 'app.function.add' })}
        </FullwidthButton>
        <FullwidthButton
          iconProps={{ iconName: 'PageEdit' }}
          onClick={showEditNamesModal}
        >
          {intl.formatMessage({ id: 'app.function.edit' })}
        </FullwidthButton>
        <FullwidthButton
          iconProps={{ iconName: 'Zoom' }}
          onClick={validateAndShowInvalidNamedRanges}
        >
          {intl.formatMessage({ id: 'app.function.validate' })}
        </FullwidthButton>
      </MainWrapper>
      <LoadingModal isLoading={isLoading} message="Loading..." />
      {addNamesModal}
      {editNamesModal}
      {errorDialog}
      {validateNamedRangesDialog}
    </div>
  );
};

export default Home;
