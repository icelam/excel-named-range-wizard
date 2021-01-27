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
} from '../excelUtils';

// images references in the manifest
import '../../../assets/icon-16.png';
import '../../../assets/icon-32.png';
import '../../../assets/icon-80.png';

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
  const [isAddNameModalOpen, setIsAddNameModalOpen] = useState(false);
  const [
    addNameErrorCode,
    setAddNameErrorCode,
  ] = useState<'' | 'InvalidInput' | 'FailedToAdd' | 'NothingToAdd'>('');
  const [
    addNameValidationResult,
    setAddNameValidationResult,
  ] = useState<NamedRangeValidationResult[]>([]);
  const [isAddNameSuccess, setIsAddNameSuccess] = useState(false);

  const showAddNamesModal = async (): Promise<void> => {
    setIsLoading(true);
    const { success: isInsertSuccess, errorCode } = await insertAddNamedRangesForm();
    if (!isInsertSuccess) {
      setErrorMessage(intl.formatMessage({ id: 'app.function.add.error.insertForm' }, { ERROR_CODE: errorCode }));
      setIsLoading(false);
      return;
    }
    setIsAddNameModalOpen(true);
    setIsLoading(false);
  };

  const hideAddNamesModal = async (shouldDeleteWorkSheet = true): Promise<void> => {
    setIsAddNameModalOpen(false);
    setAddNameErrorCode('');
    setAddNameValidationResult([]);
    setIsAddNameSuccess(false);
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
        setAddNameErrorCode(errorCode);
        setAddNameValidationResult(validationResult);
        return;
      }

      setErrorMessage(intl.formatMessage({ id: 'app.function.add.error' }, { ERROR_CODE: errorCode }));
      await hideAddNamesModal();
      return;
    }

    await deleteAddNamedRangeForm();
    setIsAddNameSuccess(true);
  };

  const addNameHowTo = (
    <>
      <p>{intl.formatMessage({ id: 'app.function.add.howTo' })}</p>
      <ModalDetailsWrapper>
        <HeroList
          items={Array.from({ length: 7 }, (_, index) => ({
            primaryText: intl.formatMessage({ id: `app.function.add.tips${index + 1}` }),
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

  const addNameErrorDetailsList = useMemo(() => {
    if (addNameErrorCode === 'InvalidInput') {
      const errorList: HeroListItem[] = [];

      if (!addNameValidationResult.every(({ validations }) => validations.isNameNonEmpty)) {
        errorList.push({
          primaryText: intl.formatMessage({ id: 'app.function.add.error.missingName' }),
          icon: 'IncidentTriangle',
        });
      }

      addNameValidationResult.forEach(({ name, validations }) => {
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

    if (addNameErrorCode === 'FailedToAdd') {
      return addNameValidationResult.map(({ name, runtimeError }) => ({
        primaryText: `${name}: ${runtimeError}`,
        icon: 'IncidentTriangle',
      }));
    }

    return [];
  }, [addNameValidationResult, addNameErrorCode]);

  const addNameErrorDetails = (
    <>
      <p>{intl.formatMessage({ id: `app.function.add.error.${addNameErrorCode}` })}</p>
      {
        addNameValidationResult.length
          ? (
            <ModalDetailsWrapper>
              <HeroList items={addNameErrorDetailsList} />
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

  const addNameSuccessMessage = (
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
      isOpen={isAddNameModalOpen}
      modalId="addNameModal"
      isClosable={false}
      theme={isAddNameSuccess ? 'success' : addNameErrorCode ? 'failure' : undefined}
    >
      {
        isAddNameSuccess
          ? addNameSuccessMessage
          : addNameErrorCode
            ? addNameErrorDetails
            : addNameHowTo
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
   * Vaidate Name ranges
   */
  const editNamedRanges = async (): Promise<void> => {
    // TODO: write something
  };

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
          onClick={editNamedRanges}
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
      {errorDialog}
      {validateNamedRangesDialog}
    </div>
  );
};

export default Home;
