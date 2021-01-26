import React, {
  FC, useState, useMemo, useEffect,
} from 'react';
import styled from 'styled-components';
import { useIntl } from 'react-intl';
import {
  DefaultButton, DialogFooter,
} from 'office-ui-fabric-react';
import {
  Header, Modal, Progress, HeroList,
} from '../components';
import {
  exportNamedRangesToWorksheet,
  validateNamedRanges,
  NamedRangeType,
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

const InvalidNamedRangeListWrapper = styled.div`
  max-height: 200px;
  overflow-y: auto;
  background-color: ${(props) => props.theme.color.background};
  padding: 0 1rem;
`;

const Home: FC = () => {
  const intl = useIntl();
  const [isLoading, setIsLoading] = useState(false);

  /**
   * Add Named Ranges Modal
   */
  const [isAddNameModalOpen, setIsAddNameModalOpen] = useState(false);

  const showAddNamesModal = (): void => {
    setIsAddNameModalOpen(true);
  };

  const hideAddNamesModal = (): void => {
    setIsAddNameModalOpen(false);
  };

  const addNamesModal = (
    <Modal
      onDismiss={hideAddNamesModal}
      title="Add Names"
      isOpen={isAddNameModalOpen}
      modalId="addNameModal"
      isClosable
    >
      To-Do
    </Modal>
  );

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
   * Vaidate Name ranges
   */
  const [isValidateNamedRangesDialogOpen, setIsValidateNamedRangesDialogOpen] = useState(false);
  const [invalidNamedRanges, setInvalidNameRanges] = useState<NamedRangeType[]>([]);

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
    setInvalidNameRanges(names);
    setIsLoading(false);
  };

  const hideValidateNameRangesDialog = (): void => {
    setIsValidateNamedRangesDialogOpen(false);
  };

  const invalidNamedRangesList = useMemo(() => invalidNamedRanges.map(({ name }) => (
    { icon: 'IncidentTriangle', primaryText: name }
  )), [invalidNamedRanges]);

  const validateNamedRangesDialog = (
    <Modal
      onDismiss={hideValidateNameRangesDialog}
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
            <InvalidNamedRangeListWrapper>
              <HeroList items={invalidNamedRangesList} />
            </InvalidNamedRangeListWrapper>
          )
          : null
      }

      <DialogFooter>
        <DefaultButton onClick={hideValidateNameRangesDialog} text={intl.formatMessage({ id: 'app.modal.ok' })} />
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
