import React, { FC } from 'react';
import styled from 'styled-components';
import { Modal as FabricModal, IconButton, Icon } from 'office-ui-fabric-react';

interface ModalProps {
  title: string;
  onDismiss: () => void;
  isOpen: boolean;
  isBlocking?: boolean;
  modalId: string;
  theme?: 'success' | 'failure';
  isClosable?: boolean;
}

const Header = styled.div<{varient: ModalProps['theme']}>`
  flex: 1 1 auto;
  display: flex;
  align-items: center;
  font-weight: 600;
  font-size: 1.25rem;
  padding: 0.75rem 0.75rem 0.875rem 1.5rem;
  justify-content: space-between;
  ${(props) => props.varient && `border-top: 4px solid ${props.theme.color[props.varient]};`}

  span {
    vertical-align: middle;
  }

  .ms-Icon {
    color: ${(props) => props.theme.color.body};
  }
`;

const StatusIcon = styled<{ colorVarient: string }>(Icon)`
  vertical-align: middle;
  margin-right: 0.5rem;
  ${(props) => props.colorVarient && `color: ${props.theme.color[props.colorVarient]};`}
`;

const Content = styled.div`
  padding: 0 1.5rem 1.5rem 1.5rem;
`;

const getIcon = (varient: string): JSX.Element | null => {
  let icon;
  switch (varient) {
    case 'success':
      icon = 'CompletedSolid';
      break;
    case 'failure':
      icon = 'StatusErrorFull';
      break;
    default:
      icon = '';
  }

  return icon ? <StatusIcon iconName={icon} colorVarient={varient} /> : null;
};

const Modal: FC<ModalProps> = ({
  children, onDismiss, title, isOpen, isBlocking = true, modalId, theme, isClosable = false,
}) => (
  <FabricModal
    titleAriaId={modalId}
    isOpen={isOpen}
    onDismiss={onDismiss}
    isBlocking={isBlocking}
  >
    <Header varient={theme}>
      <span>
        {getIcon(theme)}{title}
      </span>
      {isClosable && (
        <IconButton
          iconProps={{ iconName: 'Cancel' }}
          ariaLabel="Close"
          onClick={onDismiss}
        />
      )}
    </Header>
    <Content>
      {children}
    </Content>
  </FabricModal>
);

export default Modal;
