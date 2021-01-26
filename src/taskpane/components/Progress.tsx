import React, { FC } from 'react';
import styled from 'styled-components';
import { Spinner, SpinnerType } from 'office-ui-fabric-react';

export interface ProgressProps {
  message: string;
  className?: string;
}

const Wrapper = styled.div`
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  width: 100%;
  height: 100%;
`;

const Progress: FC<ProgressProps> = ({ message, className }) => (
  <Wrapper className={className}>
    <Spinner type={SpinnerType.large} label={message} />
  </Wrapper>
);

export default Progress;
