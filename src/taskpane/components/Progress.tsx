import React, { FC } from 'react';
import styled from 'styled-components';
import { Spinner, SpinnerType } from 'office-ui-fabric-react';

export interface ProgressProps {
  message: string;
}

const Wrapper = styled.div`
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  width: 100%;
  height: 100%;
`;

const Progress: FC<ProgressProps> = ({ message }) => (
  <Wrapper>
    <Spinner type={SpinnerType.large} label={message} />
  </Wrapper>
);

export default Progress;
