import React, { FC } from 'react';
import styled from 'styled-components';

export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}

const HeaderSection = styled.section`
  padding: 1.25rem;
  padding: 2.5rem 1.25rem;
  display: flex;
  flex-direction: column;
  align-items: center;
  background-color: ${(props) => props.theme.color.background};
`;

const AppLogo = styled.img`
  width: 5.625rem;
  height: auto;
`;

const HeaderTitle = styled.h1`
  font-size: 2.625rem;
  font-weight: 100;
  margin-bottom: 0;
`;

const Header: FC<HeaderProps> = ({ title, logo, message }) => (
  <HeaderSection>
    <AppLogo src={logo} alt={title} title={title} />
    <HeaderTitle>{message}</HeaderTitle>
  </HeaderSection>
);

export default Header;
