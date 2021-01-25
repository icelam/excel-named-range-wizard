import React, { FC } from 'react';
import styled from 'styled-components';

export interface HeroListItem {
  icon: string;
  primaryText: string;
}

export interface HeroListProps {
  message: string;
  items: HeroListItem[];
}

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
  color: #333333;
`;

const UnorderList = styled.ul`
  margin: 0;
  padding: 0;
  list-style-type: none;
  margin-top: 1.25rem;

  li {
    padding-bottom: 1.25rem;;
    display: -webkit-flex;
    display: flex;

    i.ms-Icon {
      margin-right: 0.625rem;
    }

    span {
      font-size: 0.75rem;
    }
  }
`;

const HeroList: FC<HeroListProps> = ({ children, items, message }) => {
  const listItems = items.map((item) => (
    <li key={item.primaryText}>
      <i className={`ms-Icon ms-Icon--${item.icon}`} />
      <span>{item.primaryText}</span>
    </li>
  ));

  return (
    <MainWrapper>
      <Title>{message}</Title>
      <UnorderList>{listItems}</UnorderList>
      {children}
    </MainWrapper>
  );
};

export default HeroList;
