import React, { FC } from 'react';
import styled from 'styled-components';

export interface HeroListItem {
  icon: string;
  primaryText: string;
}

export interface HeroListProps {
  items: HeroListItem[];
}

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

const HeroList: FC<HeroListProps> = ({ items }) => {
  const listItems = items.map((item) => (
    <li key={item.primaryText}>
      <i className={`ms-Icon ms-Icon--${item.icon}`} />
      <span>{item.primaryText}</span>
    </li>
  ));

  return <UnorderList>{listItems}</UnorderList>;
};

export default HeroList;
