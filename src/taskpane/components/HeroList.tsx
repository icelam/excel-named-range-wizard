import React, { FC } from 'react';
import styled from 'styled-components';
import { Icon as FrabicIcon } from 'office-ui-fabric-react';

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
    padding-bottom: 1.25rem;
    display: -webkit-flex;
    display: flex;

    span {
      font-size: 0.75rem;
      vertical-align: middle;
    }
  }
`;

const Icon = styled<{ colorVarient: string }>(FrabicIcon)`
  vertical-align: middle;
  margin-right: 0.625rem;
`;

const HeroList: FC<HeroListProps> = ({ items }) => {
  const listItems = items.map(({ primaryText, icon }, index) => (
    // eslint-disable-next-line react/no-array-index-key
    <li key={`${primaryText}_${index}`}>
      <Icon iconName={icon} /><span>{primaryText}</span>
    </li>
  ));

  return <UnorderList>{listItems}</UnorderList>;
};

export default HeroList;
