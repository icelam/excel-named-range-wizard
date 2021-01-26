import React, { FC } from 'react';
import styled from 'styled-components';
import { useIntl } from 'react-intl';
import { DefaultButton } from 'office-ui-fabric-react';
import Header from '../components/Header';
import { getNamedRanges } from '../excelUtils';

// images references in the manifest
import '../../../assets/icon-16.png';
import '../../../assets/icon-32.png';
import '../../../assets/icon-80.png';

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

const Home: FC = () => {
  const intl = useIntl();

  const onClick = async (): Promise<void> => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load('address');

        // Update the fill color
        range.format.fill.color = 'yellow';

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
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
          onClick={getNamedRanges}
        >
          {intl.formatMessage({ id: 'app.function.export' })}
        </FullwidthButton>
        <FullwidthButton
          iconProps={{ iconName: 'PageAdd' }}
          onClick={onClick}
        >
          {intl.formatMessage({ id: 'app.function.add' })}
        </FullwidthButton>
        <FullwidthButton
          iconProps={{ iconName: 'PageEdit' }}
          onClick={onClick}
        >
          {intl.formatMessage({ id: 'app.function.edit' })}
        </FullwidthButton>
      </MainWrapper>
    </div>
  );
};

export default Home;
