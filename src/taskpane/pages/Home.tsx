import React, { FC, useMemo } from 'react';
import { useIntl } from 'react-intl';
import { Button, ButtonType } from 'office-ui-fabric-react';
import Header from '../components/Header';
import HeroList, { HeroListItem } from '../components/HeroList';

// images references in the manifest
import '../../../assets/icon-16.png';
import '../../../assets/icon-32.png';
import '../../../assets/icon-80.png';

const Home: FC = () => {
  const intl = useIntl();
  const listItems: HeroListItem[] = useMemo(() => [
    {
      icon: 'Ribbon',
      primaryText: 'Lorem Ipsum',
    },
    {
      icon: 'Unlock',
      primaryText: 'Lorem Ipsum',
    },
    {
      icon: 'Design',
      primaryText: 'Lorem Ipsum',
    },
  ], []);

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
        message={intl.formatMessage({ id: 'app.welcome' })}
      />
      <HeroList message={intl.formatMessage({ id: 'app.home.tagline' })} items={listItems}>
        <p>
          Modify the source files, then click <b>Run</b>.
        </p>
        <Button
          buttonType={ButtonType.hero}
          iconProps={{ iconName: 'ChevronRight' }}
          onClick={onClick}
        >
          Run
        </Button>
      </HeroList>
    </div>
  );
};

export default Home;
