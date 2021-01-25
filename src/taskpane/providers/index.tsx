import React, { FC } from 'react';
import { IntlProvider } from 'react-intl';
import { ThemeProvider } from 'styled-components';
import translations from '../locale';
import { themes } from '../styles';

const currentLocale = 'en';

const Providers: FC = ({ children }) => (
  <IntlProvider
    locale={currentLocale}
    messages={translations[currentLocale]}
  >
    <ThemeProvider theme={themes.light}>
      {children}
    </ThemeProvider>
  </IntlProvider>
);

export default Providers;
