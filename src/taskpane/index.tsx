import 'office-ui-fabric-react/dist/css/fabric.min.css';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import React from 'react';
import * as ReactDOM from 'react-dom';
// eslint-disable-next-line import/no-extraneous-dependencies
import { AppContainer } from 'react-hot-loader';
import App from './App';

initializeIcons();

let isOfficeInitialized = false;

const render = (Component): void => {
  ReactDOM.render(
    <AppContainer>
      <Component isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById('container'),
  );
};

/* Render application after Office initializes */
Office.initialize = (): void => {
  isOfficeInitialized = true;
  render(App);
};

if ((module as any).hot) {
  (module as any).hot.accept('./App', () => {
    // eslint-disable-next-line global-require
    const NextApp = require('./App').default;
    render(NextApp);
  });
}
