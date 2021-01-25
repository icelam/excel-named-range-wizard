import React, { FC } from 'react';
import Providers from './providers';
import { GlobalStyle } from './styles';
import Progress from './components/Progress';
import Home from './pages/Home';

export interface AppProps {
  isOfficeInitialized: boolean;
}

const App: FC<AppProps> = ({ isOfficeInitialized }) => (
  <Providers>
    <GlobalStyle />
    {
      !isOfficeInitialized
        ? <Progress message="Please sideload your addin to see app body." />
        : <Home />
    }
  </Providers>
);

export default App;
