import { createGlobalStyle } from 'styled-components';

const GlobalStyle = createGlobalStyle`
  html, body {
    color: ${(props) => props.theme.color.body};
  }

  #container {
    width: 100%;
    height: 100%;
  }
`;

export default GlobalStyle;
