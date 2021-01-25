import mediaQueries from './media';

const colors = {
  white: '#ffffff',
  black: '#000000',
  lightGrey: '#f4f4f4',
  darkGrey: '#333333',
};

const themes = {
  light: {
    color: {
      background: colors.lightGrey,
      body: colors.darkGrey,
    },
    mediaQueries: {
      ...mediaQueries,
    },
  },
};

export default themes;
