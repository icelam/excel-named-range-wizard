import mediaQueries from './media';

const colors = {
  white: '#ffffff',
  black: '#000000',
  lightGrey: '#f4f4f4',
  darkGrey: '#333333',
  red: '#f0533d',
  green: '#4caf50',
};

const themes = {
  light: {
    color: {
      background: colors.lightGrey,
      body: colors.darkGrey,
      success: colors.green,
      failure: colors.red,
    },
    mediaQueries: {
      ...mediaQueries,
    },
  },
};

export default themes;
