/* Color Schemes */
export const mediaColorScheme = {
  light: '(prefers-color-scheme: light)',
  dark: '(prefers-color-scheme: dark)',
};

/*
 * Breakpoints
 * xs < 576 | sm 576 > < 768 | md 768 > < 992 | lg 992 > < 1200 | xl 1200 > < 1600 | xxl > 1600
 */
export const breakpoints = {
  sm: '576',
  md: '768',
  lg: '992',
  xl: '1200',
  xxl: '1600',
};

/* Media queries */
export const mediaQueries = {
  mobile: `only screen and (max-width: ${breakpoints.sm - 1}px)`,
  tablet: `only screen and (min-width: ${breakpoints.sm}px) and (max-width: ${breakpoints.md - 1}px)`,
  desktop: `only screen and (min-width: ${breakpoints.md}px)`,
  smallScreen: `only screen and (max-width: ${breakpoints.lg - 1}px)`,
  largerScreen: `only screen and (min-width: ${breakpoints.lg}px)`,
  extraLargeScreen: `only screen and (min-width: ${breakpoints.xl}px)`,
  tabletDesktop: `only screen and (min-width: ${breakpoints.sm}px)`,
  mobileTablet: `only screen and (max-width: ${breakpoints.md - 1}px)`,
};
