import globalTheme from '../config/theme/globalTheme.json';

export interface ITheme {
  app: {
    backgroundColor: string;
    fontFamily: string;
  };
  card: {
    background: string;
    borderColor: string;
    borderRadius: string;
    padding: string;
    shadow: string;
  };
  header: {
    background: string;
    textColor: string;
    fontSize: string;
    fontWeight: number;
  };
  section: {
    headerBackground: string;
    headerText: string;
    borderRadius: string;
    padding: string;
  };
  text: {
    primary: string;
    secondary: string;
    labelSize: string;
  };
  input: {
    borderColor: string;
    borderRadius: string;
    padding: string;
    fontSize: string;
    background: string;
  };
  button: {
    primaryBackground: string;
    primaryText: string;
    secondaryBackground: string;
    borderRadius: string;
    padding: string;
    fontWeight: number;
  };
}

export function getTheme(): ITheme {
  return globalTheme as ITheme;
}
