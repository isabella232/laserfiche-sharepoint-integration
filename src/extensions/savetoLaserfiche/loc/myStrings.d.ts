declare interface ISavetoLaserficheCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'SavetoLaserficheCommandSetStrings' {
  const strings: ISavetoLaserficheCommandSetStrings;
  export = strings;
}
