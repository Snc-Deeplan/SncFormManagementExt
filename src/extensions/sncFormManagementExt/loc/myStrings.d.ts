declare interface ISncFormManagementExtCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'SncFormManagementExtCommandSetStrings' {
  const strings: ISncFormManagementExtCommandSetStrings;
  export = strings;
}
