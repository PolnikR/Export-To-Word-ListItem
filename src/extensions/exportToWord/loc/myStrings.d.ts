declare interface IExportToWordCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ExportToWordCommandSetStrings' {
  const strings: IExportToWordCommandSetStrings;
  export = strings;
}
