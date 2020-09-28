declare interface ITestCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'TestCommandSetStrings' {
  const strings: ITestCommandSetStrings;
  export = strings;
}
