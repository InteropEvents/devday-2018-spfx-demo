declare interface ITodoCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'TodoCommandSetStrings' {
  const strings: ITodoCommandSetStrings;
  export = strings;
}
