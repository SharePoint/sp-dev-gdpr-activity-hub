declare interface IGdprDashboardStrings {
  BasicGroupName: string;
  PropertyPaneDescription: string;

  TaskCompletedColumnTitle: string;
  TaskAssigneeColumnTitle: string;
  TaskTitleColumnTitle: string;
  TaskDueDateColumnTitle: string;
}

declare module 'gdprDashboardStrings' {
  const strings: IGdprDashboardStrings;
  export = strings;
}
