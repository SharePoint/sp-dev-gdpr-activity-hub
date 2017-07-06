declare interface IGdprInsertRequestStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TargetListFieldLabel: string;

  YesText: string;
  NoText: string;

  RequestTypeFieldLabel: string;
  RequestTypeAccessLabel: string;
  RequestTypeCorrectLabel: string;
  RequestTypeExportLabel: string;
  RequestTypeObjectionLabel: string;
  RequestTypeEraseLabel: string;

  TitleFieldLabel: string;
  TitleFieldPlaceholder: string;
  TitleFieldValidationErrorMessage: string;
  DataSubjectFieldLabel: string;
  DataSubjectFieldPlaceholder: string;
  DataSubjectEmailFieldLabel: string;
  DataSubjectEmailFieldPlaceholder: string;
  DataSubjectEmailFieldValidationErrorMessage: string;
  VerifiedDataSubjectFieldLabel: string;
  RequestAssignedToFieldLabel: string;
  RequestAssignedToFieldPlaceholder: string;
  RequestInsertionDateFieldLabel: string;
  RequestInsertionDateFieldPlaceholder: string;
  RequestDueDateFieldLabel: string;
  RequestDueDateFieldPlaceholder: string;
  AdditionalNotesFieldLabel: string;
  DeliveryMethodFieldLabel: string;
  DeliveryMethodFieldPlaceholder: string;
  CorrectionDefinitionFieldLabel: string;
  DeliveryFormatFieldLabel: string;
  DeliveryFormatFieldPlaceholder: string;
  PersonalDataFieldLabel: string;
  ProcessingTypeFieldLabel: string;
  ProcessingTypeFieldPlaceholder: string;
  ReasonFieldLabel: string;
  NotifyApplicableFieldLabel: string;

  SaveButtonText: string;
  CancelButtonText: string;
  ItemSavedMessage: string;
  InsertNextLabel: string;
  GoHomeLabel: string;
  ItemInsertedDialogTitle: string;
  ItemInsertedDialogSubText: string;
}

declare module 'gdprInsertRequestStrings' {
  const strings: IGdprInsertRequestStrings;
  export = strings;
}
