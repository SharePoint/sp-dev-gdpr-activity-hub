declare interface IGdprInsertEventStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  
  YesText: string;
  NoText: string;

  EventTypeFieldLabel: string;
  EventTypeDataBreachLabel: string;
  EventTypeIdentityRiskLabel: string;
  EventTypeDataConsentLabel: string;
  EventTypeDataConsentWithdrawalLabel: string;
  EventTypeDataProcessingLabel: string;
  EventTypeDataArchivedLabel: string;

  TitleFieldLabel: string;
  TitleFieldPlaceholder: string;
  TitleFieldValidationErrorMessage: string;
  NotifiedByFieldLabel: string;
  NotifiedByFieldPlaceholder: string;
  NotifiedByFieldValidationErrorMessage: string;
  EventAssignedToFieldLabel: string;
  EventAssignedToFieldPlaceholder: string;
  EventStartDateFieldLabel: string;
  EventStartDateFieldPlaceholder: string;
  EventStartTimeHoursFieldLabel: string;
  EventStartTimeMinutesFieldLabel: string;
  EventEndDateFieldLabel: string;
  EventEndDateFieldPlaceholder: string;
  EventEndTimeHoursFieldLabel: string;
  EventEndTimeMinutesFieldLabel: string;
  PostEventReportFieldLabel: string;
  AdditionalNotesFieldLabel: string;
  BreachTypeFieldLabel: string;
  BreachTypeFieldPlaceholder: string;
  SeverityFieldLabel: string;
  SeverityFieldPlaceholder: string;
  DPANotifiedFieldLabel: string;
  DPANotificationDateFieldLabel: string;
  DPANotificationDateFieldPlaceholder: string;
  DPANotificationTimeHoursFieldLabel: string;
  DPANotificationTimeMinutesFieldLabel: string;
  EstimatedAffectedSubjectsFieldLabel: string;
  EstimatedAffectedSubjectsFieldValidationErrorMessage: string;
  ToBeDeterminedFieldLabel: string;
  IncludesChildrenFieldLabel: string;
  IncludesChildrenInProgressFieldLabel: string;
  ActionPlanFieldLabel: string;
  BreachResolvedFieldLabel: string;
  ActionsTakenFieldLabel: string;
  RiskTypeFieldLabel: string;
  RiskTypeFieldPlaceholder: string;
  IncludesSensitiveDataFieldLabel: string;
  IncludesSensitiveDataFieldPlaceholder: string;
  DataSubjectIsChildFieldLabel: string;
  IndirectDataProviderFieldLabel: string;
  DataProviderFieldLabel: string;
  ConsentNotesFieldLabel: string;
  ConsentTypeFieldLabel: string;
  ConsentTypeFieldPlaceholder: string;
  ConsentIsInternalFieldLabel: string;
  InternalConsentText: string;
  ExternalConsentText: string;
  ConsentWithdrawalTypeFieldLabel: string;
  ConsentWithdrawalTypeFieldPlaceholder: string;
  ConsentWithdrawalNotesFieldLabel: string;
  OriginalConsentAvailableFieldLabel: string;
  OriginalConsentFieldLabel: string;
  OriginalConsentFieldPlaceholder: string;
  NotifyApplicableFieldLabel: string;
  ProcessingTypeFieldLabel: string;
  ProcessingTypeFieldPlaceholder: string;
  ProcessorsFieldLabel: string;
  ProcessorsFieldPlaceholder: string;
  ArchivedDataFieldLabel: string;
  AnonymizeFieldLabel: string;
  ArchivingNotesFieldLabel: string; 

  SaveButtonText: string;
  CancelButtonText: string;
  ItemSavedMessage: string;
  InsertNextLabel: string;
  GoHomeLabel: string;
  ItemInsertedDialogTitle: string;
  ItemInsertedDialogSubText: string;

  HoursValidationError: string;
  MinutesValidationError: string;
}

declare module 'gdprInsertEventStrings' {
  const strings: IGdprInsertEventStrings;
  export = strings;
}
