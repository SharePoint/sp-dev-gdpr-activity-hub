export interface ITaxonomyTerm {
    Label?: string | number;
    TermGuid: string;
    WssId?: number;
}

export interface IItem {
  kind: string;
  title: string;
}

export interface IGDPRContact extends IItem {
  role: ITaxonomyTerm;
  user: string;
}

export interface IGDPREvent extends IItem {
    notifiedBy: string;
    eventAssignedTo: string;
    eventStartDate: any;
    eventEndDate: any;
    postReport: string;
    additionalNotes: string;
}

export interface IGDPRRequest extends IItem {
    dataSubject: string;
    dataSubjectEmail: string;
    verifiedDataSubject: boolean;
    requestAssignedTo: string;
    requestInsertionDate: any;
    requestDueDate: any;
    additionalNotes: string;
}

export interface IGDPRIncident extends IGDPREvent {
    severity: ITaxonomyTerm;
}

export interface IIncidentDataBreach extends IGDPRIncident {
    kind: "DataBreach";
    breachType: ITaxonomyTerm;
    dpaNotified: boolean;
    dpaNotificationDate: any;
    estimatedNumberOfAffectedDataSubjects: number;
    toBeDetermined: boolean;
    includesChildrenData: boolean;
    inProgress: boolean;
    actionPlan: string;
    breachResolved: boolean;
    actionsTaken: string;
}

export interface IIncidentIdentityRisk extends IGDPRIncident {
    kind: "IdentityRisk";
    riskType: Array<ITaxonomyTerm>;
}

export interface IEventDataArchived extends IGDPREvent {
    kind: "DataArchived";
    archivedData: string;
    includesSensitiveData: ITaxonomyTerm;
    includesChildrenData: boolean;
    anonymize: boolean;
    archivingNotes: string;
}

export interface IEventDataConsent extends IGDPREvent {
    kind: "DataConsent";
    consentIsInternal: boolean;
    includesSensitiveData: ITaxonomyTerm;
    dataSubjectIsChild: boolean;
    indirectDataProvider: boolean;
    dataProvider: string;
    consentNotes: string;
    consentType: Array<ITaxonomyTerm>;
}

export interface IEventDataConsentWithdrawal extends IGDPREvent {
    kind: "DataConsentWithdrawal";
    withdrawalType: Array<ITaxonomyTerm>;
    withdrawalNotes: string;
    originalConsentAvailable: boolean;
    originalConsentId: string;
    notifyThirdParties: boolean;
}

export interface IEventDataProcessing extends IGDPREvent {
    kind: "DataProcessing";
    processingType: Array<ITaxonomyTerm>;
    processors: Array<string>;
}

export interface IRequestAccessPersonalData extends IGDPRRequest {
    kind: "Access";
    deliveryMethod: ITaxonomyTerm;
}

export interface IRequestCorrectPersonalData extends IGDPRRequest {
    kind: "Correct";
    correctionDefinition: string;
}

export interface IRequestErasePersonalData extends IGDPRRequest {
    kind: "Erase";
    notifyThirdParties: boolean;
    reason: string;
}

export interface IRequestExportPersonalData extends IGDPRRequest {
    kind: "Export";
    deliveryMethod: ITaxonomyTerm;
    deliveryFormat: ITaxonomyTerm;
}

export interface IRequestObjectionToProcessing extends IGDPRRequest {
    kind: "Objection";
    personalData: string;
    processingType: Array<ITaxonomyTerm>;
    reason: string;
}
