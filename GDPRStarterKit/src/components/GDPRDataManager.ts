/// <reference types="es6-promise" />

import { 
    // IItem, 
    // IGDPREvent,
    // IGDPRRequest,
    // IGDPRIncident,
    ITaxonomyTerm, 
    IIncidentDataBreach,
    IIncidentIdentityRisk,
    IEventDataArchived,
    IEventDataConsent,
    IEventDataConsentWithdrawal,
    IEventDataProcessing,
    IRequestAccessPersonalData,
    IRequestCorrectPersonalData,
    IRequestErasePersonalData,
    IRequestExportPersonalData,
    IRequestObjectionToProcessing     
} from '../domain_model/GDPR_domain_model';
import { default as pnp, ItemAddResult, WebEnsureUserResult }  from "sp-pnp-js";

export interface IGDPRDataManager {
    setup(settings: any): void;
    insertNewRequest(request: IRequestAccessPersonalData | IRequestCorrectPersonalData |
        IRequestErasePersonalData | IRequestExportPersonalData |
        IRequestObjectionToProcessing): Promise<number>;
    insertNewEvent(event: IEventDataConsent | IEventDataConsentWithdrawal | 
        IEventDataProcessing | IEventDataArchived |
        IIncidentDataBreach | IIncidentIdentityRisk): Promise<number>;
}

export interface IGDPRDataManagerSharePointSettings {
    requestsListId?: string;
    eventsListId?: string;
}

export class GDPRDataManager implements IGDPRDataManager {

    private requestsListId: string;
    private eventsListId: string;

    public setup(settings: any): void {
        let s = <IGDPRDataManagerSharePointSettings>settings;
        if (s != null)
        {
            if (s.requestsListId != null && s.requestsListId.length > 0) this.requestsListId = s.requestsListId;
            if (s.eventsListId != null && s.eventsListId.length > 0) this.eventsListId = s.eventsListId;
        }
    }

    public insertNewRequest(request: IRequestAccessPersonalData | IRequestCorrectPersonalData |
        IRequestErasePersonalData | IRequestExportPersonalData |
        IRequestObjectionToProcessing): Promise<number> {
        
        return(new Promise<number>((resolve, reject) => {
            this.resolveUserId(request.requestAssignedTo).then((requestAssignedToId: number) => {

                let mappedRequest: any = {
                    Title: request.title,
                    GDPRDataSubject: request.dataSubject,
                    GDPRDataSubjectEmail: request.dataSubjectEmail,
                    GDPRVerifiedDataSubject: request.verifiedDataSubject,
                    GDPRRequestAssignedToId: requestAssignedToId,
                    GDPRRequestInsertionDate: request.requestInsertionDate,
                    GDPRRequestDueDate: request.requestDueDate,
                    GDPRNotes: request.additionalNotes,
                };

                switch (request.kind)
                {
                    case "Access":
                        mappedRequest.ContentTypeId = "0x0100A16621D3EDF4F141847640F9058A5B730100CAD083221153F240A5DA2C3CFDF2C451";
                        mappedRequest.GDPRDeliveryMethod = {
                            __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                            Label: request.deliveryMethod.Label,
                            TermGuid: request.deliveryMethod.TermGuid,
                            WssId: -1
                        };
                        break;
                    case "Correct":
                        mappedRequest.ContentTypeId = "0x0100A16621D3EDF4F141847640F9058A5B7302000457A3587EA71A4BA545274EF3432D89";
                        mappedRequest.GDPRCorrectionDefinition = request.correctionDefinition;
                        break;
                    case "Erase":
                        mappedRequest.ContentTypeId = "0x0100A16621D3EDF4F141847640F9058A5B730300185B78202037C741B91654045C963841";
                        mappedRequest.GDPRNotifyExternalProcessor = request.notifyThirdParties;
                        mappedRequest.GDPRReason = request.reason;
                        break;
                    case "Export":
                        mappedRequest.ContentTypeId = "0x0100A16621D3EDF4F141847640F9058A5B730400CA60A3E25D2AFB429FCCAB901C990383";
                        mappedRequest.GDPRDeliveryMethod = {
                            __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                            Label: request.deliveryMethod.Label,
                            TermGuid: request.deliveryMethod.TermGuid,
                            WssId: -1
                        };
                        mappedRequest.GDPRDeliveryFormat = {
                            __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                            Label: request.deliveryFormat.Label,
                            TermGuid: request.deliveryFormat.TermGuid,
                            WssId: -1
                        };
                        break;
                    case "Objection":
                        mappedRequest.ContentTypeId = "0x0100A16621D3EDF4F141847640F9058A5B730500CCB415E01CEA9C41B6C9E7979D9EF06C";
                        mappedRequest.GDPRPersonalData = request.personalData;
                        mappedRequest['f9daa214b1dd4b0dafb30c8f454778ec'] = this.prepareTaxonomyMultivalues(request.processingType);
                        mappedRequest.GDPRReason = request.reason;
                        break;
                }

                pnp.sp.web.lists.getById(this.requestsListId).items.add(mappedRequest).then((iar: ItemAddResult) => {
                    resolve(iar.data.Id);
                }).catch((ex: any) => {
                    reject(ex);
                });
            });
        }));
    }

    public insertNewEvent(event: IEventDataConsent | IEventDataConsentWithdrawal | 
        IEventDataProcessing | IEventDataArchived |
        IIncidentDataBreach | IIncidentIdentityRisk): Promise<number> {

        return(new Promise<number>((resolve, reject) => {

            let dataProcessors: number[] = [];

            if (event.kind == "DataProcessing")
            {
                let processorsResolvers = event.processors.map(p => {
                    return(this.resolveUserId(p));
                });
                Promise.all(processorsResolvers).then(processors => {
                    dataProcessors = processors;
                });
            }

            this.resolveUserId(event.eventAssignedTo).then((eventAssignedToId: number) => {

                let mappedEvent: any = {
                    Title: event.title,
                    GDPRNotifiedBy: event.notifiedBy,
                    GDPRReportAssignedToId: eventAssignedToId,
                    GDPRReportStartDateTime: event.eventStartDate,
                    GDPRReportEndDateTime: event.eventEndDate,
                    GDPRPostEventReport: event.postReport,
                    GDPRNotes: event.additionalNotes,
                };

                switch (event.kind)
                {
                    case "DataConsent":
                        mappedEvent.ContentTypeId = "0x0100B506463210A9D340A2E0E0A0889DC892040086569F519007EB40AE2411EDAB6A61E8";
                        mappedEvent.GDPRConsentIsInternal = event.consentIsInternal;
                        if (event.includesSensitiveData)
                        {
                            mappedEvent.GDPRIncludesSensitiveData = {
                                __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                                Label: event.includesSensitiveData.Label,
                                TermGuid: event.includesSensitiveData.TermGuid,
                                WssId: -1
                            };
                        }
                        mappedEvent.GDPRDataSubjectIsChild = event.dataSubjectIsChild;
                        mappedEvent.GDPRIndirectDataProvider = event.indirectDataProvider;
                        mappedEvent.GDPRDataProvider = event.dataProvider;
                        mappedEvent.GDPRConsentNotes = event.consentNotes;
                        mappedEvent['a0dce788628b4762be4c0fae9bee3096'] = this.prepareTaxonomyMultivalues(event.consentType);
                        break;
                    case "DataConsentWithdrawal":
                        mappedEvent.ContentTypeId = "0x0100B506463210A9D340A2E0E0A0889DC8920500EBBB731827924A4E8CEF1F4569BFCA2D";
                        // mappedEvent.GDPRWithdrawalType = event.withdrawalType; // TODO: Replace with a direct call like CSOM
                        mappedEvent['b900b350bfe4474e923938eb73598802'] = this.prepareTaxonomyMultivalues(event.withdrawalType);
                        mappedEvent.GDPRWithdrawalNotes = event.withdrawalNotes;
                        mappedEvent.GDPRConsentLookupAvailable = event.originalConsentAvailable;
                        mappedEvent.GDPRConsentLookupId = event.originalConsentId;
                        mappedEvent.GDPRNotifyExternalProcessor = event.notifyThirdParties;                
                        break;
                    case "DataProcessing":
                        mappedEvent.ContentTypeId = "0x0100B506463210A9D340A2E0E0A0889DC8920300BA08C9FB525ED848AC38713D45981877";
                        mappedEvent['f9daa214b1dd4b0dafb30c8f454778ec'] = this.prepareTaxonomyMultivalues(event.processingType);
                        mappedEvent.GDPRProcessorsId = {
                            results: dataProcessors,
                        };
                        break;
                    case "DataArchived":
                        mappedEvent.ContentTypeId = "0x0100B506463210A9D340A2E0E0A0889DC89207001AD2CE357E64F546B5BEE2786EB34571";
                        if (event.includesSensitiveData)
                        {
                            mappedEvent.GDPRIncludesSensitiveData = {
                                __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                                Label: event.includesSensitiveData.Label,
                                TermGuid: event.includesSensitiveData.TermGuid,
                                WssId: -1
                            };
                        }
                        mappedEvent.GDPRArchivedData = event.archivedData;
                        mappedEvent.GDPRDataDeIdentified = event.anonymize;
                        mappedEvent.GDPRArchivingNotes = event.archivingNotes;
                        mappedEvent.GDPRIncludesChildrenData = event.includesChildrenData;
                        break;
                    case "DataBreach":
                        mappedEvent.ContentTypeId = "0x0100B506463210A9D340A2E0E0A0889DC8920100CEADB5787C515947893D97E56E3E2CEC";
                        mappedEvent.GDPRBreachType = {
                            __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                            Label: event.breachType.Label,
                            TermGuid: event.breachType.TermGuid,
                            WssId: -1
                        };
                        mappedEvent.GDPRSeverity = {
                            __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                            Label: event.severity.Label,
                            TermGuid: event.severity.TermGuid,
                            WssId: -1
                        };
                        mappedEvent.GDPRDPANotified = event.dpaNotified;
                        mappedEvent.GDPRDPANotificationDate = event.dpaNotificationDate;
                        mappedEvent.GDPREstimatedAffectedDataSubject = event.estimatedNumberOfAffectedDataSubjects; // TODO: Fix during provisioning
                        mappedEvent.GDPRIncidentToBeDetermined = event.toBeDetermined;
                        mappedEvent.GDPRIncludesChildrenData = event.includesChildrenData;
                        mappedEvent.GDPRIncidentInProgress = event.inProgress;
                        mappedEvent.GDPRPlannedRecoveryAction = event.actionPlan;
                        mappedEvent.GDPRBreachSolved = event.breachResolved;
                        mappedEvent.GDPRActionsTaken = event.actionsTaken;
                        break;
                    case "IdentityRisk":
                        mappedEvent.ContentTypeId = "0x0100B506463210A9D340A2E0E0A0889DC8920200A6F17E1F42E84E409416BF705FB5DC41";
                        mappedEvent['m23b2cf910684f13aff8b226db423e20'] = this.prepareTaxonomyMultivalues(event.riskType);
                        mappedEvent.GDPRSeverity = {
                            __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                            Label: event.severity.Label,
                            TermGuid: event.severity.TermGuid,
                            WssId: -1
                        };
                        break;
                }

                pnp.sp.web.lists.getById(this.eventsListId).items.add(mappedEvent).then((iar: ItemAddResult) => {
                    resolve(iar.data.Id);
                }).catch((ex: any) => {
                    reject(ex);
                });
            });
        }));
    }

    private resolveUserId(username: string) : Promise<number> {
        return(new Promise<number>((resolve, reject) => {
            if (username != undefined)
            {
                pnp.sp.web.ensureUser("i:0#.f|membership|" + username)
                    .then((result: WebEnsureUserResult) => {
                        resolve(result.data.Id);
                    })
                    .catch((e: any) => {
                        reject(e);                
                    });
            }
            else
            {
                resolve(0);
            }
        }));
    }

    private prepareTaxonomyMultivalues(terms: ITaxonomyTerm[]) : string {

        let termsValuesString: string = "";
        terms.forEach(t => {
            termsValuesString += "-1;#" + t.Label + "|" + t.TermGuid + ";#";
        });
        termsValuesString = termsValuesString.substring(0, termsValuesString.length - 1);

        return(termsValuesString);
    }
}