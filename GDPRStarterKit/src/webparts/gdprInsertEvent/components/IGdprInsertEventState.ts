import { ISPTermObject } from '../../../components/SPTermStoreService';

export interface IGdprInsertEventState {
  currentEventType: string;
  isValid: boolean;
  showDialogResult: boolean;
  
  title?: string;
  notifiedBy?: string;
  eventAssignedTo?: string;
  eventStartDate?: Date;
  eventEndDate?: Date;
  postEventReport?: string;
  additionalNotes?: string;
  breachType?: ISPTermObject;
  riskType?: ISPTermObject[];
  severity?: ISPTermObject;
  dpaNotified?: boolean;
  dpaNotificationDate?: Date;
  estimatedAffectedSubjects?: Number;
  toBeDetermined?: boolean;
  includesChildren?: boolean;
  includesChildrenInProgress?: boolean;
  actionPlan?: string;
  breachResolved?: boolean;
  actionsTaken?: string;
  includesSensitiveData?: ISPTermObject;
  dataSubjectIsChild?: boolean;
  indirectDataProvider?: boolean;
  dataProvider?: string;
  consentNotes?: string;
  consentType?: ISPTermObject[];
  consentIsInternal?: boolean;
  consentWithdrawalType?: ISPTermObject[];
  consentWithdrawalNotes?: string;
  originalConsentAvailable?: boolean;
  originalConsent?: number;
  notifyApplicable?: boolean;
  processingType?: ISPTermObject[];
  processors?: string[];
  archivedData?: string;
  anonymize?: boolean;
  archivingNotes?: string;
}