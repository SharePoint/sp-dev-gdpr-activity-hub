import { ISPTermObject } from '../../../components/SPTermStoreService';

export interface IGdprInsertRequestState {
  currentRequestType: string;
  isValid: boolean;
  showDialogResult: boolean;

  title?: string;
  dataSubject?: string;
  dataSubjectEmail?: string;
  verifiedDataSubject?: boolean;
  requestAssignedTo?: string;
  requestInsertionDate?: Date;
  requestDueDate?: Date;
  additionalNotes?: string;
  deliveryMethod?: ISPTermObject;
  correctionDefinition?: string;
  deliveryFormat?: ISPTermObject;
  personalData?: string;
  processingType?: ISPTermObject[];
  notifyApplicable?: boolean;
  reason?: string;
}