import * as React from 'react';
import styles from './GdprInsertRequest.module.scss';
import { IGdprInsertRequestProps } from './IGdprInsertRequestProps';

import * as strings from 'gdprInsertRequestStrings';

import pnp from "sp-pnp-js";

import { SPPeoplePicker } from '../../../components/SPPeoplePicker';
import { SPTaxonomyPicker } from '../../../components/SPTaxonomyPicker';
import { ISPTermObject } from '../../../components/SPTermStoreService';
import { SPDateTimePicker } from '../../../components/SPDateTimePicker';

import { GDPRDataManager } from '../../../components/GDPRDataManager';

/**
 * Common Infrastructure
 */
import {
  BaseComponent,
  assign,
  autobind
} from 'office-ui-fabric-react/lib/Utilities';

/**
 * Dialog
 */
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';

/**
 * Choice Group
 */
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';

/**
 * Text Field
 */
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

/**
 * Toggle
 */
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';

/**
 * Button
 */
import { PrimaryButton, DefaultButton, Button, IButtonProps } from 'office-ui-fabric-react/lib/Button';

import { IGdprInsertRequestState } from './IGdprInsertRequestState';

export default class GdprInsertRequest extends React.Component<IGdprInsertRequestProps, IGdprInsertRequestState> {

  /**
   * Main constructor for the component
   */
  constructor() {
    super();
    
    this.state = {
      currentRequestType : "Access",
      isValid: false,
      showDialogResult: false,
    };
  }

  public render(): React.ReactElement<IGdprInsertRequestProps> {
    return (
      <div className={styles.gdprRequest}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
              <ChoiceGroup
                label={ strings.RequestTypeFieldLabel }
                onChange={ this._onChangedRequestType }
                options={ [
                  {
                    key: 'Access',
                    iconProps: { iconName: 'QuickNote' },
                    text: strings.RequestTypeAccessLabel,
                    checked: true,                    
                  },
                  {
                    key: 'Correct',
                    iconProps: { iconName: 'EditNote' },
                    text: strings.RequestTypeCorrectLabel,
                  },
                  {
                    key: 'Export',
                    iconProps: { iconName: 'NoteForward' },
                    text: strings.RequestTypeExportLabel,
                  },
                  {
                    key: 'Objection',
                    iconProps: { iconName: 'NoteReply' },
                    text: strings.RequestTypeObjectionLabel,
                  },
                  {
                    key: 'Erase',
                    iconProps: { iconName: 'EraseTool' },
                    text: strings.RequestTypeEraseLabel,
                  }
                ]}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField 
                label={ strings.TitleFieldLabel } 
                placeholder={ strings.TitleFieldPlaceholder } 
                required={ true } 
                value={ this.state.title }
                onChanged={ this._onChangedTitle }
                onGetErrorMessage={ this._getErrorMessageTitle }
                />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField 
                label={ strings.DataSubjectFieldLabel } 
                placeholder={ strings.DataSubjectFieldPlaceholder } 
                required={ true }
                value={ this.state.dataSubject }
                onChanged={ this._onChangedDataSubject }
                />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField 
                label={ strings.DataSubjectEmailFieldLabel } 
                placeholder={ strings.DataSubjectEmailFieldPlaceholder } 
                required={ false } 
                value={ this.state.dataSubjectEmail }
                onChanged={ this._onChangedDataSubjectEmail }
                onGetErrorMessage={ this._getErrorMessageDataSubjectEmail }
                />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <Toggle
                defaultChecked={ false }
                label={ strings.VerifiedDataSubjectFieldLabel }
                onText={ strings.YesText }
                offText={ strings.NoText }
                checked={ this.state.verifiedDataSubject }
                onChanged={ this._onChangedVerifiedDataSubject }
                />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <SPPeoplePicker
                context={ this.props.context }
                label={ strings.RequestAssignedToFieldLabel }
                required={ true }
                placeholder={ strings.RequestAssignedToFieldPlaceholder }
                onChanged={ this._onChangedRequestAssignedTo }
                />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <SPDateTimePicker 
                showTime={ false }
                includeSeconds={ false }
                dateLabel={ strings.RequestInsertionDateFieldLabel }
                datePlaceholder={ strings.RequestInsertionDateFieldPlaceholder }
                isRequired={ true }
                initialDateTime={ this.state.requestInsertionDate }
                onChanged={ this._onChangedRequestInsertionDate }
                />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <SPDateTimePicker 
                showTime={ false }
                includeSeconds={ false }
                dateLabel={ strings.RequestDueDateFieldLabel }
                datePlaceholder={ strings.RequestDueDateFieldPlaceholder } 
                isRequired={ true } 
                initialDateTime={ this.state.requestDueDate }
                onChanged={ this._onChangedRequestDueDate }
                />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField
                label={ strings.AdditionalNotesFieldLabel }
                multiline 
                autoAdjustHeight
                value={ this.state.additionalNotes }
                onChanged={ this._onChangedAdditionalNotes }
                />
            </div>
          </div>
          {
            (this.state.currentRequestType === "Access" || this.state.currentRequestType === "Export") ?
            <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
              <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                <SPTaxonomyPicker 
                  context={ this.props.context }
                  termSetName="Delivery Methods"
                  label={ strings.DeliveryMethodFieldLabel }
                  placeholder={ strings.DeliveryMethodFieldPlaceholder }
                  required={ true }
                  onChanged={ this._onChangedDeliveryMethod }
                  />
              </div>
            </div>
            : null
          }
          {
            (this.state.currentRequestType === "Correct") ?
            <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
              <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                <TextField
                  label={ strings.CorrectionDefinitionFieldLabel }
                  multiline 
                  autoAdjustHeight
                  required={ true }
                  value={ this.state.correctionDefinition }
                  onChanged={ this._onChangedCorrectionDefinition }
                  />
              </div>
            </div>
            : null
          }
          {
            (this.state.currentRequestType === "Export") ?
            <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
              <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                <SPTaxonomyPicker 
                  context={ this.props.context }
                  termSetName="Delivery Format"
                  label={ strings.DeliveryFormatFieldLabel }
                  placeholder={ strings.DeliveryFormatFieldPlaceholder }
                  required={ true } 
                  onChanged={ this._onChangedDeliveryFormat }
                  />
              </div>
            </div>
            : null
          }
          {
            (this.state.currentRequestType === "Objection") ?
            <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
              <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                <TextField
                  label={ strings.PersonalDataFieldLabel }
                  multiline 
                  autoAdjustHeight
                  value={ this.state.personalData }
                  onChanged={ this._onChangedPersonalData }
                  />
              </div>
            </div>
            : null
          }
          {
            (this.state.currentRequestType === "Objection") ?
            <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
              <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                <SPTaxonomyPicker 
                  context={ this.props.context }
                  termSetName="Processing Type"
                  label={ strings.ProcessingTypeFieldLabel }
                  placeholder={ strings.ProcessingTypeFieldPlaceholder }
                  required={ true }
                  onChanged={ this._onChangedProcessingType }
                  />
              </div>
            </div>
            : null
          }
          {
            (this.state.currentRequestType === "Erase") ?
            <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
              <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                <Toggle
                  defaultChecked={ false }
                  label={ strings.NotifyApplicableFieldLabel }
                  onText={ strings.YesText }
                  offText={ strings.NoText }
                  checked={ this.state.notifyApplicable }
                  onChanged={ this._onChangedNotifyApplicable }
                  />
              </div>
            </div>
            : null
          }
          {
            (this.state.currentRequestType === "Objection" || this.state.currentRequestType === "Erase") ?
            <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
              <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                <TextField
                  label={ strings.ReasonFieldLabel }
                  multiline 
                  autoAdjustHeight
                  value={ this.state.reason }
                  onChanged={ this._onChangedReason }
                  />
              </div>
            </div>
            : null
          }
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <PrimaryButton
                data-automation-id='saveRequest'
                label={ strings.SaveButtonText  }
                disabled={ !this.state.isValid }
                onClick={ this._saveClick }
                />
                &nbsp;&nbsp;
              <Button
                data-automation-id='cancel'
                label={ strings.CancelButtonText  }
                onClick={ this._cancelClick }
                />
            </div>
          </div>
        </div>
        <Dialog
            isOpen={ this.state.showDialogResult }
            type={ DialogType.normal }
            onDismiss={ this._closeInsertDialogResult }
            title={ strings.ItemInsertedDialogTitle }
            subText={ strings.ItemInsertedDialogSubText }
            isBlocking={ true }
          >
          <DialogFooter>
            <PrimaryButton
              onClick={ this._insertNextClick } 
              label={ strings.InsertNextLabel } />
            <DefaultButton 
              onClick={ this._goHomeClick }
              label={ strings.GoHomeLabel } />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  @autobind
  private _onChangedRequestType(ev: React.FormEvent<HTMLInputElement>, option: any) {
    this.state.currentRequestType = option.key;
    this.state.deliveryMethod = null;
    this.state.deliveryFormat = null;
    this.state.processingType = [];

    this._updateState(this.state);
  }

  @autobind
  private _getErrorMessageTitle(value: string): string {

    return (value == null || value.length == 0 || value.length >= 10)
      ? ''
      : `${strings.TitleFieldValidationErrorMessage} ${value.length}.`;
  }

  private _getErrorMessageDataSubjectEmail(value: string): string {

    let emailRegEx: RegExp = new RegExp(/^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/);

    if (value != null && value.length > 0 && !emailRegEx.test(value))
    {
      return(strings.DataSubjectEmailFieldValidationErrorMessage);
    }
    else
    {
      return("");
    }
  }

  @autobind
  private _updateState(state: IGdprInsertRequestState): void {
    state.isValid = this._formIsValid();
    this.setState(state);
  }

  @autobind
  private _onChangedTitle(newValue: string): void {
    this.state.title = newValue;
    this._updateState(this.state);
  }

  @autobind
  private _onChangedDataSubject(newValue: string): void {
    this.state.dataSubject = newValue;
    this._updateState(this.state);
  }

  @autobind
  private _onChangedDataSubjectEmail(newValue: string): void {
    this.state.dataSubjectEmail = newValue;
    this._updateState(this.state);
  }

  @autobind
  private _onChangedVerifiedDataSubject(checked: boolean): void {
    this.state.verifiedDataSubject = checked;
    this._updateState(this.state);
  }  

  @autobind
  private _onChangedRequestAssignedTo(items: string[]): void {
    this.state.requestAssignedTo = items[0];
    this._updateState(this.state);
  }

  @autobind
  private _onChangedRequestInsertionDate(newValue: Date): void {
    this.state.requestInsertionDate = newValue;
    this._updateState(this.state);
  }

  @autobind
  private _onChangedRequestDueDate(newValue: Date): void {
    this.state.requestDueDate = newValue;
    this._updateState(this.state);
  }

  @autobind
  private _onChangedAdditionalNotes(newValue: string): void {
    this.state.additionalNotes = newValue;
    this._updateState(this.state);
  }

  @autobind
  private _onChangedDeliveryMethod(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0)
    {
      this.state.deliveryMethod = terms[0];
    }
    else
    {
      this.state.deliveryMethod = null;
    }
    this._updateState(this.state);
  }

  @autobind
  private _onChangedCorrectionDefinition(newValue: string): void {
    this.state.correctionDefinition = newValue;
    this._updateState(this.state);
  }

  @autobind
  private _onChangedDeliveryFormat(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0)
    {
      this.state.deliveryFormat = terms[0];
    }
    else
    {
      this.state.deliveryFormat = null;
    }
    this._updateState(this.state);
  }

  @autobind
  private _onChangedPersonalData(newValue: string): void {
    this.state.personalData = newValue;
    this._updateState(this.state);
  }

  @autobind
  private _onChangedProcessingType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0)
    {
      this.state.processingType = terms;
    }
    else
    {
      this.state.processingType = [];
    }
    this._updateState(this.state);
  }

  @autobind
  private _onChangedNotifyApplicable(checked: boolean): void {
    this.state.notifyApplicable = checked;
    this._updateState(this.state);
  }

  @autobind
  private _onChangedReason(newValue: string): void {
    this.state.reason = newValue;
    this._updateState(this.state);
  }

  @autobind
  private _saveClick(event) {
    event.preventDefault();
    if (this._formIsValid())
    {
      let dataManager = new GDPRDataManager();
      dataManager.setup({
        requestsListId: this.props.targetList,
      });

      let request : any = {
          kind: this.state.currentRequestType,
          title: this.state.title,
          dataSubject: this.state.dataSubject,
          dataSubjectEmail: this.state.dataSubjectEmail,
          verifiedDataSubject: this.state.verifiedDataSubject,
          requestAssignedTo: this.state.requestAssignedTo,
          requestInsertionDate: this.state.requestInsertionDate,
          requestDueDate: this.state.requestDueDate,
          additionalNotes: this.state.additionalNotes,
        };

      switch (request.kind)
      {
        case "Access":
          request.deliveryMethod = {
            Label: this.state.deliveryMethod.name,
            TermGuid: this.state.deliveryMethod.guid,
            WssId: -1,
          };
          break;
        case "Correct":
          request.correctionDefinition = this.state.correctionDefinition;
          break;
        case "Erase":
          request.notifyThirdParties = this.state.notifyApplicable;
          request.reason = this.state.reason;
          break;
        case "Export":
          request.deliveryMethod = {
            Label: this.state.deliveryMethod.name,
            TermGuid: this.state.deliveryMethod.guid,
            WssId: -1,
          };
          request.deliveryFormat = {
            Label: this.state.deliveryFormat.name,
            TermGuid: this.state.deliveryFormat.guid,
            WssId: -1,
          };
          break;
        case "Objection":
          request.personalData = this.state.personalData;
          request.processingType = this.state.processingType.map(i => {
            return {
              Label: i.name,
              TermGuid: i.guid,
              WssId: -1,
            };
          });
          request.reason = this.state.reason;
          break;
      }

      dataManager.insertNewRequest(request).then((itemId: number) => {
        this.state.showDialogResult = true;
        this._updateState(this.state);
      });
    }
  }

  @autobind
  private _cancelClick(event) {
    event.preventDefault();
    window.history.back();
  }

  private _formIsValid() : boolean {
    let isValid: boolean = 
      (this.state.title != null && this.state.title.length > 0) &&
      (this.state.dataSubject != null && this.state.dataSubject.length > 0) &&
      (this.state.requestAssignedTo != null && this.state.requestAssignedTo.length > 0) &&
      (this.state.requestInsertionDate != null) &&
      (this.state.requestDueDate != null);

    if (this.state.currentRequestType == "Access" || this.state.currentRequestType == "Export") {
      isValid = isValid && this.state.deliveryMethod != null;
    }
    if (this.state.currentRequestType == "Export") {
      isValid = isValid && this.state.deliveryFormat != null;
    } 
    if (this.state.currentRequestType == "Correct") {
      isValid = isValid && this.state.correctionDefinition != null && this.state.correctionDefinition.length > 0;
    }
    if (this.state.currentRequestType == "Objection") {
      isValid = isValid && this.state.processingType != null && this.state.processingType.length > 0;
    }

    return(isValid);
  }

  @autobind
  private _closeInsertDialogResult() {
    this.state.showDialogResult = false;
    this._updateState(this.state);
  }

  @autobind
  private _insertNextClick(event) {
    event.preventDefault();
    this._closeInsertDialogResult();
  }

  @autobind
  private _goHomeClick(event) {
    event.preventDefault();
    pnp.sp.web.select("Url").get().then((web) => {
      window.location.replace(web.Url);
    });
  }
}
