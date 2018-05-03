import * as React from 'react';

import { ISPDateTimePickerProps } from './ISPDateTimePickerProps';
import { ISPDateTimePickerState } from './ISPDateTimePickerState';

/**
 * Common Infrastructure
 */
import {
  BaseComponent,
  assign,
  autobind
} from 'office-ui-fabric-react/lib/Utilities';

/**
 * Label
 */
import { Label } from 'office-ui-fabric-react/lib/Label';

/**
 * Text Field
 */
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

/**
 * Date Picker
 */
import {
  DatePicker,
  DayOfWeek,
  IDatePickerStrings
} from 'office-ui-fabric-react/lib/DatePicker';

const DayPickerStrings: IDatePickerStrings = {
  months: [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ],

  shortMonths: [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec'
  ],

  days: [
    'Sunday',
    'Monday',
    'Tuesday',
    'Wednesday',
    'Thursday',
    'Friday',
    'Saturday'
  ],

  shortDays: [
    'S',
    'M',
    'T',
    'W',
    'T',
    'F',
    'S'
  ],

  goToToday: 'Go to today',
  isRequiredErrorMessage: 'Field is required.',
  invalidInputErrorMessage: 'Invalid date format.'
};

export interface IDatePickerRequiredExampleState {
  firstDayOfWeek?: DayOfWeek;
}

export class SPDateTimePicker extends React.Component<ISPDateTimePickerProps, ISPDateTimePickerState> {

  /**
   * Constructor
   */
  constructor(props: ISPDateTimePickerProps) {
    super(props);

    this.state = {
      date: (this.props.initialDateTime != null) ? new Date(this.props.initialDateTime) : null,
      hours: (this.props.initialDateTime != null) ? new Date(this.props.initialDateTime).getHours() : 0,
      minutes: (this.props.initialDateTime != null) ? new Date(this.props.initialDateTime).getMinutes() : 0,
      seconds: (this.props.initialDateTime != null) ? new Date(this.props.initialDateTime).getSeconds() : 0,
    };
  }

  public render(): React.ReactElement<ISPDateTimePickerProps> {
    return (
      <div className="ms-Grid">
        <div className="ms-Grid-row">
          <div className={(this.props.showTime) ? "ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg6" : "ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"}>
            <DatePicker
              value={this.state.date}
              onSelectDate={this._dateSelected}
              label={this.props.dateLabel}
              isRequired={this.props.isRequired}
              firstDayOfWeek={DayOfWeek.Sunday}
              strings={DayPickerStrings}
              placeholder={this.props.datePlaceholder} />
          </div>
          {(this.props.showTime) ?
            <div className={(this.props.includeSeconds) ? "ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2" : "ms-Grid-col ms-u-sm3 ms-u-md3 ms-u-lg3"}>
              <TextField
                type="number"
                label={this.props.hoursLabel}
                onChanged={this._hoursChanged}
                onGetErrorMessage={this._getErrorMessageHours}
                min="0"
                max="23" />
            </div>
            : null
          }
          {(this.props.showTime) ?
            <div className={(this.props.includeSeconds) ? "ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2" : "ms-Grid-col ms-u-sm3 ms-u-md3 ms-u-lg3"}>
              <TextField
                type="number"
                label={this.props.minutesLabel}
                onChanged={this._minutesChanged}
                onGetErrorMessage={this._getErrorMessageMinutes}
                min="0"
                max="59" />
            </div>
            : null
          }
          {(this.props.showTime && this.props.includeSeconds) ?
            <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2">
              <TextField
                type="number"
                label={this.props.secondsLabel}
                onChanged={this._secondsChanged}
                onGetErrorMessage={this._getErrorMessageSeconds}
                min="0"
                max="59" />
            </div>
            : null
          }
        </div>
      </div>
    );
  }

  @autobind
  private _dateSelected(date: Date): void {
    if (date == null)
      return;
    //this.state.date = date;
    //this.setState(this.state);
    this.setState((current) => ({ ...current, date: date }));
    this.saveFullDate();
  }

  @autobind
  private _hoursChanged(value: string): void {
    // this.state.hours = Number(value);
    // this.setState(this.state);
    this.setState((current) => ({
      ...current,
      hours: Number(value)
      
    }));
    
    this.saveFullDate();
  }

  @autobind
  private _getErrorMessageHours(value: string): string {
    let hoursValue = Number(value);
    return (hoursValue >= 0 && hoursValue <= 23)
      ? ''
      : `${this.props.hoursValidationError}.`;
  }

  @autobind
  private _minutesChanged(newValue: string): void {
    // this.state.minutes = Number(newValue);
    // this.setState(this.state);
    this.setState((current) => ({
      ...current,
      minutes: Number(newValue)
      
    }));
    
    this.saveFullDate();
  }

  @autobind
  private _getErrorMessageMinutes(value: string): string {
    let minutesValue = Number(value);
    return (minutesValue >= 0 && minutesValue <= 59)
      ? ''
      : `${this.props.minutesValidationError}.`;
  }

  @autobind
  private _secondsChanged(newValue: string): void {
    // this.state.seconds = Number(newValue);
    // this.setState(this.state);
    this.setState((current) => ({
      ...current,
      seconds: Number(newValue)
      
    }));
    this.saveFullDate();
  }

  @autobind
  private _getErrorMessageSeconds(value: string): string {
    let secondsValue = Number(value);
    return (secondsValue >= 0 && secondsValue <= 59)
      ? ''
      : `${this.props.secondsValidationError}.`;
  }

  private saveFullDate(): void {
    if (this.state.date == null)
      return;
    var finalDate = new Date(this.state.date.toISOString());
    finalDate.setHours(this.state.hours);
    finalDate.setMinutes(this.state.minutes);
    finalDate.setSeconds(this.props.includeSeconds ? this.state.seconds : 0);

    if (finalDate != null) {
      var finalDateAsString: string = '';
      if (this.props.formatDate) {
        finalDateAsString = this.props.formatDate(finalDate);
      }
      else {
        finalDateAsString = finalDate.toString();
      }
    }

    this.setState((current) => ({
      ...current,
      fullDate: finalDateAsString
      
    }));

    if (this.props.onChanged != null) {
      this.props.onChanged(finalDate);
    }
  }
}