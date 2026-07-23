import * as React from 'react';
import styles from './LeaveRequest.module.scss';
import { ILeaveRequestProps } from './ILeaveRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';


export interface ILeaveRequestState {

}

require('../assets/style.css');

export default class LeaveRequest extends React.Component<ILeaveRequestProps, ILeaveRequestState> {

  constructor(props: ILeaveRequestProps, state: ILeaveRequestState) {
    
    super(props);

    this.state = {
      
    };
  }

  public render(): React.ReactElement<ILeaveRequestProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="leaveRequest">
       
      </section>
    );
  }
}
