import * as React from 'react';
import styles from './LeaveRequest.module.scss';
import { ILeaveRequestProps } from './ILeaveRequestProps';
import { DetailsList, IColumn, Pivot, PivotItem, SelectionMode } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { escape } from '@microsoft/sp-lodash-subset';
import { LeaveService, INewLeavePayload, IBalanceUpdate } from '../services/LeaveService';
import { formatDisplayDate, calcBusinessDays, isSameCalendarDay, parseIsoDate } from '../services/dateUtils';
import {
  IEmployeeInfo,
  ILeaveBalance,
  ILeaveHistoryItem,
  ILeaveFormValues,
  DurationValue,
  LeaveTypeValue,
  emptyLeaveForm,
  getLeaveStatus
} from '../models/ILeaveModels';


export interface ILeaveRequestState {
  loading: boolean;
  errorMessage: string;
  successMessage: string;

  currentEmployee?: IEmployeeInfo;
  balance?: ILeaveBalance;
  isHRTeam: boolean;
  isReportingManager: boolean;
  directReports: IEmployeeInfo[];

  allLeaves: ILeaveHistoryItem[];
  pendingApproval: ILeaveHistoryItem[];

  activeTab: 'approval' | 'myPending' | 'history';
  searchTerm: string;
  sortColumn: string;
  sortDescending: boolean;

  isFormOpen: boolean;
  formValues: ILeaveFormValues;
  formErrors: string[];
  totalDays: number;
  isSaving: boolean;

  formOwnerBalance?: ILeaveBalance;

  isCommentsOpen: boolean;
  commentsAction: 'approve' | 'reject' | '';
  commentsTarget?: ILeaveHistoryItem;
  comments: string;
}

require('../assets/style.css');

export default class LeaveRequest extends React.Component<ILeaveRequestProps, ILeaveRequestState> {

  private service: LeaveService;

  constructor(props: ILeaveRequestProps, state: ILeaveRequestState) {
    
    super(props);

    this.service = new LeaveService();

    this.state = {

      loading: true,
      errorMessage: '',
      successMessage: '',
      isHRTeam: false,
      isReportingManager: false,
      directReports: [],
      allLeaves: [],
      pendingApproval: [],
      activeTab: 'myPending',
      searchTerm: '',
      sortColumn: 'StartDate',
      sortDescending: false,
      isFormOpen: false,
      formValues: emptyLeaveForm(),
      formErrors: [],
      totalDays: 0,
      isSaving: false,
      isCommentsOpen: false,
      commentsAction: '',
      comments: ''
      
    };
  }

  public componentDidMount(): void {
    this.loadAll().catch((e) => this.setState({ errorMessage: this.errText(e), loading: false }));
  }

  private errText(e: any): string {
    return e && e.message ? e.message : 'Something went wrong. Please try again.';
  }

  /* ============================= data loading ============================= */

  private async loadAll(): Promise<void> {
    this.setState({ loading: true, errorMessage: '' });

    const userEmail = this.props.context.pageContext.user.email || this.props.context.pageContext.user.loginName;

    const currentEmployee = await this.service.getCurrentEmployee(userEmail);
    const balance = await this.service.getLeaveBalance(userEmail);
    const isHRTeam = await this.service.isCurrentUserInHRTeam();

    let isReportingManager = false;
    let directReports: IEmployeeInfo[] = [];

    if (!isHRTeam && balance) {
      directReports = await this.service.getDirectReports(balance.EmployeeNameId);
      isReportingManager = directReports.length > 0;
    }

    await new Promise<void>((resolve) => {
      this.setState({ currentEmployee, balance, isHRTeam, isReportingManager, directReports }, resolve);
    });

    try {
      await this.refreshTables();
    } finally {
      this.setState({ loading: false });
    }
  }

  private async refreshTables(): Promise<void> {
    const { isHRTeam, balance, isReportingManager, directReports } = this.state;
    const employeeId = balance ? balance.EmployeeNameId : 0;

    const allLeaves = await this.service.getAllLeaves(isHRTeam, employeeId);

    let pendingApproval: ILeaveHistoryItem[] = [];
    if (isHRTeam || isReportingManager) {
      pendingApproval = await this.service.getPendingApprovalLeaves(
        isHRTeam,
        directReports.map((r) => r.Id)
      );
    }

    this.setState({ allLeaves, pendingApproval });
  }

  /* ============================= derived data ============================= */

  private myPendingLeaves(): ILeaveHistoryItem[] {
    const { allLeaves, balance, isHRTeam } = this.state;
    return allLeaves.filter(
      (item) =>
        !item.isApproved &&
        !item.isRejected &&
        !item.isCancelled &&
        (!isHRTeam || (balance && item.EmpNameId === balance.EmployeeNameId))
    );
  }

  private currentTabRows(): ILeaveHistoryItem[] {
    const { activeTab, pendingApproval, allLeaves } = this.state;
    if (activeTab === 'approval') return pendingApproval;
    if (activeTab === 'myPending') return this.myPendingLeaves();
    return allLeaves;
  }

  private searchFilter(items: ILeaveHistoryItem[]): ILeaveHistoryItem[] {
    const term = this.state.searchTerm.trim().toLowerCase();
    if (!term) return items;
    return items.filter((item) => {
      const empName = item.EmpName ? item.EmpName.Title : '';
      return (
        empName.toLowerCase().indexOf(term) > -1 ||
        (item.LeaveType || '').toLowerCase().indexOf(term) > -1 ||
        (item.Description || '').toLowerCase().indexOf(term) > -1
      );
    });
  }

  /* ============================= form (add / edit) ============================= */

  private recalcTotals(values: ILeaveFormValues): { totalDays: number } {
    const start = parseIsoDate(values.startDate);
    const end = parseIsoDate(values.endDate);
    if (!start || !end || end < start) return { totalDays: 0 };

    const sameDay = isSameCalendarDay(start, end);
    if (sameDay) {
      return { totalDays: values.duration === 'First half' || values.duration === 'Second half' ? 0.5 : 1 };
    }

    const base = calcBusinessDays(start, end);
    if (values.startDateInterval === 'First half' && values.endDateInterval === 'First half') {
      return { totalDays: base - 0.5 };
    }
    if (values.startDateInterval === 'Second half' && values.endDateInterval === 'First half') {
      return { totalDays: base - 1 };
    }
    if (values.startDateInterval === 'Second half' && values.endDateInterval === 'Second half') {
      return { totalDays: base - 0.5 };
    }
    return { totalDays: base };
  }

  private openNewLeaveForm = (): void => {
    this.setState({ isFormOpen: true, formValues: emptyLeaveForm(), formErrors: [], totalDays: 0, formOwnerBalance: this.state.balance });
  };

  private openEditLeaveForm = async (itemId: number): Promise<void> => {
    const item = await this.service.getLeaveById(itemId);
    if (!item) return;

    const ownerBalance = await this.service.getLeaveBalanceByEmployeeId(item.Id);

    const notifyUsers:any = item.NotifyUsers ? item.NotifyUsers : [];
    const values: ILeaveFormValues = {
      itemId: item.Id,
      startDate: item.StartDate ? item.StartDate.substring(0, 10) : '',
      endDate: item.EndDate ? item.EndDate.substring(0, 10) : '',
      duration: (item.Duration as any) || '',
      startDateInterval: (item.StartDateInterval as any) || 'First half',
      endDateInterval: (item.EndDateInterval as any) || 'Second half',
      leaveType: item.LeaveType,
      description: item.Description,
      notifyUsers: notifyUsers.map((u) => ({ id: u.Id, text: u.Title, loginName: u.Title }))
    };
    const { totalDays } = this.recalcTotals(values);
    this.setState({ isFormOpen: true, formValues: values, formErrors: [], totalDays, formOwnerBalance: ownerBalance });
  };

  private closeForm = (): void => {
    this.setState({ isFormOpen: false, formValues: emptyLeaveForm(), formErrors: [] });
  };

  private onDateChange = (field: 'startDate' | 'endDate', value: string): void => {
    const values: ILeaveFormValues = { ...this.state.formValues, [field]: value } as ILeaveFormValues;

    const start = parseIsoDate(values.startDate);
    const end = parseIsoDate(values.endDate);

    if (start && end && end < start) {
      this.setState({ formValues: values, formErrors: ['End Date should not be less than start date.'], totalDays: 0 });
      return;
    }

    if (start && end && isSameCalendarDay(start, end)) {
      values.duration = values.duration || 'Full day';
    } else if (start && end) {
      values.startDateInterval = values.startDateInterval || 'First half';
      values.endDateInterval = values.endDateInterval || 'Second half';
    }

    const { totalDays } = this.recalcTotals(values);
    this.setState({ formValues: values, formErrors: [], totalDays });
  };

  private onFieldChange = (field: keyof ILeaveFormValues, value: any): void => {
    const values: ILeaveFormValues = { ...this.state.formValues, [field]: value };
    const { totalDays } = this.recalcTotals(values);
    this.setState({ formValues: values, totalDays });
  };

  private onNotifyUsersChange = (items: any[]): void => {
    this.setState({
      formValues: {
        ...this.state.formValues,
        notifyUsers: items.map((i) => ({
          id: i.id ? parseInt(i.id, 10) : 0,
          text: i.text,
          loginName: i.loginName || i.secondaryText || i.text
        }))
      }
    });
  };

  private validate(): string[] {
    const { formValues, totalDays, balance } = this.state;
    const errors: string[] = [];
    if (!formValues.startDate) errors.push('Start date is required.');
    if (!formValues.endDate) errors.push('End date is required.');
    if (!formValues.leaveType) errors.push('Leave Type is required.');
    if (!formValues.description) errors.push('Description is required.');
    if (formValues.notifyUsers.length === 0) errors.push('Notify Users is required.');

    const remainingPL = balance ? balance.PL : 0;
    if (formValues.leaveType === 'Paid Leave' && totalDays > remainingPL) {
      errors.push(`You do not have sufficient leaves left to apply for selected days. You have only ${remainingPL} leave remaining.`);
    }
    return errors;
  }

  private onSubmitLeaveForm = async (): Promise<void> => {
    const errors = this.validate();
    if (errors.length > 0) {
      this.setState({ formErrors: errors });
      return;
    }

    // const { balance, currentEmployee, formValues, totalDays } = this.state;
    // if (!balance || !currentEmployee) return;

    // this.setState({ isSaving: true, errorMessage: '' });
    // try {

      

    //   const start = parseIsoDate(formValues.startDate) as Date;
    //   const end = parseIsoDate(formValues.endDate) as Date;
    //   const sameDay = isSameCalendarDay(start, end);

    //   const notifyLoginNames = formValues.notifyUsers.map((u) => u.loginName);
    //   const notifyIds = await this.service.ensureUsers(notifyLoginNames);

    //   // IMPORTANT: don't use start.toISOString() / end.toISOString() here.
    //   // parseIsoDate() builds a *local* midnight Date, and .toISOString()
    //   // converts that to UTC — in any timezone ahead of UTC (e.g. IST) that
    //   // rolls the date back by one day before it ever reaches SharePoint.
    //   // Sending the yyyy-MM-dd value as UTC midnight keeps the calendar date
    //   // stable end-to-end (save -> read back on edit -> re-save).
    //   const payload: INewLeavePayload = {
    //     EmpNameId: balance.EmployeeNameId,
    //     StartDate: `${formValues.startDate}T00:00:00.000Z`,
    //     EndDate: `${formValues.endDate}T00:00:00.000Z`,
    //     Description: formValues.description,
    //     LeaveType: formValues.leaveType,
    //     TotalDays: totalDays.toString(),
    //     Duration: sameDay ? formValues.duration || null : null,
    //     StartDateInterval: sameDay ? null : formValues.startDateInterval,
    //     EndDateInterval: sameDay ? null : formValues.endDateInterval,
    //     NotifyUsersId: { results: notifyIds }
    //   };

    //   let editLeaveType = '';
    //   let editTotalDays = 0;
    //   if (formValues.itemId > 0) {
    //     const original = await this.service.getLeaveById(formValues.itemId);
    //     if (original) {
    //       editLeaveType = original.LeaveType;
    //       editTotalDays = parseFloat(original.TotalDays);
    //     }
    //     await this.service.updateLeave(formValues.itemId, payload);
    //   } else {
    //     await this.service.createLeave(payload);
    //   }

    //   const balancePatch = this.computeBalancePatch({
    //     isEdit: formValues.itemId > 0,
    //     newLeaveType: formValues.leaveType,
    //     newDifference: totalDays,
    //     editLeaveType,
    //     editTotalDays,
    //     remainingPL: balance.PL,
    //     currentUnpaidLeaves: balance.UnpaidLeave
    //   });
    //   await this.service.updateLeaveBalance(balance.Id, balancePatch);

    //   this.setState({
    //     isFormOpen: false,
    //     formValues: emptyLeaveForm(),
    //     successMessage: formValues.itemId > 0 ? 'Leave details are successfully updated.' : 'Request for leave has been submitted.'
    //   });

    //   await this.loadAll();
    // } catch (e) {
    //   this.setState({ errorMessage: this.errText(e) });
    // } finally {
    //   this.setState({ isSaving: false });
    // }

    const { currentEmployee, formValues, totalDays } = this.state;
    if (!currentEmployee) return;

    this.setState({ isSaving: true, errorMessage: '' });
    try {
      const isEdit = formValues.itemId > 0;

      // Resolve who this leave actually belongs to, and THEIR balance.
      // Never assume it's the current viewer — HR/a manager can be editing
      // a direct report's leave. Using the viewer's own balance here was
      // the root cause of edits silently reassigning the leave and
      // debiting/crediting the wrong person.
      let original: ILeaveHistoryItem | undefined;
      let ownerBalance: ILeaveBalance | undefined;

      if (isEdit) {
        original = await this.service.getLeaveById(formValues.itemId);
        if (!original) {
          this.setState({ errorMessage: 'This leave request could no longer be found. It may have been deleted.' });
          return;
        }
        ownerBalance = await this.service.getLeaveBalanceByEmployeeId(original.EmpNameId);
      } else {
        ownerBalance = this.state.balance;
      }

      if (!ownerBalance) {
        this.setState({ errorMessage: 'Could not find a leave balance record for this employee.' });
        return;
      }

      const editLeaveType = original ? original.LeaveType : '';
      const editTotalDays = original ? parseFloat(original.TotalDays) : 0;

      // Days already reserved under the SAME leave type (on edit) count back
      // toward what's available, so editing without changing the day count
      // never falsely reports "insufficient balance".
      const availablePL = ownerBalance.PL + (isEdit && editLeaveType === 'Paid Leave' ? editTotalDays : 0);
      if (formValues.leaveType === 'Paid Leave' && totalDays > availablePL) {
        this.setState({
          isSaving: false,
          formErrors: [`You do not have sufficient leaves left to apply for selected days. You have only ${availablePL} leave remaining.`]
        });
        return;
      }

      const start = parseIsoDate(formValues.startDate) as Date;
      const end = parseIsoDate(formValues.endDate) as Date;
      const sameDay = isSameCalendarDay(start, end);

      const notifyLoginNames = formValues.notifyUsers.map((u) => u.loginName);
      const notifyIds = await this.service.ensureUsers(notifyLoginNames);

      // IMPORTANT: don't use start.toISOString() / end.toISOString() here.
      // parseIsoDate() builds a *local* midnight Date, and .toISOString()
      // converts that to UTC — in any timezone ahead of UTC (e.g. IST) that
      // rolls the date back by one day before it ever reaches SharePoint.
      // Sending the yyyy-MM-dd value as UTC midnight keeps the calendar date
      // stable end-to-end (save -> read back on edit -> re-save).
      const payload: INewLeavePayload = {
        EmpNameId: ownerBalance.EmployeeNameId,
        StartDate: `${formValues.startDate}T00:00:00.000Z`,
        EndDate: `${formValues.endDate}T00:00:00.000Z`,
        Description: formValues.description,
        LeaveType: formValues.leaveType,
        TotalDays: totalDays.toString(),
        Duration: sameDay ? formValues.duration || null : null,
        StartDateInterval: sameDay ? null : formValues.startDateInterval,
        EndDateInterval: sameDay ? null : formValues.endDateInterval,
        NotifyUsersId: { results: notifyIds }
      };

      if (isEdit) {
        await this.service.updateLeave(formValues.itemId, payload);
      } else {
        await this.service.createLeave(payload);
      }

      const balancePatch = this.computeBalancePatch({
        isEdit,
        newLeaveType: formValues.leaveType,
        newDifference: totalDays,
        editLeaveType,
        editTotalDays,
        remainingPL: ownerBalance.PL,
        currentUnpaidLeaves: ownerBalance.UnpaidLeave
      });
      await this.service.updateLeaveBalance(ownerBalance.Id, balancePatch);

      this.setState({
        isFormOpen: false,
        formValues: emptyLeaveForm(),
        formErrors: [],
        successMessage: isEdit ? 'Leave details are successfully updated.' : 'Request for leave has been submitted.'
      });

      await this.loadAll();
    } catch (e) {
      this.setState({ errorMessage: this.errText(e) });
    } finally {
      this.setState({ isSaving: false });
    }
  };

  /** Mirrors the legacy balance-adjustment logic: restores days under the old
   * leave type (on edit) and reserves days under the new one. */
  private computeBalancePatch(params: {
    isEdit: boolean;
    newLeaveType: string;
    newDifference: number;
    editLeaveType: string;
    editTotalDays: number;
    remainingPL: number;
    currentUnpaidLeaves: number;
  }): IBalanceUpdate {
    const { isEdit, newLeaveType, newDifference, editLeaveType, editTotalDays, remainingPL, currentUnpaidLeaves } = params;
    const patch: IBalanceUpdate = {};

    if (isEdit && editLeaveType && editLeaveType !== newLeaveType) {
      if (editLeaveType === 'Paid Leave') patch.PL = (remainingPL + editTotalDays).toString();
      else if (editLeaveType === 'Unpaid Leave') patch.UnpaidLeave = (currentUnpaidLeaves - editTotalDays).toString();

      if (newLeaveType === 'Paid Leave') patch.PL = (remainingPL - newDifference).toString();
      else if (newLeaveType === 'Unpaid Leave') patch.UnpaidLeave = (currentUnpaidLeaves + newDifference).toString();
    } else if (isEdit) {
      if (newLeaveType === 'Paid Leave') patch.PL = (remainingPL + editTotalDays - newDifference).toString();
      else if (newLeaveType === 'Unpaid Leave') patch.UnpaidLeave = (currentUnpaidLeaves - editTotalDays + newDifference).toString();
    } else {
      if (newLeaveType === 'Paid Leave') patch.PL = (remainingPL - newDifference).toString();
      else if (newLeaveType === 'Unpaid Leave') patch.UnpaidLeave = (currentUnpaidLeaves + newDifference).toString();
    }
    return patch;
  }

  /* ============================= cancel / delete ============================= */

  private cancelLeave = async (itemId: number): Promise<void> => {
    // eslint-disable-next-line no-alert
    if (!window.confirm('Are you sure you want to delete the leave?')) return;

    this.setState({ isSaving: true, errorMessage: '' });
    try {
      const item = await this.service.getLeaveById(itemId);

        if (!item) return;

        // Prevent double cancel
        if (item.isCancelled || item.isRejected) {
            this.setState({
                errorMessage: "Leave has already been cancelled or rejected."
            });
            return;
        }

        await this.service.cancelLeave(itemId);

      const empBalance = await this.service.getLeaveBalanceByEmployeeId(item.EmpNameId);
      // if (empBalance) {
      //   const totalDays = parseFloat(item.TotalDays);
      //   const patch: IBalanceUpdate = {};
      //   if (item.LeaveType === 'Paid Leave') patch.PL = (empBalance.PL + totalDays).toString();
      //   else if (item.LeaveType === 'Unpaid Leave') patch.UnpaidLeave = (empBalance.UnpaidLeave - totalDays).toString();
      //   await this.service.updateLeaveBalance(empBalance.Id, patch);
      // }

      if (empBalance) {

        const totalDays = parseFloat(item.TotalDays);
    
        const patch: IBalanceUpdate = {};
    
        if (item.LeaveType === "Paid Leave") {
    
            patch.PL = (empBalance.PL + totalDays).toString();
    
        } else if (item.LeaveType === "Unpaid Leave") {
    
            patch.UnpaidLeave = (
                empBalance.UnpaidLeave - totalDays
            ).toString();
    
        }
    
        await this.service.updateLeaveBalance(
            empBalance.Id,
            patch
        );
    }

      this.setState({ successMessage: 'Leave has been deleted.' });
      await this.loadAll();
    } catch (e) {
      this.setState({ errorMessage: this.errText(e) });
    } finally {
      this.setState({ isSaving: false });
    }
  };

  /* ============================= approve / reject ============================= */

  private openComments = (item: ILeaveHistoryItem, action: 'approve' | 'reject'): void => {
    this.setState({ isCommentsOpen: true, commentsAction: action, commentsTarget: item, comments: '' });
  };

  private closeComments = (): void => {
    this.setState({ isCommentsOpen: false, commentsAction: '', commentsTarget: undefined, comments: '' });
  };

  private submitComments = async (): Promise<void> => {
    const { commentsAction, commentsTarget, currentEmployee, comments } = this.state;
    if (!commentsTarget || !currentEmployee) return;

    this.setState({ isSaving: true, errorMessage: '' });
    try {
      if (commentsAction === 'approve') {
        await this.service.setApproval(commentsTarget.Id, true, currentEmployee.Title, comments);
        this.setState({ successMessage: 'Leave has been approved.' });
      } else {

        // Prevent double reject
        const latestLeave = await this.service.getLeaveById(commentsTarget.Id);
    
        if (latestLeave && (latestLeave.isRejected || latestLeave.isCancelled)) {
            this.setState({
                errorMessage: "This leave request has already been processed."
            });
            return;
        }
    
        await this.service.setApproval(
            commentsTarget.Id,
            false,
            currentEmployee.Title,
            comments
        );
    
        const empBalance = await this.service.getLeaveBalanceByEmployeeId(
            commentsTarget.EmpName.Id
        );
    
        if (empBalance) {
    
            const totalDays = parseFloat(commentsTarget.TotalDays);
    
            const patch: IBalanceUpdate = {};
    
            if (commentsTarget.LeaveType === "Paid Leave") {
    
                patch.PL = (empBalance.PL + totalDays).toString();
    
            } else if (commentsTarget.LeaveType === "Unpaid Leave") {
    
                patch.UnpaidLeave = (
                    empBalance.UnpaidLeave - totalDays
                ).toString();
    
            }
    
            await this.service.updateLeaveBalance(empBalance.Id, patch);
        }
    
        this.setState({
            successMessage: "Leave request has been rejected."
        });
    }
      this.closeComments();
      await this.loadAll();
    } catch (e) {
      this.setState({ errorMessage: this.errText(e) });
    } finally {
      this.setState({ isSaving: false });
    }
  };

  /* ============================= render helpers ============================= */

  // private renderRow(item: ILeaveHistoryItem, mode: 'approval' | 'myPending' | 'history'): React.ReactElement {
  //   const { status, className } = getLeaveStatus(item);
  //   const { isHRTeam } = this.state;

  //   return (
  //     <tr key={item.Id}>
  //       {(mode !== 'myPending') && <td>{item.EmpName ? item.EmpName.Title : ''}</td>}
  //       <td>{formatDisplayDate(item.StartDate)}</td>
  //       <td>{formatDisplayDate(item.EndDate)}</td>
  //       <td>{item.LeaveType}</td>
  //       <td>{item.Description}</td>
  //       {mode === 'history' && <td>{item.ApproverName || ''}</td>}
  //       {mode === 'history' && (
  //         <td>
  //           <span className={`status-badge ${className}`}>{status}</span>
  //         </td>
  //       )}
  //       <td>
  //         {mode === 'approval' && (
  //           <>
  //             <button className="btn btn-primary" onClick={() => this.openComments(item, 'approve')}>Approve</button>
  //             <button className="btn btn-danger" onClick={() => this.openComments(item, 'reject')}>Reject</button>
  //           </>
  //         )}
  //         {mode === 'myPending' && (
  //           <>
  //             <button className="btn btn-primary" onClick={() => this.openEditLeaveForm(item.Id)}>Edit</button>
  //             <button className="btn btn-danger" onClick={() => this.cancelLeave(item.Id)}>Delete</button>
  //           </>
  //         )}
  //         {mode === 'history' && status === 'Pending' && (
  //           <>
  //             {isHRTeam && <button className="btn btn-primary" onClick={() => this.openEditLeaveForm(item.Id)}>Edit</button>}
  //             <button className="btn btn-danger" onClick={() => this.cancelLeave(item.Id)}>Cancel</button>
  //           </>
  //         )}
  //       </td>
  //     </tr>
  //   );
  // }

  private sortRows(items: ILeaveHistoryItem[]): ILeaveHistoryItem[] {
    const { sortColumn, sortDescending } = this.state;
    const sorted = items.slice();
    sorted.sort((a, b) => {
      let cmp = 0;
      switch (sortColumn) {
        case 'EmpName':
          cmp = (a.EmpName ? a.EmpName.Title : '').localeCompare(b.EmpName ? b.EmpName.Title : '');
          break;
        case 'StartDate':
          cmp = new Date(a.StartDate).getTime() - new Date(b.StartDate).getTime();
          break;
        case 'EndDate':
          cmp = new Date(a.EndDate).getTime() - new Date(b.EndDate).getTime();
          break;
        case 'LeaveType':
          cmp = (a.LeaveType || '').localeCompare(b.LeaveType || '');
          break;
        case 'Description':
          cmp = (a.Description || '').localeCompare(b.Description || '');
          break;
        case 'ApproverName':
          cmp = (a.ApproverName || '').localeCompare(b.ApproverName || '');
          break;
        case 'Status':
          cmp = getLeaveStatus(a).status.localeCompare(getLeaveStatus(b).status);
          break;
        default:
          cmp = 0;
      }
      return sortDescending ? -cmp : cmp;
    });
    return sorted;
  }


  private onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { sortColumn, sortDescending } = this.state;
    const isSameColumn = sortColumn === column.key;
    this.setState({
      sortColumn: column.key,
      sortDescending: isSameColumn ? !sortDescending : false
    });
  };

  private getColumns(mode: 'approval' | 'myPending' | 'history'): IColumn[] {
    const { isHRTeam } = this.state;
    const { sortColumn, sortDescending } = this.state;

    const sortableHeader = (key: string, name: string, minWidth: number): IColumn => ({
      key,
      name,
      fieldName: key,
      minWidth,
      isResizable: true,
      isSorted: sortColumn === key,
      isSortedDescending: sortDescending,
      onColumnClick: this.onColumnClick
    });

    const columns: IColumn[] = [];

    if (mode !== 'myPending') {
      columns.push({
        ...sortableHeader('EmpName', 'Employee Name', 140),
        onRender: (item: ILeaveHistoryItem) => (item.EmpName ? item.EmpName.Title : '')
      });
    }

    columns.push({
      ...sortableHeader('StartDate', 'Start Date', 100),
      onRender: (item: ILeaveHistoryItem) => formatDisplayDate(item.StartDate)
    });
    columns.push({
      ...sortableHeader('EndDate', 'End Date', 100),
      onRender: (item: ILeaveHistoryItem) => formatDisplayDate(item.EndDate)
    });
    columns.push(sortableHeader('LeaveType', 'Leave Type', 100));
    columns.push(sortableHeader('Description', 'Description', 180));

    if (mode === 'history') {
      columns.push({
        ...sortableHeader('ApproverName', 'Approver', 100),
        onRender: (item: ILeaveHistoryItem) => item.ApproverName || ''
      });
      columns.push({
        ...sortableHeader('Status', 'Status', 100),
        onRender: (item: ILeaveHistoryItem) => {
          const { status, className } = getLeaveStatus(item);
          return <span className={`status-badge ${className}`}>{status}</span>;
        }
      });
    }

    columns.push({
      key: 'ActionItem',
      name: 'Action Item',
      fieldName: 'ActionItem',
      minWidth: 140,
      isResizable: true,
      onRender: (item: ILeaveHistoryItem) => {
        const { status } = getLeaveStatus(item);
        return (
          <>
            {mode === 'approval' && (
              <>
                <button className="btn btn-primary" onClick={() => this.openComments(item, 'approve')}>Approve</button>
                <button className="btn btn-danger" onClick={() => this.openComments(item, 'reject')}>Reject</button>
              </>
            )}
            {mode === 'myPending' && (
              <>
                <button className="btn btn-primary" onClick={() => this.openEditLeaveForm(item.Id)}>Edit</button>
                <button className="btn btn-danger" onClick={() => this.cancelLeave(item.Id)}>Delete</button>
              </>
            )}
            {mode === 'history' && status === 'Pending' && (
              <>
                {isHRTeam && <button className="btn btn-primary" onClick={() => this.openEditLeaveForm(item.Id)}>Edit</button>}
                <button className="btn btn-danger" onClick={() => this.cancelLeave(item.Id)}>Cancel</button>
              </>
            )}
          </>
        );
      }
    });

    return columns;
  }


  // private renderTable(mode: 'approval' | 'myPending' | 'history', emptyMessage: string): React.ReactElement {
  //   const rows = this.searchFilter(this.currentTabRows());
  //   const { isHRTeam } = this.state;

  //   return (
  //     <>
  //       <div className="leave-search">
  //         <input
  //           type="text"
  //           placeholder="Search"
  //           value={this.state.searchTerm}
  //           onChange={(e) => this.setState({ searchTerm: e.target.value })}
  //         />
  //       </div>

  //       {rows.length === 0 ? (
  //         <p className="leave-empty">{emptyMessage}</p>
  //       ) : (
  //         <div className="leave-table-wrap">
  //           <table className="leave-table">
  //             <thead>
  //               <tr>
  //                 {mode !== 'myPending' && <th>Employee Name</th>}
  //                 <th>Start Date</th>
  //                 <th>End Date</th>
  //                 <th>Leave Type</th>
  //                 <th>Description</th>
  //                 {mode === 'history' && <th>Approver</th>}
  //                 {mode === 'history' && <th>Status</th>}
  //                 <th>Action Item</th>
  //               </tr>
  //             </thead>
  //             <tbody>{rows.map((item) => this.renderRow(item, mode))}</tbody>
  //           </table>
  //         </div>
  //       )}
  //     </>
  //   );
  // }

  private renderTable(mode: 'approval' | 'myPending' | 'history', emptyMessage: string): React.ReactElement {
    const rows = this.sortRows(this.searchFilter(this.currentTabRows()));

    return (
      <>
        <div className="leave-search">
          <input
            type="text"
            placeholder="Search"
            aria-label="Search"
            value={this.state.searchTerm}
            onChange={(e) => this.setState({ searchTerm: e.target.value })}
          />
        </div>

        {rows.length === 0 ? (
          <p className="leave-empty">{emptyMessage}</p>
        ) : (
          <div className="leave-table-wrap">
            <DetailsList
              items={rows}
              columns={this.getColumns(mode)}
              selectionMode={SelectionMode.none}
              getKey={(item: ILeaveHistoryItem) => item.Id.toString()}
              setKey={mode}
            />
          </div>
        )}
      </>
    );
  }

  private renderLeaveForm(): React.ReactElement | null {
    if (!this.state.isFormOpen) return null;
    const { formValues, formErrors, totalDays, isSaving, formOwnerBalance } = this.state;
    const sameDay =
      !!formValues.startDate && !!formValues.endDate && formValues.startDate === formValues.endDate;
    const showIntervals = !!formValues.startDate && !!formValues.endDate && !sameDay;
    const offerPaidLeave = !!formOwnerBalance && formOwnerBalance.PL > 0;

    return (
      <div className="leave-overlay">
        <div className="leave-form-panel">
          <div className="panel-title">{formValues.itemId > 0 ? 'Edit Leave Request' : 'Request Leave'}</div>

          {formErrors.length > 0 && (
            <div className="form-errors">
              {formErrors.map((e, i) => (
                <div key={i}>{e}</div>
              ))}
            </div>
          )}

          <div className="form-row">
            <div className="form-group">
              <label className="required">Start Date</label>
              <input
                type="date"
                value={formValues.startDate}
                min={new Date().toISOString().substring(0, 10)}
                onChange={(e) => this.onDateChange('startDate', e.target.value)}
              />
            </div>
            <div className="form-group">
              <label className="required">End Date</label>
              <input
                type="date"
                value={formValues.endDate}
                min={formValues.startDate || new Date().toISOString().substring(0, 10)}
                onChange={(e) => this.onDateChange('endDate', e.target.value)}
              />
            </div>
          </div>

          {sameDay && (
            <div className="form-group">
              <label>Duration</label>
              <div className="radio-group">
                {(['Full day', 'First half', 'Second half'] as DurationValue[]).map((d) => (
                  <label key={d}>
                    <input
                      type="radio"
                      name="duration"
                      checked={formValues.duration === d}
                      onChange={() => this.onFieldChange('duration', d)}
                    />
                    {d}
                  </label>
                ))}
              </div>
            </div>
          )}

          {showIntervals && (
            <div className="form-row">
              <div className="form-group">
                <label>Start Date covers</label>
                <div className="radio-group">
                  {(['First half', 'Second half'] as const).map((d) => (
                    <label key={d}>
                      <input
                        type="radio"
                        name="startInterval"
                        checked={formValues.startDateInterval === d}
                        onChange={() => this.onFieldChange('startDateInterval', d)}
                      />
                      {d}
                    </label>
                  ))}
                </div>
              </div>
              <div className="form-group">
                <label>End Date covers</label>
                <div className="radio-group">
                  {(['First half', 'Second half'] as const).map((d) => (
                    <label key={d}>
                      <input
                        type="radio"
                        name="endInterval"
                        checked={formValues.endDateInterval === d}
                        onChange={() => this.onFieldChange('endDateInterval', d)}
                      />
                      {d}
                    </label>
                  ))}
                </div>
              </div>
            </div>
          )}

          {totalDays > 0 && <div className="total-days">Total days: {totalDays}</div>}

          <div className="form-group">
            <label className="required">Leave Type</label>
            <select
              value={formValues.leaveType}
              onChange={(e) => this.onFieldChange('leaveType', e.target.value as LeaveTypeValue)}
            >
              <option value="">Select Leave Type</option>
              {(offerPaidLeave || formValues.leaveType === 'Paid Leave') && (
                <option value="Paid Leave">Paid Leave</option>
              )}
              <option value="Unpaid Leave">Unpaid Leave</option>
            </select>
            {formOwnerBalance && (
              <small style={{ color: '#666' }}>
                Paid leave balance: {formOwnerBalance.PL} &nbsp;|&nbsp; Unpaid leave taken: {formOwnerBalance.UnpaidLeave}
              </small>
            )}
          </div>

          <div className="form-group">
            <label className="required">Notify Users</label>
            <PeoplePicker
              context={this.props.context}
              personSelectionLimit={50}
              showtooltip={true}
              required={true}
              principalTypes={[PrincipalType.User]}
              resolveDelay={300}
              defaultSelectedUsers={formValues.notifyUsers.map((u) => u.loginName)}
              // onChange={this.onNotifyUsersChange}
              onChange={(e) => this.setState({formValues : {...this.state.formValues, notifyUsers: e.map(u => ({id: u.id ? parseInt(u.id, 10) : 0, text: u.text, loginName: u.title || u.secondaryText || u.text}))}})}   
              ensureUser={true}
            />
          </div>

          <div className="form-group">
            <label className="required">Description</label>
            <textarea
              rows={4}
              value={formValues.description}
              onChange={(e) => this.onFieldChange('description', e.target.value)}
            />
          </div>

          <div className="form-actions">
            <button className="btn btn-primary" disabled={isSaving} onClick={this.onSubmitLeaveForm}>
              Request leave
            </button>
            <button className="btn btn-secondary" disabled={isSaving} onClick={this.closeForm}>
              Cancel
            </button>
          </div>
        </div>
      </div>
    );
  }

  private renderCommentsDialog(): React.ReactElement | null {
    if (!this.state.isCommentsOpen) return null;
    const { commentsAction, comments, isSaving } = this.state;
    const title = commentsAction === 'approve' ? 'Approve Leave Request' : 'Reject Leave Request';

    return (
      <div className="leave-overlay comments-overlay">
        <div className="comments-panel">
          <div className="panel-title">{title}</div>
          <div className="form-group">
            <label>Comments</label>
            <textarea rows={4} value={comments} onChange={(e) => this.setState({ comments: e.target.value })} />
          </div>
          <div className="form-actions">
            <button className="btn btn-primary" disabled={isSaving} onClick={this.submitComments}>
              Submit
            </button>
            <button className="btn btn-secondary" disabled={isSaving} onClick={this.closeComments}>
              Cancel
            </button>
          </div>
        </div>
      </div>
    );
  }

  public render(): React.ReactElement<ILeaveRequestProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const { loading, errorMessage, successMessage, balance, isHRTeam, isReportingManager } = this.state;

    if (loading) {
      return (
        <div className="leave-spinner-wrap">
          <div className="leave-spinner" />
          <span>Loading leave dashboard…</span>
        </div>
      );
    }

    const showApprovalTab = isHRTeam || isReportingManager;

    return (
      <section className='leaveManagementSystem'>
        <div className="leave-header">
          <h2 className="section-title">{this.props.description || 'Leave Management'}</h2>
        </div>
        <div className="title-underline"></div>

        {errorMessage && (
          <div className="message-bar error" onClick={() => this.setState({ errorMessage: '' })}>
            {errorMessage}
          </div>
        )}
        {successMessage && (
          <div className="message-bar success" onClick={() => this.setState({ successMessage: '' })}>
            {successMessage}
          </div>
        )}

        <div className="balance-cards">
          <div className="balance-card paid">
            <span className="balance-value">{balance ? balance.PL : 0}</span>
            <span className="balance-label">Paid Leaves Balance</span>
          </div>
          <div className="balance-card unpaid">
            <span className="balance-value">{balance ? balance.UnpaidLeave : 0}</span>
            <span className="balance-label">Unpaid Leaves applied</span>
          </div>
        </div>

        <div className="leave-toolbar">
          <button className="btn btn-primary" onClick={this.openNewLeaveForm}>+ Add new</button>
        </div>

        <div className="leave-filters">
          <Pivot
            selectedKey={this.state.activeTab}
            onLinkClick={(item) =>
              item && this.setState({ activeTab: item.props.itemKey as any, searchTerm: '' })
            }
          >
            {showApprovalTab && <PivotItem headerText="Pending for Approval" itemKey="approval" />}
            <PivotItem headerText="My Leaves Pending for Approval" itemKey="myPending" />
            <PivotItem headerText="Leave History" itemKey="history" />
          </Pivot>
        </div>

        {this.state.activeTab === 'approval' &&
          this.renderTable('approval', 'There are no leaves pending at your end.')}
        {this.state.activeTab === 'myPending' &&
          this.renderTable('myPending', 'There are no pending leaves as of now.')}
        {this.state.activeTab === 'history' &&
          this.renderTable('history', 'There are no leaves as of now.')}

        {this.renderLeaveForm()}
        {this.renderCommentsDialog()}
      </section>
    );
  }
  
}
