export interface IEmployeeInfo {
  Id: number;
  Title: string;
  EmailId: string;
  ReportsToId?: number;
}

export interface ILeaveBalance {
  Id: number;
  SL: number;
  PL: number;
  UnpaidLeave: number;
  EmployeeNameId: number;
}

export const LeaveType = {
  Paid: 'Paid Leave',
  Unpaid: 'Unpaid Leave',
  Sick: 'Sick Leave'
} as const;

export type LeaveTypeValue = typeof LeaveType[keyof typeof LeaveType];

export type DurationValue = 'Full day' | 'First half' | 'Second half' | '';

export type LeaveStatus = 'Pending' | 'Approved' | 'Rejected' | 'Cancelled';

export interface INotifyUser {
  Id: number;
  Title: string;
}

export interface ILeaveHistoryItem {
  Id: number;
  EmpNameId: number;
  EmpName?: { Title: string };
  StartDate: string;
  EndDate: string;
  LeaveType: LeaveTypeValue | '';
  Duration: DurationValue | null;
  StartDateInterval: 'First half' | 'Second half' | null;
  EndDateInterval: 'First half' | 'Second half' | null;
  TotalDays: string;
  Description: string;
  NotifyUsersId?: number[];
  NotifyUsers?: { results: INotifyUser[] };
  isApproved: boolean;
  isRejected: boolean;
  isCancelled: boolean;
  ApproverName: string | null;
  ApproverComment: string | null;
}

export interface ILeaveFormValues {
  itemId: number;
  startDate: string; // yyyy-MM-dd, matches a plain <input type="date">
  endDate: string;
  duration: DurationValue;
  startDateInterval: 'First half' | 'Second half';
  endDateInterval: 'First half' | 'Second half';
  leaveType: LeaveTypeValue | '';
  description: string;
  notifyUsers: { id: number; text: string; loginName: string }[];
}

export function emptyLeaveForm(): ILeaveFormValues {
  return {
    itemId: 0,
    startDate: '',
    endDate: '',
    duration: '',
    startDateInterval: 'First half',
    endDateInterval: 'Second half',
    leaveType: '',
    description: '',
    notifyUsers: []
  };
}

export function getLeaveStatus(item: ILeaveHistoryItem): { status: LeaveStatus; className: string } {
  if (item.isApproved) return { status: 'Approved', className: 'status-approved' };
  if (item.isRejected) return { status: 'Rejected', className: 'status-rejected' };
  if (item.isCancelled) return { status: 'Cancelled', className: 'status-cancelled' };
  return { status: 'Pending', className: 'status-pending' };
}
