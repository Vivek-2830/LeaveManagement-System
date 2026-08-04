import { sp } from '@pnp/sp/rest';
import { IEmployeeInfo, ILeaveBalance, ILeaveHistoryItem, LeaveTypeValue } from '../models/ILeaveModels';

const EMPLOYEE_INFO_LIST = 'Employee Information';
const LEAVE_BALANCE_LIST = 'Employee Leave Balance';
const LEAVE_HISTORY_LIST = 'Employee Leave History';
const HR_GROUP_NAME = 'HR Team';

export interface INewLeavePayload {
  EmpNameId: number;
  StartDate: string;
  EndDate: string;
  Description: string;
  LeaveType: LeaveTypeValue | '';
  TotalDays: string;
  Duration: string | null;
  StartDateInterval: string | null;
  EndDateInterval: string | null;
  NotifyUsersId?: { results: number[] };
}

export interface IBalanceUpdate {
  SL?: string;
  PL?: string;
  UnpaidLeave?: string;
}

/**
 * Data access layer, built the same way as the "News Announcements" reads in
 * HomePageAllNewsAnnouncementPage.tsx (`sp.web.lists.getByTitle(...)`).
 * NOTE: `sp.setup({ spfxContext })` must be called once, in the web part's
 * onInit(), before any method on this class is used.
 */
export class LeaveService {
  public async getCurrentEmployee(userLoginEmail: string): Promise<IEmployeeInfo | undefined> {
    const items = await sp.web.lists
      .getByTitle(EMPLOYEE_INFO_LIST)
      .items.select('Id', 'Title', 'EmailId', 'ReportsTo/Id')
      .expand('ReportsTo')
      .filter(`EmailId eq '${this.escape(userLoginEmail)}'`)
      .top(1)
      .get();
    if (items.length === 0) return undefined;
    const r = items[0];
    return {
      Id: r.Id,
      Title: r.Title,
      EmailId: r.EmailId,
      ReportsToId: r.ReportsTo ? r.ReportsTo.Id : undefined
    };
  }

  public async getDirectReports(managerEmployeeId: number): Promise<IEmployeeInfo[]> {
    const items = await sp.web.lists
      .getByTitle(EMPLOYEE_INFO_LIST)
      .items.select('Id', 'Title', 'EmailId', 'ReportsTo/Id')
      .expand('ReportsTo')
      .filter(`ReportsTo/Id eq ${managerEmployeeId}`)
      .top(5000)
      .orderBy('Id', false)
      .get();
    return items.map((r: any) => ({
      Id: r.Id,
      Title: r.Title,
      EmailId: r.EmailId,
      ReportsToId: r.ReportsTo ? r.ReportsTo.Id : undefined
    }));
  }

  public async isCurrentUserInHRTeam(): Promise<boolean> {
    try {
      const groups = await sp.web.currentUser.groups.get();
      return groups.some((g: any) => g.Title === HR_GROUP_NAME);
    } catch {
      return false;
    }
  }

  public async getLeaveBalance(userLoginEmail: string): Promise<ILeaveBalance | undefined> {
    const items = await sp.web.lists
      .getByTitle(LEAVE_BALANCE_LIST)
      .items.select('Id', 'SL', 'PL', 'UnpaidLeave', 'EmployeeName/Id', 'EmployeeName/Title', 'EmployeeName/EmailId')
      .expand('EmployeeName/EmailId')
      .filter(`EmployeeName/EmailId eq '${this.escape(userLoginEmail)}'`)
      .top(1)
      .orderBy('Id', false)
      .get();
    if (items.length === 0) return undefined;
    const r = items[0];
    return {
      Id: r.Id,
      SL: r.SL != null ? parseFloat(r.SL) : 0,
      PL: r.PL != null ? parseFloat(r.PL) : 0,
      UnpaidLeave: r.UnpaidLeave != null ? parseFloat(r.UnpaidLeave) : 0,
      EmployeeNameId: r.EmployeeName ? r.EmployeeName.Id : 0
    };
  }

  public async getLeaveBalanceByEmployeeId(employeeId: number): Promise<ILeaveBalance | undefined> {
    const items = await sp.web.lists
      .getByTitle(LEAVE_BALANCE_LIST)
      .items.select('Id', 'SL', 'PL', 'UnpaidLeave', 'EmployeeName/Id')
      .expand('EmployeeName')
      .filter(`EmployeeName/Id eq ${employeeId}`)
      .top(1)
      .get();
    if (items.length === 0) return undefined;
    const r = items[0];
    return {
      Id: r.Id,
      SL: r.SL ? parseFloat(r.SL) : 0,
      PL: r.PL ? parseFloat(r.PL) : 0,
      UnpaidLeave: r.UnpaidLeave ? parseFloat(r.UnpaidLeave) : 0,
      EmployeeNameId: r.EmployeeName ? r.EmployeeName.Id : 0
    };
  }

  public async updateLeaveBalance(balanceId: number, patch: IBalanceUpdate): Promise<void> {
    await sp.web.lists.getByTitle(LEAVE_BALANCE_LIST).items.getById(balanceId).update(patch);
  }

  public async getAllLeaves(isHRTeam: boolean, employeeId: number): Promise<ILeaveHistoryItem[]> {
    let query = sp.web.lists
      .getByTitle(LEAVE_HISTORY_LIST)
      .items.select(
        'Id', 'EmpName/Id', 'EmpName/Title', 'StartDate', 'EndDate', 'LeaveType', 'Duration',
        'StartDateInterval', 'EndDateInterval', 'TotalDays', 'Description',
        'isApproved', 'isRejected', 'isCancelled', 'ApproverName', 'ApproverComment',
        'NotifyUsers/Id', 'NotifyUsers/Title'
      )
      .expand('EmpName', 'NotifyUsers')
      .top(5000)
      .orderBy('Id', false);

    if (!isHRTeam) {
      query = query.filter(`EmpName/Id eq ${employeeId}`);
    }
    return (await query.get()) as unknown as ILeaveHistoryItem[];
  }

  public async getPendingApprovalLeaves(isHRTeam: boolean, reportIds: number[]): Promise<ILeaveHistoryItem[]> {
    let filter = 'isApproved eq 0 and isRejected eq 0 and isCancelled eq 0';
    if (!isHRTeam) {
      if (reportIds.length === 0) return [];
      const idFilter = reportIds.map((id) => `EmpName/Id eq ${id}`).join(' or ');
      filter = `(${idFilter}) and ${filter}`;
    }
    const items = await sp.web.lists
      .getByTitle(LEAVE_HISTORY_LIST)
      .items.select('Id', 'EmpName/Id', 'EmpName/Title', 'StartDate', 'EndDate', 'LeaveType', 'TotalDays', 'Description')
      .expand('EmpName')
      .filter(filter)
      .top(5000)
      .orderBy('Id', false)
      .get();
    return items as unknown as ILeaveHistoryItem[];
  }

  public async getLeaveById(itemId: number): Promise<ILeaveHistoryItem | undefined> {
    const items = await sp.web.lists
      .getByTitle(LEAVE_HISTORY_LIST)
      .items.select(
        'Id', 'EmpName/Id', 'StartDate', 'EndDate', 'LeaveType', 'Duration',
        'StartDateInterval', 'EndDateInterval', 'TotalDays', 'Description',
        'NotifyUsers/Id', 'NotifyUsers/Title' , 'NotifyUsers/EMail'
      )
      .expand('EmpName', 'NotifyUsers')
      .filter(`Id eq ${itemId}`)
      .top(1)
      .get();
    if (items.length === 0) return undefined;
    return items[0] as unknown as ILeaveHistoryItem;
  }

  public async createLeave(payload: INewLeavePayload): Promise<{ Id: number }> {
    const result = await sp.web.lists.getByTitle(LEAVE_HISTORY_LIST).items.add(payload);
    return { Id: result.data.Id };
  }

  public async updateLeave(itemId: number, payload: Partial<INewLeavePayload>): Promise<void> {
    await sp.web.lists.getByTitle(LEAVE_HISTORY_LIST).items.getById(itemId).update(payload);
  }

  public async setApproval(itemId: number, approve: boolean, approverName: string, comment: string): Promise<void> {
    const patch: any = { ApproverName: approverName, ApproverComment: comment || null };
    if (approve) {
      patch.isApproved = true;
    } else {
      patch.isRejected = true;
    }
    await sp.web.lists.getByTitle(LEAVE_HISTORY_LIST).items.getById(itemId).update(patch);
  }

  public async cancelLeave(itemId: number): Promise<void> {
    await sp.web.lists.getByTitle(LEAVE_HISTORY_LIST).items.getById(itemId).update({
      isCancelled: true,
      isApproved: false,
      isRejected: false
    });
  }

  public async ensureUsers(loginNames: string[]): Promise<number[]> {
    const ids: number[] = [];
    for (const loginName of loginNames) {
      const user = await sp.web.ensureUser(loginName);
      ids.push(user.data.Id);
    }
    return ids;
  }

  private escape(value: string): string {
    return value.replace(/'/g, "''");
  }
}
