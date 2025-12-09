import { SPFI } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";

export type ReviewerType = "Employee" | "Supervisor" | "Reviewer";

export interface IUserInfo {
  Id: number;
  Email: string;
  Title?: string;
}

export interface IAssignment {
  Id: number;
  Title: string;
  ReviewPeriodStart?: string;
  ReviewPeriodEnd?: string;
  Status?: string;

  SelfEvalSubmitted?: boolean;
  SupervisorSubmitted?: boolean;
  ReviewerSubmitted?: boolean;

  Employee?: IUserInfo;
  Supervisor?: IUserInfo;
  OptionalReviewer?: IUserInfo;
  ProposedReviewer?: IUserInfo;
}

export interface IEvaluationResponse extends Record<string, unknown> {
  Id: number;
  Title?: string;
  AssignmentIDId: number;
  ReviewerType: ReviewerType;
  SubmittedDate?: string;
}

export interface IPendingAssignment extends IAssignment {
  MyRole: ReviewerType;
}

// Internal raw SharePoint shape for assignments
interface IRawAssignment {
  Id: number;
  Title: string;
  ReviewPeriodStart?: string;
  ReviewPeriodEnd?: string;
  Status?: string;

  SelfEvalSubmitted?: boolean;
  SupervisorSubmitted?: boolean;
  ReviewerSubmitted?: boolean;

  Employee?: { Id: number; EMail?: string; Title?: string };
  Supervisor?: { Id: number; EMail?: string; Title?: string };
  OptionalReviewer?: { Id: number; EMail?: string; Title?: string };
  ProposedReviewer?: { Id: number; EMail?: string; Title?: string };
}

export class EvaluationService {
  private assignmentsList = "EvaluationAssignments";
  private responsesList = "EvaluationResponses";

  constructor(private sp: SPFI) {}

  public async getCurrentUser(): Promise<IUserInfo> {
    // Retry logic with exponential backoff for SharePoint context initialization
    const maxRetries = 3;
    let lastError: Error | undefined;

    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        const u = await this.sp.web.currentUser();

        // Validate the user object has required properties
        if (u && typeof u.Id === 'number' && u.Email) {
          return { Id: u.Id, Email: u.Email, Title: u.Title };
        }

        // User object is incomplete, retry after a delay
        if (attempt < maxRetries - 1) {
          await new Promise(resolve => setTimeout(resolve, Math.pow(2, attempt) * 250));
        }
      } catch (error) {
        lastError = error as Error;
        // Wait before retrying
        if (attempt < maxRetries - 1) {
          await new Promise(resolve => setTimeout(resolve, Math.pow(2, attempt) * 250));
        }
      }
    }

    throw new Error(
      'Unable to load current user information. Please refresh the page. ' +
      (lastError ? `Details: ${lastError.message}` : '')
    );
  }

  public async getAssignment(id: number): Promise<IAssignment> {
    const a = await this.sp.web.lists
      .getByTitle(this.assignmentsList)
      .items.getById(id)
      .select(
        "Id,Title,ReviewPeriodStart,ReviewPeriodEnd,Status," +
          "SelfEvalSubmitted,SupervisorSubmitted,ReviewerSubmitted," +
          "Employee/Id,Employee/Title,Employee/EMail," +
          "Supervisor/Id,Supervisor/Title,Supervisor/EMail," +
          "OptionalReviewer/Id,OptionalReviewer/Title,OptionalReviewer/EMail"
      )
      .expand("Employee,Supervisor,OptionalReviewer")();

    const raw = a as unknown as IRawAssignment;

    // Validate that we got a valid assignment with an Id
    if (!raw || typeof raw.Id !== 'number') {
      throw new Error(`Invalid assignment data received for ID ${id}`);
    }

    const mapUser = (
      p?: { Id: number; EMail?: string; Title?: string }
    ): IUserInfo | undefined =>
      p && p.EMail && typeof p.Id === 'number' ? { Id: p.Id, Email: p.EMail, Title: p.Title } : undefined;

    return {
      Id: raw.Id,
      Title: raw.Title,
      ReviewPeriodStart: raw.ReviewPeriodStart,
      ReviewPeriodEnd: raw.ReviewPeriodEnd,
      Status: raw.Status,
      SelfEvalSubmitted: raw.SelfEvalSubmitted,
      SupervisorSubmitted: raw.SupervisorSubmitted,
      ReviewerSubmitted: raw.ReviewerSubmitted,
      Employee: mapUser(raw.Employee),
      Supervisor: mapUser(raw.Supervisor),
      OptionalReviewer: mapUser(raw.OptionalReviewer)
    };
  }

  /**
   * Broad pending fetch:
   * - include items assigned to me in any role
   * - exclude clearly completed/archived items
   * - allow blank Status
   * - fallback to client-side email match if server filter returns none
   */
  public async getPendingAssignmentsForUser(): Promise<IPendingAssignment[]> {
    const me = await this.getCurrentUser();
    const emailLower = me.Email.replace("'", "''").toLowerCase();

    const filter =
      "(" +
        "Employee/EMail eq '" + emailLower + "' or " +
        "Supervisor/EMail eq '" + emailLower + "' or " +
        "OptionalReviewer/EMail eq '" + emailLower + "'" +
      ")" +
      " and (" +
        "Status ne 'Complete' and Status ne 'Completed' and Status ne 'Closed' and Status ne 'Archived' " +
        "or Status eq null" +
      ")";

    const selectExpand =
      "Id,Title,ReviewPeriodStart,ReviewPeriodEnd,Status," +
      "SelfEvalSubmitted,SupervisorSubmitted,ReviewerSubmitted," +
      "Employee/Id,Employee/Title,Employee/EMail," +
      "Supervisor/Id,Supervisor/Title,Supervisor/EMail," +
      "OptionalReviewer/Id,OptionalReviewer/Title,OptionalReviewer/EMail";

    let items = await this.sp.web.lists
      .getByTitle(this.assignmentsList)
      .items.filter(filter)
      .select(selectExpand)
      .expand("Employee,Supervisor,OptionalReviewer")();

    // Fallback if server filter returns nothing
    if (!items || items.length === 0) {
      const base = await this.sp.web.lists
        .getByTitle(this.assignmentsList)
        .items.top(200)();

      const ids = (base as Array<{ Id: number }>)
        .filter((b) => typeof b.Id === 'number')
        .map((b) => b.Id);

      items = await Promise.all(
        ids.map((id: number) =>
          this.sp.web.lists
            .getByTitle(this.assignmentsList)
            .items.getById(id)
            .select(selectExpand)
            .expand("Employee,Supervisor,OptionalReviewer")()
        )
      );
    }

    const rawItems = items as unknown as IRawAssignment[];

    const normalizeUser = (
      p?: { Id: number; EMail?: string; Title?: string }
    ): IUserInfo | undefined =>
      p && p.EMail && typeof p.Id === 'number' ? { Id: p.Id, Email: p.EMail, Title: p.Title } : undefined;

    const meLower = me.Email.toLowerCase();
    const doneStatuses = ["complete", "completed", "closed", "archived"];

    return rawItems
      .filter((it: IRawAssignment) => {
        // First filter: must have a valid Id
        return typeof it.Id === 'number';
      })
      .filter((it: IRawAssignment) => {
        const e = it.Employee && it.Employee.EMail ? it.Employee.EMail.toLowerCase() : undefined;
        const s = it.Supervisor && it.Supervisor.EMail ? it.Supervisor.EMail.toLowerCase() : undefined;
        const r =
          it.OptionalReviewer && it.OptionalReviewer.EMail
            ? it.OptionalReviewer.EMail.toLowerCase()
            : undefined;
        return e === meLower || s === meLower || r === meLower;
      })
      .filter((it: IRawAssignment) => {
        const st = (it.Status || "").toLowerCase();
        return doneStatuses.indexOf(st) === -1; // ES5-safe
      })
      .map((it: IRawAssignment): IPendingAssignment => {
        let role: ReviewerType = "Reviewer";
        if (it.Employee && it.Employee.EMail && it.Employee.EMail.toLowerCase() === meLower) {
          role = "Employee";
        } else if (
          it.Supervisor &&
          it.Supervisor.EMail &&
          it.Supervisor.EMail.toLowerCase() === meLower
        ) {
          role = "Supervisor";
        } else {
          role = "Reviewer";
        }

        return {
          Id: it.Id,
          Title: it.Title,
          ReviewPeriodStart: it.ReviewPeriodStart,
          ReviewPeriodEnd: it.ReviewPeriodEnd,
          Status: it.Status,
          SelfEvalSubmitted: it.SelfEvalSubmitted,
          SupervisorSubmitted: it.SupervisorSubmitted,
          ReviewerSubmitted: it.ReviewerSubmitted,
          Employee: normalizeUser(it.Employee),
          Supervisor: normalizeUser(it.Supervisor),
          OptionalReviewer: normalizeUser(it.OptionalReviewer),
          MyRole: role
        };
      });
  }

  public async getMyResponse(
    assignmentId: number,
    reviewerType: ReviewerType,
    userEmail: string
  ): Promise<IEvaluationResponse | undefined> {
    const email = userEmail.replace("'", "''");
    const items = await this.sp.web.lists
      .getByTitle(this.responsesList)
      .items.filter(
        "AssignmentIDId eq " +
          assignmentId +
          " and ReviewerType eq '" +
          reviewerType +
          "' and ReviewerName/EMail eq '" +
          email +
          "'"
      )
      .select("Id,Title,SubmittedDate,AssignmentIDId,ReviewerType,* ,ReviewerName/EMail")
      .expand("ReviewerName")();

    if (!items || items.length === 0) return undefined;
    return items[0] as unknown as IEvaluationResponse;
  }

  // Safe create: add then query for the created item
  public async createResponse(
    payload: Record<string, unknown>
  ): Promise<IEvaluationResponse> {
    try {
      // Create the item
      await this.sp.web.lists
        .getByTitle(this.responsesList)
        .items.add(payload);

      // Extract the filter criteria from the payload to find the created item
      const assignmentId = payload.AssignmentIDId as number;
      const reviewerType = payload.ReviewerType as string;

      // Get current user to match the ReviewerName
      const me = await this.getCurrentUser();

      // Wait a moment for SharePoint to index the new item
      await new Promise(resolve => setTimeout(resolve, 500));

      // Query for the created item using the same logic as getMyResponse
      const items = await this.sp.web.lists
        .getByTitle(this.responsesList)
        .items.filter(
          "AssignmentIDId eq " +
            assignmentId +
            " and ReviewerType eq '" +
            reviewerType +
            "' and ReviewerName/EMail eq '" +
            me.Email.replace("'", "''") +
            "'"
        )
        .select("Id,Title,AssignmentIDId,ReviewerType,SubmittedDate,*,ReviewerName/Id,ReviewerName/EMail")
        .expand("ReviewerName")
        .top(1)();

      if (!items || items.length === 0) {
        throw new Error('Created response not found after creation');
      }

      const result = items[0] as unknown as IEvaluationResponse;

      // Final validation
      if (!result || typeof result.Id !== 'number') {
        throw new Error('Created response has invalid Id');
      }

      return result;
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : 'Unknown error';
      throw new Error(`Create response failed: ${errorMsg}`);
    }
  }

  public async updateResponse(
    id: number,
    payload: Record<string, unknown>
  ): Promise<void> {
    await this.sp.web.lists
      .getByTitle(this.responsesList)
      .items.getById(id)
      .update(payload);
  }

  public async markSubmitted(
    assignmentId: number,
    role: ReviewerType
  ): Promise<void> {
    const payload: Record<string, unknown> = {};

    if (role === "Employee") payload.SelfEvalSubmitted = true;
    if (role === "Supervisor") payload.SupervisorSubmitted = true;
    if (role === "Reviewer") payload.ReviewerSubmitted = true;

    // Update Status to "In Progress" when at least one evaluation is submitted
    payload.Status = "In Progress";

    await this.sp.web.lists
      .getByTitle(this.assignmentsList)
      .items.getById(assignmentId)
      .update(payload);
  }

  /**
   * Get all assignments where the current user is the supervisor
   */
  public async getAssignmentsWhereSupervisor(): Promise<IAssignment[]> {
    const me = await this.getCurrentUser();
    const emailLower = me.Email.replace("'", "''").toLowerCase();

    const filter = "Supervisor/EMail eq '" + emailLower + "'";

    const selectExpand =
      "Id,Title,ReviewPeriodStart,ReviewPeriodEnd,Status," +
      "SelfEvalSubmitted,SupervisorSubmitted,ReviewerSubmitted," +
      "Employee/Id,Employee/Title,Employee/EMail," +
      "Supervisor/Id,Supervisor/Title,Supervisor/EMail," +
      "OptionalReviewer/Id,OptionalReviewer/Title,OptionalReviewer/EMail," +
      "ProposedReviewer/Id,ProposedReviewer/Title,ProposedReviewer/EMail";

    const items = await this.sp.web.lists
      .getByTitle(this.assignmentsList)
      .items.filter(filter)
      .select(selectExpand)
      .expand("Employee,Supervisor,OptionalReviewer,ProposedReviewer")();

    const rawItems = items as unknown as IRawAssignment[];

    const normalizeUser = (
      p?: { Id: number; EMail?: string; Title?: string }
    ): IUserInfo | undefined =>
      p && p.EMail && typeof p.Id === 'number' ? { Id: p.Id, Email: p.EMail, Title: p.Title } : undefined;

    return rawItems
      .filter((it: IRawAssignment) => {
        return typeof it.Id === 'number';
      })
      .map((it: IRawAssignment): IAssignment => {
        return {
          Id: it.Id,
          Title: it.Title,
          ReviewPeriodStart: it.ReviewPeriodStart,
          ReviewPeriodEnd: it.ReviewPeriodEnd,
          Status: it.Status,
          SelfEvalSubmitted: it.SelfEvalSubmitted,
          SupervisorSubmitted: it.SupervisorSubmitted,
          ReviewerSubmitted: it.ReviewerSubmitted,
          Employee: normalizeUser(it.Employee),
          Supervisor: normalizeUser(it.Supervisor),
          OptionalReviewer: normalizeUser(it.OptionalReviewer),
          ProposedReviewer: normalizeUser(it.ProposedReviewer)
        };
      });
  }

  /**
   * Update or clear the OptionalReviewer field for an assignment
   * @param assignmentId - The assignment ID to update
   * @param reviewerId - The user ID of the optional reviewer, or undefined to clear
   */
  public async updateOptionalReviewer(
    assignmentId: number,
    reviewerId: number | undefined
  ): Promise<void> {
    const payload: Record<string, unknown> = {
      OptionalReviewerId: reviewerId ?? undefined
    };

    await this.sp.web.lists
      .getByTitle(this.assignmentsList)
      .items.getById(assignmentId)
      .update(payload);
  }

  /**
   * Add the proposed reviewer as an optional approver
   * @param assignmentId - The assignment ID to update
   * @param reviewerId - The user ID of the proposed reviewer to add as optional approver
   */
  public async addOptionalApprover(
    assignmentId: number,
    reviewerId: number
  ): Promise<void> {
    const payload: Record<string, unknown> = {
      OptionalReviewerId: reviewerId,
      SendOptionalEmail: true
    };

    await this.sp.web.lists
      .getByTitle(this.assignmentsList)
      .items.getById(assignmentId)
      .update(payload);
  }

  /**
   * Check if the current user is an admin
   * Only specific users can access the admin dashboard
   */
  public async isAdmin(): Promise<boolean> {
    const me = await this.getCurrentUser();
    const adminEmails = [
      "kane@taylorwiseman.com",
      "scimone@taylorwiseman.com",
      "vecchioj@taylorwiseman.com"
    ];
    return adminEmails.some(email => email.toLowerCase() === me.Email.toLowerCase());
  }

  /**
   * Get all evaluation assignments for admin dashboard
   * Returns all assignments regardless of user
   */
  public async getAllAssignments(): Promise<IAssignment[]> {
    const selectExpand =
      "Id,Title,ReviewPeriodStart,ReviewPeriodEnd,Status," +
      "SelfEvalSubmitted,SupervisorSubmitted,ReviewerSubmitted," +
      "Employee/Id,Employee/Title,Employee/EMail," +
      "Supervisor/Id,Supervisor/Title,Supervisor/EMail," +
      "OptionalReviewer/Id,OptionalReviewer/Title,OptionalReviewer/EMail," +
      "ProposedReviewer/Id,ProposedReviewer/Title,ProposedReviewer/EMail";

    const items = await this.sp.web.lists
      .getByTitle(this.assignmentsList)
      .items.select(selectExpand)
      .expand("Employee,Supervisor,OptionalReviewer,ProposedReviewer")
      .top(5000)();

    const rawItems = items as unknown as IRawAssignment[];

    const normalizeUser = (
      p?: { Id: number; EMail?: string; Title?: string }
    ): IUserInfo | undefined =>
      p && p.EMail && typeof p.Id === 'number' ? { Id: p.Id, Email: p.EMail, Title: p.Title } : undefined;

    return rawItems
      .filter((it: IRawAssignment) => {
        return typeof it.Id === 'number';
      })
      .map((it: IRawAssignment): IAssignment => {
        return {
          Id: it.Id,
          Title: it.Title,
          ReviewPeriodStart: it.ReviewPeriodStart,
          ReviewPeriodEnd: it.ReviewPeriodEnd,
          Status: it.Status,
          SelfEvalSubmitted: it.SelfEvalSubmitted,
          SupervisorSubmitted: it.SupervisorSubmitted,
          ReviewerSubmitted: it.ReviewerSubmitted,
          Employee: normalizeUser(it.Employee),
          Supervisor: normalizeUser(it.Supervisor),
          OptionalReviewer: normalizeUser(it.OptionalReviewer),
          ProposedReviewer: normalizeUser(it.ProposedReviewer)
        };
      });
  }

  /**
   * Send reminders webhook to Power Automate
   * @param webhookUrl - The Power Automate webhook URL
   * @param incompleteAssignments - Array of assignments with incomplete evaluations
   */
  public async sendReminders(
    webhookUrl: string,
    incompleteAssignments: IAssignment[]
  ): Promise<void> {
    const payload = {
      assignments: incompleteAssignments.map(a => ({
        id: a.Id,
        title: a.Title,
        ...(a.Employee?.Email && { employee: a.Employee.Email }),
        ...(a.Supervisor?.Email && { supervisor: a.Supervisor.Email }),
        ...(a.OptionalReviewer?.Email && { optionalReviewer: a.OptionalReviewer.Email }),
        ...(a.ProposedReviewer?.Email && { proposedReviewer: a.ProposedReviewer.Email }),
        selfEvalSubmitted: a.SelfEvalSubmitted || false,
        supervisorSubmitted: a.SupervisorSubmitted || false,
        reviewerSubmitted: a.ReviewerSubmitted || false
      }))
    };

    const response = await fetch(webhookUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Webhook request failed with status ${response.status}: ${errorText || response.statusText}`);
    }
  }

  /**
   * Send go live webhook to Power Automate
   * @param webhookUrl - The Power Automate webhook URL
   * @param incompleteAssignments - Array of assignments with incomplete evaluations
   */
  public async sendGoLive(
    webhookUrl: string,
    incompleteAssignments: IAssignment[]
  ): Promise<void> {
    const payload = {
      assignments: incompleteAssignments.map(a => ({
        id: a.Id,
        title: a.Title,
        ...(a.Employee?.Email && { employee: a.Employee.Email }),
        ...(a.Supervisor?.Email && { supervisor: a.Supervisor.Email }),
        ...(a.OptionalReviewer?.Email && { optionalReviewer: a.OptionalReviewer.Email }),
        ...(a.ProposedReviewer?.Email && { proposedReviewer: a.ProposedReviewer.Email }),
        selfEvalSubmitted: a.SelfEvalSubmitted || false,
        supervisorSubmitted: a.SupervisorSubmitted || false,
        reviewerSubmitted: a.ReviewerSubmitted || false
      }))
    };

    const response = await fetch(webhookUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Webhook request failed with status ${response.status}: ${errorText || response.statusText}`);
    }
  }

  /**
   * Send rejection webhook to Power Automate
   * @param webhookUrl - The Power Automate webhook URL
   * @param assignment - The assignment being rejected
   * @param rejectionReason - The reason for rejection
   * @param rejectEmployee - Whether to reject the employee's submission
   * @param rejectSupervisor - Whether to reject the supervisor's submission
   * @param rejectReviewer - Whether to reject the reviewer's submission
   */
  public async sendRejection(
    webhookUrl: string,
    assignment: IAssignment,
    rejectionReason: string,
    rejectEmployee: boolean,
    rejectSupervisor: boolean,
    rejectReviewer: boolean
  ): Promise<void> {
    const payload = {
      assignmentId: assignment.Id,
      assignmentTitle: assignment.Title,
      rejectionReason: rejectionReason,
      rejectedSubmitters: {
        employee: rejectEmployee,
        supervisor: rejectSupervisor,
        reviewer: rejectReviewer
      },
      ...(assignment.Employee?.Email && { employeeEmail: assignment.Employee.Email }),
      ...(assignment.Supervisor?.Email && { supervisorEmail: assignment.Supervisor.Email }),
      ...(assignment.OptionalReviewer?.Email && { reviewerEmail: assignment.OptionalReviewer.Email })
    };

    const response = await fetch(webhookUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Rejection webhook failed with status ${response.status}: ${errorText || response.statusText}`);
    }
  }

  /**
   * Update assignment submission flags after rejection
   * @param assignmentId - The assignment ID to update
   * @param rejectEmployee - Whether to reset employee submission flag
   * @param rejectSupervisor - Whether to reset supervisor submission flag
   * @param rejectReviewer - Whether to reset reviewer submission flag
   */
  public async updateRejectedSubmissions(
    assignmentId: number,
    rejectEmployee: boolean,
    rejectSupervisor: boolean,
    rejectReviewer: boolean
  ): Promise<void> {
    const payload: Record<string, unknown> = {};

    // First, update the submission flags based on what was rejected
    if (rejectEmployee) {
      payload.SelfEvalSubmitted = false;
    }
    if (rejectSupervisor) {
      payload.SupervisorSubmitted = false;
    }
    if (rejectReviewer) {
      payload.ReviewerSubmitted = false;
    }

    // Then update the status to In Progress
    payload.Status = "In Progress";

    await this.sp.web.lists
      .getByTitle(this.assignmentsList)
      .items.getById(assignmentId)
      .update(payload);
  }
}
