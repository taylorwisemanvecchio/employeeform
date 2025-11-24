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

  // Safe create: add then re-read
  public async createResponse(
    payload: Record<string, unknown>
  ): Promise<IEvaluationResponse> {
    try {
      const addRes = await this.sp.web.lists
        .getByTitle(this.responsesList)
        .items.add(payload);

      // Try to get Id from the response data
      let createdId = addRes?.data?.Id as number | undefined;

      // If no Id in data, try to get it from the item reference
      if (!createdId || typeof createdId !== 'number') {
        try {
          const itemData = await addRes.item.select("Id")();
          createdId = itemData?.Id as number | undefined;
        } catch (itemError) {
          // Ignore error and continue to next attempt
        }
      }

      // If we still don't have an Id, this is a problem
      if (!createdId || typeof createdId !== 'number') {
        throw new Error('Failed to get Id from created response');
      }

      // Re-read the full item to ensure we have complete data
      const fullItem = await this.sp.web.lists
        .getByTitle(this.responsesList)
        .items.getById(createdId)
        .select("Id,Title,AssignmentIDId,ReviewerType,SubmittedDate")();

      const result = fullItem as unknown as IEvaluationResponse;

      // Final validation
      if (!result || typeof result.Id !== 'number') {
        throw new Error('Failed to retrieve created response: Invalid data');
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

    await this.sp.web.lists
      .getByTitle(this.assignmentsList)
      .items.getById(assignmentId)
      .update(payload);
  }
}
