import * as React from "react";
import { useState, useEffect, useMemo } from "react";
import { EvaluationService, IAssignment } from "../services/EvaluationService";
import styles from "./EmployeeEvak.module.scss";
import {
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  TextField,
  Checkbox
} from "@fluentui/react";

interface IAdminDashboardProps {
  service: EvaluationService;
}

interface IAdminRow {
  id: number;
  title: string;
  status: string;
  employeeName: string;
  employeeEmail: string;
  employeeSubmitted: boolean;
  supervisorName: string;
  supervisorEmail: string;
  supervisorSubmitted: boolean;
  proposedReviewerName: string;
  proposedReviewerEmail: string;
  proposedReviewerAdded: boolean;
  proposedReviewerSubmitted: boolean;
}

export const AdminDashboard: React.FC<IAdminDashboardProps> = ({ service }) => {
  const [assignments, setAssignments] = useState<IAssignment[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | undefined>();
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState<string>("All");
  const [submissionFilter, setSubmissionFilter] = useState<string>("All");
  const [sendingReminders, setSendingReminders] = useState(false);
  const [webhookUrl, setWebhookUrl] = useState("");
  const [showWebhookInput, setShowWebhookInput] = useState(false);
  const [sendingGoLive, setSendingGoLive] = useState(false);
  const [goLiveWebhookUrl, setGoLiveWebhookUrl] = useState("");
  const [showGoLiveWebhookInput, setShowGoLiveWebhookInput] = useState(false);

  // Rejection dialog state
  const [showRejectDialog, setShowRejectDialog] = useState(false);
  const [selectedAssignment, setSelectedAssignment] = useState<IAssignment | undefined>();
  const [rejectionReason, setRejectionReason] = useState("");
  const [rejectEmployee, setRejectEmployee] = useState(false);
  const [rejectSupervisor, setRejectSupervisor] = useState(false);
  const [rejectReviewer, setRejectReviewer] = useState(false);
  const [sendingRejection, setSendingRejection] = useState(false);
  const rejectionWebhookUrl = "https://defaultdbf39b203d1a468094f8b0aade3398.82.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/35615e9615d745828af325e55871ab7b/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=65ifFfq4dvZMXpSv06DXM9Z3cYUfWWulFRwvLCTRViA";

  const loadAssignments = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      setError(undefined);
      const data = await service.getAllAssignments();
      setAssignments(data);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : "Failed to load assignments";
      setError(errorMessage);
    } finally {
      setLoading(false);
    }
  }, [service]);

  useEffect(() => {
    loadAssignments().catch(() => {
      // Error is handled in loadAssignments
    });
  }, [loadAssignments]);

  // Transform assignments into admin rows
  const adminRows: IAdminRow[] = useMemo(() => {
    return assignments.map(a => {
      const proposedReviewerAdded =
        a.ProposedReviewer && a.OptionalReviewer &&
        a.ProposedReviewer.Id === a.OptionalReviewer.Id;

      return {
        id: a.Id,
        title: a.Title,
        status: a.Status || "Not Started",
        employeeName: a.Employee?.Title || "N/A",
        employeeEmail: a.Employee?.Email || "",
        employeeSubmitted: a.SelfEvalSubmitted || false,
        supervisorName: a.Supervisor?.Title || "N/A",
        supervisorEmail: a.Supervisor?.Email || "",
        supervisorSubmitted: a.SupervisorSubmitted || false,
        proposedReviewerName: a.ProposedReviewer?.Title || "N/A",
        proposedReviewerEmail: a.ProposedReviewer?.Email || "",
        proposedReviewerAdded: proposedReviewerAdded || false,
        proposedReviewerSubmitted: a.ReviewerSubmitted || false
      };
    });
  }, [assignments]);

  // Filter rows based on search and filters
  const filteredRows = useMemo(() => {
    return adminRows.filter(row => {
      // Search filter - ES5 compatible
      const searchLower = searchTerm.toLowerCase();
      const matchesSearch =
        row.title.toLowerCase().indexOf(searchLower) !== -1 ||
        row.employeeName.toLowerCase().indexOf(searchLower) !== -1 ||
        row.employeeEmail.toLowerCase().indexOf(searchLower) !== -1 ||
        row.supervisorName.toLowerCase().indexOf(searchLower) !== -1 ||
        row.supervisorEmail.toLowerCase().indexOf(searchLower) !== -1 ||
        row.proposedReviewerName.toLowerCase().indexOf(searchLower) !== -1 ||
        row.proposedReviewerEmail.toLowerCase().indexOf(searchLower) !== -1;

      if (!matchesSearch) return false;

      // Status filter
      if (statusFilter !== "All" && row.status !== statusFilter) return false;

      // Submission filter
      if (submissionFilter === "All Completed") {
        return row.employeeSubmitted && row.supervisorSubmitted &&
               (!row.proposedReviewerAdded || row.proposedReviewerSubmitted);
      } else if (submissionFilter === "Has Incomplete") {
        return !row.employeeSubmitted || !row.supervisorSubmitted ||
               (row.proposedReviewerAdded && !row.proposedReviewerSubmitted);
      }

      return true;
    });
  }, [adminRows, searchTerm, statusFilter, submissionFilter]);

  // Calculate statistics
  const stats = useMemo(() => {
    const total = adminRows.length;
    const inProgress = adminRows.filter(r => r.status === "In Progress").length;
    const completed = adminRows.filter(r =>
      r.employeeSubmitted && r.supervisorSubmitted &&
      (!r.proposedReviewerAdded || r.proposedReviewerSubmitted)
    ).length;
    const employeeCompleted = adminRows.filter(r => r.employeeSubmitted).length;
    const supervisorCompleted = adminRows.filter(r => r.supervisorSubmitted).length;
    const reviewerCompleted = adminRows.filter(r =>
      r.proposedReviewerAdded && r.proposedReviewerSubmitted
    ).length;
    const reviewerPending = adminRows.filter(r =>
      r.proposedReviewerAdded && !r.proposedReviewerSubmitted
    ).length;

    return {
      total,
      inProgress,
      completed,
      completionRate: total > 0 ? Math.round((completed / total) * 100) : 0,
      employeeCompleted,
      employeeRate: total > 0 ? Math.round((employeeCompleted / total) * 100) : 0,
      supervisorCompleted,
      supervisorRate: total > 0 ? Math.round((supervisorCompleted / total) * 100) : 0,
      reviewerCompleted,
      reviewerPending,
      reviewerRate: reviewerPending + reviewerCompleted > 0
        ? Math.round((reviewerCompleted / (reviewerPending + reviewerCompleted)) * 100)
        : 0
    };
  }, [adminRows]);

  // Get unique status values for filter dropdown - ES5 compatible
  const uniqueStatuses = useMemo(() => {
    const statusMap: { [key: string]: boolean } = {};
    adminRows.forEach(r => {
      statusMap[r.status] = true;
    });
    const statuses: string[] = [];
    for (const status in statusMap) {
      if (Object.prototype.hasOwnProperty.call(statusMap, status)) {
        statuses.push(status);
      }
    }
    return statuses.sort();
  }, [adminRows]);

  // Get incomplete assignments for reminders
  const incompleteAssignments = useMemo(() => {
    return assignments.filter(a =>
      !a.SelfEvalSubmitted ||
      !a.SupervisorSubmitted ||
      (a.ProposedReviewer && a.OptionalReviewer &&
       a.ProposedReviewer.Id === a.OptionalReviewer.Id && !a.ReviewerSubmitted)
    );
  }, [assignments]);

  const handleSendReminders = async (): Promise<void> => {
    if (!webhookUrl) {
      setShowWebhookInput(true);
      return;
    }

    try {
      setSendingReminders(true);
      await service.sendReminders(webhookUrl, incompleteAssignments);
      alert(`Reminders sent successfully to ${incompleteAssignments.length} incomplete evaluations!`);
      setShowWebhookInput(false);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : "Failed to send reminders";
      alert(`Error sending reminders: ${errorMessage}`);
    } finally {
      setSendingReminders(false);
    }
  };

  const handleSendGoLive = async (): Promise<void> => {
    if (!goLiveWebhookUrl) {
      setShowGoLiveWebhookInput(true);
      return;
    }

    try {
      setSendingGoLive(true);
      await service.sendGoLive(goLiveWebhookUrl, incompleteAssignments);
      alert(`Go Live sent successfully to ${incompleteAssignments.length} incomplete evaluations!`);
      setShowGoLiveWebhookInput(false);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : "Failed to send go live";
      alert(`Error sending go live: ${errorMessage}`);
    } finally {
      setSendingGoLive(false);
    }
  };

  const handleOpenRejectDialog = (assignment: IAssignment): void => {
    setSelectedAssignment(assignment);
    setRejectionReason("");
    setRejectEmployee(false);
    setRejectSupervisor(false);
    setRejectReviewer(false);
    setShowRejectDialog(true);
  };

  const handleCloseRejectDialog = (): void => {
    setShowRejectDialog(false);
    setSelectedAssignment(undefined);
    setRejectionReason("");
    setRejectEmployee(false);
    setRejectSupervisor(false);
    setRejectReviewer(false);
  };

  const handleSendRejection = async (): Promise<void> => {
    if (!selectedAssignment) return;

    if (!rejectionReason.trim()) {
      alert("Please provide a reason for rejection");
      return;
    }

    if (!rejectEmployee && !rejectSupervisor && !rejectReviewer) {
      alert("Please select at least one submitter to reject");
      return;
    }

    try {
      setSendingRejection(true);

      // Send rejection webhook to Power Automate
      await service.sendRejection(
        rejectionWebhookUrl,
        selectedAssignment,
        rejectionReason,
        rejectEmployee,
        rejectSupervisor,
        rejectReviewer
      );

      // Update submission flags and status in SharePoint
      await service.updateRejectedSubmissions(
        selectedAssignment.Id,
        rejectEmployee,
        rejectSupervisor,
        rejectReviewer
      );

      alert("Rejection sent successfully!");
      handleCloseRejectDialog();

      // Reload assignments to reflect the changes
      await loadAssignments();
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : "Failed to send rejection";
      alert(`Error sending rejection: ${errorMessage}`);
    } finally {
      setSendingRejection(false);
    }
  };

  if (loading) {
    return <div className={styles.container}>Loading admin dashboard...</div>;
  }

  if (error) {
    return (
      <div className={styles.container}>
        <div className={styles.error}>Error: {error}</div>
        <button onClick={loadAssignments} className={styles.button}>
          Retry
        </button>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <h2 className={styles.title}>Admin Dashboard - Evaluation Progress</h2>

      {/* Statistics Cards */}
      <div className={styles.statsGrid}>
        <div className={styles.statCard}>
          <h3>Total Evaluations</h3>
          <div className={styles.statValue}>{stats.total}</div>
        </div>
        <div className={styles.statCard}>
          <h3>In Progress</h3>
          <div className={styles.statValue}>{stats.inProgress}</div>
        </div>
        <div className={styles.statCard}>
          <h3>Fully Completed</h3>
          <div className={styles.statValue}>{stats.completed}</div>
          <div className={styles.statSubtext}>{stats.completionRate}%</div>
        </div>
        <div className={styles.statCard}>
          <h3>Employee Submissions</h3>
          <div className={styles.statValue}>{stats.employeeCompleted}/{stats.total}</div>
          <div className={styles.statSubtext}>{stats.employeeRate}%</div>
        </div>
        <div className={styles.statCard}>
          <h3>Supervisor Submissions</h3>
          <div className={styles.statValue}>{stats.supervisorCompleted}/{stats.total}</div>
          <div className={styles.statSubtext}>{stats.supervisorRate}%</div>
        </div>
        <div className={styles.statCard}>
          <h3>Reviewer Submissions</h3>
          <div className={styles.statValue}>{stats.reviewerCompleted}/{stats.reviewerPending + stats.reviewerCompleted}</div>
          <div className={styles.statSubtext}>{stats.reviewerRate}%</div>
        </div>
      </div>

      {/* Progress Bars */}
      <div className={styles.progressSection}>
        <h3>Completion Progress</h3>
        <div className={styles.progressBar}>
          <div className={styles.progressLabel}>
            <span>Overall Completion</span>
            <span>{stats.completionRate}%</span>
          </div>
          <div className={styles.progressTrack}>
            <div
              className={styles.progressFill}
              style={{ width: `${stats.completionRate}%` }}
            />
          </div>
        </div>
        <div className={styles.progressBar}>
          <div className={styles.progressLabel}>
            <span>Employee Evaluations</span>
            <span>{stats.employeeRate}%</span>
          </div>
          <div className={styles.progressTrack}>
            <div
              className={styles.progressFill}
              style={{ width: `${stats.employeeRate}%`, backgroundColor: '#0078d4' }}
            />
          </div>
        </div>
        <div className={styles.progressBar}>
          <div className={styles.progressLabel}>
            <span>Supervisor Evaluations</span>
            <span>{stats.supervisorRate}%</span>
          </div>
          <div className={styles.progressTrack}>
            <div
              className={styles.progressFill}
              style={{ width: `${stats.supervisorRate}%`, backgroundColor: '#107c10' }}
            />
          </div>
        </div>
        {(stats.reviewerPending + stats.reviewerCompleted) > 0 && (
          <div className={styles.progressBar}>
            <div className={styles.progressLabel}>
              <span>Reviewer Evaluations</span>
              <span>{stats.reviewerRate}%</span>
            </div>
            <div className={styles.progressTrack}>
              <div
                className={styles.progressFill}
                style={{ width: `${stats.reviewerRate}%`, backgroundColor: '#8764b8' }}
              />
            </div>
          </div>
        )}
      </div>

      {/* Send Reminders Section */}
      <div className={styles.reminderSection}>
        <button
          onClick={handleSendReminders}
          disabled={sendingReminders || incompleteAssignments.length === 0}
          className={styles.reminderButton}
        >
          {sendingReminders ? "Sending..." : `Send Reminders (${incompleteAssignments.length} incomplete)`}
        </button>
        {showWebhookInput && (
          <div className={styles.webhookInput}>
            <input
              type="text"
              placeholder="Enter Power Automate webhook URL"
              value={webhookUrl}
              onChange={(e) => setWebhookUrl(e.target.value)}
              className={styles.input}
            />
            <button onClick={handleSendReminders} className={styles.button}>
              Send
            </button>
            <button onClick={() => setShowWebhookInput(false)} className={styles.buttonSecondary}>
              Cancel
            </button>
          </div>
        )}
      </div>

      {/* Go Live Section */}
      <div className={styles.reminderSection}>
        <button
          onClick={handleSendGoLive}
          disabled={sendingGoLive || incompleteAssignments.length === 0}
          className={styles.reminderButton}
        >
          {sendingGoLive ? "Sending..." : `Go Live (${incompleteAssignments.length} incomplete)`}
        </button>
        {showGoLiveWebhookInput && (
          <div className={styles.webhookInput}>
            <input
              type="text"
              placeholder="Enter Power Automate webhook URL for Go Live"
              value={goLiveWebhookUrl}
              onChange={(e) => setGoLiveWebhookUrl(e.target.value)}
              className={styles.input}
            />
            <button onClick={handleSendGoLive} className={styles.button}>
              Send
            </button>
            <button onClick={() => setShowGoLiveWebhookInput(false)} className={styles.buttonSecondary}>
              Cancel
            </button>
          </div>
        )}
      </div>

      {/* Search and Filters */}
      <div className={styles.filterSection}>
        <input
          type="text"
          placeholder="Search by title, employee, supervisor, or reviewer..."
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          className={styles.searchInput}
        />
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
          className={styles.filterSelect}
        >
          <option value="All">All Statuses</option>
          {uniqueStatuses.map((status: string) => (
            <option key={status} value={status}>{status}</option>
          ))}
        </select>
        <select
          value={submissionFilter}
          onChange={(e) => setSubmissionFilter(e.target.value)}
          className={styles.filterSelect}
        >
          <option value="All">All Submissions</option>
          <option value="All Completed">All Completed</option>
          <option value="Has Incomplete">Has Incomplete</option>
        </select>
      </div>

      {/* Results Count */}
      <div className={styles.resultsCount}>
        Showing {filteredRows.length} of {adminRows.length} evaluations
      </div>

      {/* Data Table */}
      <div className={styles.tableContainer}>
        <table className={styles.adminTable}>
          <thead>
            <tr>
              <th>Title</th>
              <th>Status</th>
              <th>Employee</th>
              <th>Employee Submitted</th>
              <th>Supervisor</th>
              <th>Supervisor Submitted</th>
              <th>Proposed Reviewer</th>
              <th>Reviewer Added</th>
              <th>Reviewer Submitted</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {filteredRows.length === 0 ? (
              <tr>
                <td colSpan={10} className={styles.noResults}>
                  No evaluations found matching your criteria
                </td>
              </tr>
            ) : (
              filteredRows.map(row => {
                // ES5-compatible find alternative
                let assignment: IAssignment | undefined;
                for (let i = 0; i < assignments.length; i++) {
                  if (assignments[i].Id === row.id) {
                    assignment = assignments[i];
                    break;
                  }
                }
                return (
                  <tr key={row.id}>
                    <td>{row.title}</td>
                    <td>
                      <span className={styles.statusBadge}>
                        {row.status}
                      </span>
                    </td>
                    <td title={row.employeeEmail}>{row.employeeName}</td>
                    <td className={styles.checkboxCell}>
                      <input
                        type="checkbox"
                        checked={row.employeeSubmitted}
                        disabled
                        readOnly
                      />
                    </td>
                    <td title={row.supervisorEmail}>{row.supervisorName}</td>
                    <td className={styles.checkboxCell}>
                      <input
                        type="checkbox"
                        checked={row.supervisorSubmitted}
                        disabled
                        readOnly
                      />
                    </td>
                    <td title={row.proposedReviewerEmail}>{row.proposedReviewerName}</td>
                    <td className={styles.checkboxCell}>
                      <input
                        type="checkbox"
                        checked={row.proposedReviewerAdded}
                        disabled
                        readOnly
                      />
                    </td>
                    <td className={styles.checkboxCell}>
                      <input
                        type="checkbox"
                        checked={row.proposedReviewerSubmitted}
                        disabled
                        readOnly
                      />
                    </td>
                    <td>
                      <button
                        onClick={() => assignment && handleOpenRejectDialog(assignment)}
                        className={styles.button}
                        disabled={!assignment}
                        style={{
                          backgroundColor: '#d13438',
                          color: 'white',
                          border: 'none',
                          padding: '6px 12px',
                          borderRadius: '4px',
                          cursor: assignment ? 'pointer' : 'not-allowed'
                        }}
                      >
                        Reject
                      </button>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* Rejection Dialog */}
      <Dialog
        hidden={!showRejectDialog}
        onDismiss={handleCloseRejectDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Reject Evaluation',
          subText: selectedAssignment ? `Assignment: ${selectedAssignment.Title}` : ''
        }}
        modalProps={{
          isBlocking: true
        }}
      >
        <TextField
          label="Reason for Rejection"
          multiline
          rows={4}
          required
          value={rejectionReason}
          onChange={(_, newValue) => setRejectionReason(newValue || "")}
          placeholder="Please provide a detailed reason for the rejection..."
        />

        <div style={{ marginTop: 16 }}>
          <label style={{ fontWeight: 600, display: 'block', marginBottom: 8 }}>
            Select submitters to reject:
          </label>
          {selectedAssignment?.Employee && (
            <Checkbox
              label={`Employee: ${selectedAssignment.Employee.Title || selectedAssignment.Employee.Email}`}
              checked={rejectEmployee}
              onChange={(_, checked) => setRejectEmployee(checked || false)}
              disabled={!selectedAssignment.SelfEvalSubmitted}
            />
          )}
          {selectedAssignment?.Supervisor && (
            <Checkbox
              label={`Supervisor: ${selectedAssignment.Supervisor.Title || selectedAssignment.Supervisor.Email}`}
              checked={rejectSupervisor}
              onChange={(_, checked) => setRejectSupervisor(checked || false)}
              disabled={!selectedAssignment.SupervisorSubmitted}
            />
          )}
          {selectedAssignment?.OptionalReviewer && (
            <Checkbox
              label={`Reviewer: ${selectedAssignment.OptionalReviewer.Title || selectedAssignment.OptionalReviewer.Email}`}
              checked={rejectReviewer}
              onChange={(_, checked) => setRejectReviewer(checked || false)}
              disabled={!selectedAssignment.ReviewerSubmitted}
            />
          )}
        </div>

        <DialogFooter>
          <PrimaryButton
            onClick={handleSendRejection}
            text="Send Rejection"
            disabled={sendingRejection}
          />
          <DefaultButton
            onClick={handleCloseRejectDialog}
            text="Cancel"
            disabled={sendingRejection}
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
