import * as React from "react";
import { EvaluationService, IAssignment } from "../services/EvaluationService";
import { PrimaryButton, Stack } from "@fluentui/react";
import type { IEmployeeEvakProps } from "./IEmployeeEvakProps";

interface ISupervisorDashboardProps {
  sp: IEmployeeEvakProps["sp"];
}

export default function SupervisorDashboard(props: ISupervisorDashboardProps): React.ReactElement {
  const { sp } = props;
  const svc = React.useMemo((): EvaluationService => new EvaluationService(sp), [sp]);

  const [assignments, setAssignments] = React.useState<IAssignment[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [isMobile, setIsMobile] = React.useState<boolean>(false);
  const [updatingId, setUpdatingId] = React.useState<number | undefined>(undefined);

  // Detect mobile viewport
  React.useEffect((): (() => void) => {
    const checkMobile = (): void => {
      setIsMobile(window.innerWidth <= 768);
    };

    checkMobile();
    window.addEventListener('resize', checkMobile);
    return (): void => window.removeEventListener('resize', checkMobile);
  }, []);

  // Load assignments
  const loadAssignments = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      setError(undefined);
      const items = await svc.getAssignmentsWhereSupervisor();
      setAssignments(items);
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : "Failed to load supervisor assignments.";
      setError(msg);
    } finally {
      setLoading(false);
    }
  }, [svc]);

  React.useEffect((): void => {
    loadAssignments().catch((): void => {});
  }, [loadAssignments]);

  const handleAddOptionalApprover = async (
    assignmentId: number,
    proposedReviewerId: number
  ): Promise<void> => {
    try {
      setUpdatingId(assignmentId);

      await svc.addOptionalApprover(assignmentId, proposedReviewerId);
      // Reload assignments to reflect the change
      await loadAssignments();
      alert("Optional approver added successfully.");
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : "Failed to add optional approver.";
      alert(`Error: ${msg}`);
    } finally {
      setUpdatingId(undefined);
    }
  };

  if (loading) {
    return <div style={{ padding: isMobile ? 8 : 16 }}>Loading supervisor dashboard…</div>;
  }

  if (error) {
    return (
      <div style={{ padding: isMobile ? 8 : 16 }}>
        <h3>Something went wrong</h3>
        <p>{error}</p>
      </div>
    );
  }

  if (assignments.length === 0) {
    return (
      <div style={{ padding: isMobile ? 8 : 16 }}>
        <h2 style={{ marginTop: 0, fontSize: isMobile ? "1.2em" : "1.5em" }}>Supervisor Dashboard</h2>
        <div style={{ fontSize: isMobile ? "0.9em" : "1em" }}>
          You are not listed as a supervisor for any employees.
        </div>
      </div>
    );
  }

  return (
    <div style={{ padding: isMobile ? 8 : 16 }}>
      <h2 style={{ marginTop: 0, fontSize: isMobile ? "1.2em" : "1.5em" }}>
        Supervisor Dashboard
      </h2>
      <p style={{ color: "#555", fontSize: isMobile ? "0.9em" : "1em" }}>
        Review proposed optional reviewers for your employees
      </p>

      <Stack tokens={{ childrenGap: 12 }}>
        {assignments
          .filter((a: IAssignment) => typeof a.Id === 'number')
          .filter((a: IAssignment) => {
            // Only show assignments where ProposedReviewer hasn't been approved yet
            // Hide if OptionalReviewer.Id equals ProposedReviewer.Id
            const proposedId = a.ProposedReviewer?.Id;
            const optionalId = a.OptionalReviewer?.Id;

            // Show if there's a ProposedReviewer and it's different from OptionalReviewer
            return proposedId && proposedId !== optionalId;
          })
          .map((a: IAssignment) => {
            const hasProposedReviewer = a.ProposedReviewer && a.ProposedReviewer.Id;
            const isUpdating = updatingId === a.Id;

            return (
              <div
                key={a.Id}
                style={{
                  border: "1px solid #eee",
                  borderRadius: 8,
                  padding: isMobile ? 10 : 16,
                  backgroundColor: "#fafafa"
                }}
              >
                {/* Assignment Header */}
                <div style={{
                  marginBottom: 12,
                  paddingBottom: 12,
                  borderBottom: "1px solid #eee"
                }}>
                  <div style={{
                    fontWeight: 600,
                    fontSize: isMobile ? "1em" : "1.1em",
                    marginBottom: 4
                  }}>
                    {a.Title}
                  </div>
                  <div style={{ fontSize: isMobile ? 12 : 13, color: "#666" }}>
                    Employee: {a.Employee?.Title || "Unknown"}
                  </div>
                  {(a.ReviewPeriodStart || a.ReviewPeriodEnd) && (
                    <div style={{ fontSize: isMobile ? 11 : 12, color: "#666" }}>
                      Period: {a.ReviewPeriodStart ?? "—"} to {a.ReviewPeriodEnd ?? "—"}
                    </div>
                  )}
                </div>

                {/* Proposed Reviewer Section */}
                <div
                  style={{
                    display: "flex",
                    flexDirection: isMobile ? "column" : "row",
                    alignItems: isMobile ? "stretch" : "center",
                    gap: isMobile ? 10 : 16,
                    backgroundColor: "#fff",
                    padding: isMobile ? 10 : 12,
                    borderRadius: 6,
                    border: "1px solid #e0e0e0"
                  }}
                >
                  <div style={{ flex: 1 }}>
                    <div style={{
                      fontSize: isMobile ? 12 : 13,
                      fontWeight: 600,
                      color: "#0b6a53",
                      marginBottom: 4
                    }}>
                      Proposed Reviewer
                    </div>
                    <div style={{ fontSize: isMobile ? 13 : 14 }}>
                      {hasProposedReviewer ? (
                        <span style={{ color: "#333" }}>
                          {a.ProposedReviewer?.Title || "Unknown"}
                        </span>
                      ) : (
                        <span style={{ color: "#999", fontStyle: "italic" }}>
                          No proposed reviewer
                        </span>
                      )}
                    </div>
                  </div>

                  {/* Action Button - Only show if there's a proposed reviewer */}
                  {hasProposedReviewer && (
                    <PrimaryButton
                      text="Add Additional Reviewer"
                      disabled={isUpdating}
                      onClick={(): void => {
                        if (confirm(`Add ${a.ProposedReviewer?.Title} as optional approver?`)) {
                          handleAddOptionalApprover(
                            a.Id,
                            a.ProposedReviewer!.Id
                          ).catch((): void => {});
                        }
                      }}
                      styles={{
                        root: {
                          width: isMobile ? "100%" : "auto",
                          minWidth: 160,
                          backgroundColor: "#0b6a53",
                          borderColor: "#0b6a53"
                        },
                        rootHovered: {
                          backgroundColor: "#095847",
                          borderColor: "#095847"
                        }
                      }}
                    />
                  )}
                </div>
              </div>
            );
          })}
      </Stack>
    </div>
  );
}
