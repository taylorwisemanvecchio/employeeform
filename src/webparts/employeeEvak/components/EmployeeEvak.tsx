import * as React from "react";
import styles from "./EmployeeEvak.module.scss";
import type { IEmployeeEvakProps } from "./IEmployeeEvakProps";

import { EvaluationService, IPendingAssignment } from "../services/EvaluationService";
import EvaluationForm from "./EvaluationForm";
import SupervisorDashboard from "./SupervisorDashboard";
import { PrimaryButton, DefaultButton, Stack } from "@fluentui/react";

// Hoisted function avoids no-use-before-define
function PendingAssignmentsDashboard(props: { sp: IEmployeeEvakProps["sp"] }): React.ReactElement {
  const { sp } = props;
  const svc = React.useMemo((): EvaluationService => new EvaluationService(sp), [sp]);

  const [pending, setPending] = React.useState<IPendingAssignment[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [selected, setSelected] = React.useState<IPendingAssignment | undefined>(undefined);
  const [isMobile, setIsMobile] = React.useState<boolean>(false);

  // Detect mobile viewport
  React.useEffect((): (() => void) => {
    const checkMobile = (): void => {
      setIsMobile(window.innerWidth <= 768);
    };

    checkMobile();
    window.addEventListener('resize', checkMobile);
    return (): void => window.removeEventListener('resize', checkMobile);
  }, []);

  React.useEffect((): void => {
    (async (): Promise<void> => {
      try {
        setLoading(true);
        setError(undefined);
        const items = await svc.getPendingAssignmentsForUser();
        setPending(items);
      } catch (e: unknown) {
        const msg = e instanceof Error ? e.message : "Failed to load pending assignments.";
        setError(msg);
      } finally {
        setLoading(false);
      }
    })().catch((): void => {});
  }, [svc]);

  if (loading) {
    return <div style={{ padding: isMobile ? 8 : 16 }}>Loading pending evaluations…</div>;
  }

  if (error) {
    return (
      <div style={{ padding: isMobile ? 8 : 16 }}>
        <h3>Something went wrong</h3>
        <p>{error}</p>
      </div>
    );
  }

  if (selected) {
    // Validate selected assignment has a valid Id
    if (!selected.Id || typeof selected.Id !== 'number') {
      return (
        <div style={{ padding: isMobile ? 8 : 16 }}>
          <h3>Something went wrong</h3>
          <p>Invalid assignment selected. Please go back and try again.</p>
          <DefaultButton
            text={isMobile ? "Back" : "Back to Pending List"}
            onClick={(): void => setSelected(undefined)}
            styles={{ root: { width: isMobile ? "100%" : "auto" } }}
          />
        </div>
      );
    }

    return (
      <div style={{ padding: isMobile ? 8 : 16 }}>
        <Stack
          horizontal={!isMobile}
          tokens={{ childrenGap: 8 }}
          style={{ marginBottom: 12 }}
        >
          <DefaultButton
            text={isMobile ? "Back" : "Back to Pending List"}
            onClick={(): void => setSelected(undefined)}
            styles={{ root: { width: isMobile ? "100%" : "auto" } }}
          />
        </Stack>

        <h2 style={{ marginTop: 0, fontSize: isMobile ? "1.2em" : "1.5em" }}>{selected.Title}</h2>
        <div style={{ marginBottom: 12, color: "#555", fontSize: isMobile ? "0.9em" : "1em" }}>
          Role: <strong>{selected.MyRole}</strong>
        </div>

        {/* ✅ removed onSubmitted prop to match component signature */}
        <EvaluationForm
          sp={sp}
          assignmentId={selected.Id}
          reviewerType={selected.MyRole}
          key={`${selected.Id}|${selected.MyRole}`}
        />
      </div>
    );
  }

  return (
    <div style={{ padding: isMobile ? 8 : 16 }}>
      <h2 style={{ marginTop: 0, fontSize: isMobile ? "1.2em" : "1.5em" }}>Pending Evaluations</h2>

      {pending.length === 0 && (
        <div style={{ fontSize: isMobile ? "0.9em" : "1em" }}>
          No pending evaluations assigned to you.
        </div>
      )}

      <Stack tokens={{ childrenGap: 10 }}>
        {pending
          .filter((a: IPendingAssignment) => typeof a.Id === 'number')
          .map((a: IPendingAssignment) => (
            <div
              key={`${a.Id}|${a.MyRole}`}
              style={{
                border: "1px solid #eee",
                borderRadius: 8,
                padding: isMobile ? 10 : 12,
                display: "flex",
                flexDirection: isMobile ? "column" : "row",
                justifyContent: "space-between",
                alignItems: isMobile ? "stretch" : "center",
                gap: isMobile ? 10 : 0
              }}
            >
              <div style={{ flex: 1 }}>
                <div style={{
                  fontWeight: 600,
                  fontSize: isMobile ? "0.95em" : "1em",
                  marginBottom: isMobile ? 4 : 0
                }}>
                  {a.Title}
                </div>
                <div style={{ fontSize: isMobile ? 11 : 12, color: "#666" }}>
                  Role: {a.MyRole}
                </div>
                {(a.ReviewPeriodStart || a.ReviewPeriodEnd) && (
                  <div style={{ fontSize: isMobile ? 11 : 12, color: "#666" }}>
                    Period: {a.ReviewPeriodStart ?? "—"} to {a.ReviewPeriodEnd ?? "—"}
                  </div>
                )}
              </div>

              <PrimaryButton
                text="Open"
                onClick={(): void => setSelected(a)}
                styles={{
                  root: {
                    width: isMobile ? "100%" : "auto",
                    minWidth: isMobile ? "auto" : 80
                  }
                }}
              />
            </div>
          ))}
      </Stack>
    </div>
  );
}

type ViewType = "pending" | "supervisor";

function MainApp(props: { sp: IEmployeeEvakProps["sp"] }): React.ReactElement {
  const { sp } = props;
  const svc = React.useMemo((): EvaluationService => new EvaluationService(sp), [sp]);

  const [currentView, setCurrentView] = React.useState<ViewType>("pending");
  const [isSupervisor, setIsSupervisor] = React.useState<boolean>(false);
  const [checkingSupervisor, setCheckingSupervisor] = React.useState<boolean>(true);
  const [isMobile, setIsMobile] = React.useState<boolean>(false);

  // Detect mobile viewport
  React.useEffect((): (() => void) => {
    const checkMobile = (): void => {
      setIsMobile(window.innerWidth <= 768);
    };

    checkMobile();
    window.addEventListener('resize', checkMobile);
    return (): void => window.removeEventListener('resize', checkMobile);
  }, []);

  // Check if current user is a supervisor
  React.useEffect((): void => {
    (async (): Promise<void> => {
      try {
        const assignments = await svc.getAssignmentsWhereSupervisor();
        setIsSupervisor(assignments.length > 0);
      } catch {
        // If error checking, assume not a supervisor
        setIsSupervisor(false);
      } finally {
        setCheckingSupervisor(false);
      }
    })().catch((): void => {});
  }, [svc]);

  if (checkingSupervisor) {
    return <div style={{ padding: isMobile ? 8 : 16 }}>Loading…</div>;
  }

  return (
    <div>
      {/* Navigation tabs - only show if user is a supervisor */}
      {isSupervisor && (
        <div
          style={{
            backgroundColor: "#f5f5f5",
            borderBottom: "2px solid #0b6a53",
            padding: isMobile ? "8px 8px 0" : "12px 16px 0"
          }}
        >
          <Stack
            horizontal={!isMobile}
            tokens={{ childrenGap: isMobile ? 8 : 12 }}
            style={{ marginBottom: isMobile ? 8 : 0 }}
          >
            <button
              onClick={(): void => setCurrentView("pending")}
              style={{
                padding: isMobile ? "8px 12px" : "10px 16px",
                backgroundColor: currentView === "pending" ? "#0b6a53" : "transparent",
                color: currentView === "pending" ? "#fff" : "#333",
                border: "none",
                borderRadius: "4px 4px 0 0",
                cursor: "pointer",
                fontSize: isMobile ? "0.9em" : "1em",
                fontWeight: currentView === "pending" ? 600 : 400,
                width: isMobile ? "100%" : "auto"
              }}
            >
              My Evaluations
            </button>
            <button
              onClick={(): void => setCurrentView("supervisor")}
              style={{
                padding: isMobile ? "8px 12px" : "10px 16px",
                backgroundColor: currentView === "supervisor" ? "#0b6a53" : "transparent",
                color: currentView === "supervisor" ? "#fff" : "#333",
                border: "none",
                borderRadius: "4px 4px 0 0",
                cursor: "pointer",
                fontSize: isMobile ? "0.9em" : "1em",
                fontWeight: currentView === "supervisor" ? 600 : 400,
                width: isMobile ? "100%" : "auto"
              }}
            >
              Supervisor Dashboard
            </button>
          </Stack>
        </div>
      )}

      {/* Main content */}
      {currentView === "pending" ? (
        <PendingAssignmentsDashboard sp={sp} />
      ) : (
        <SupervisorDashboard sp={sp} />
      )}
    </div>
  );
}

export default class EmployeeEvak extends React.Component<IEmployeeEvakProps> {
  public render(): React.ReactElement<IEmployeeEvakProps> {
    const { sp, hasTeamsContext } = this.props;

    return (
      <section className={`${styles.employeeEvak} ${hasTeamsContext ? styles.teams : ""}`}>
        <MainApp sp={sp} />
      </section>
    );
  }
}
