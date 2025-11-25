import * as React from "react";
import styles from "./EmployeeEvak.module.scss";
import type { IEmployeeEvakProps } from "./IEmployeeEvakProps";

import { EvaluationService, IPendingAssignment } from "../services/EvaluationService";
import EvaluationForm from "./EvaluationForm";
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

export default class EmployeeEvak extends React.Component<IEmployeeEvakProps> {
  public render(): React.ReactElement<IEmployeeEvakProps> {
    const { sp, hasTeamsContext } = this.props;

    return (
      <section className={`${styles.employeeEvak} ${hasTeamsContext ? styles.teams : ""}`}>
        <PendingAssignmentsDashboard sp={sp} />
      </section>
    );
  }
}
