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
    return <div style={{ padding: 16 }}>Loading pending evaluations…</div>;
  }

  if (error) {
    return (
      <div style={{ padding: 16 }}>
        <h3>Something went wrong</h3>
        <p>{error}</p>
      </div>
    );
  }

  if (selected) {
    // Validate selected assignment has a valid Id
    if (!selected.Id || typeof selected.Id !== 'number') {
      return (
        <div style={{ padding: 16 }}>
          <h3>Something went wrong</h3>
          <p>Invalid assignment selected. Please go back and try again.</p>
          <DefaultButton
            text="Back to Pending List"
            onClick={(): void => setSelected(undefined)}
          />
        </div>
      );
    }

    return (
      <div style={{ padding: 16 }}>
        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 12 }}>
          <DefaultButton
            text="Back to Pending List"
            onClick={(): void => setSelected(undefined)}
          />
        </Stack>

        <h2 style={{ marginTop: 0 }}>{selected.Title}</h2>
        <div style={{ marginBottom: 12, color: "#555" }}>
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
    <div style={{ padding: 16 }}>
      <h2 style={{ marginTop: 0 }}>Pending Evaluations</h2>

      {pending.length === 0 && <div>No pending evaluations assigned to you.</div>}

      <Stack tokens={{ childrenGap: 10 }}>
        {pending
          .filter((a: IPendingAssignment) => typeof a.Id === 'number')
          .map((a: IPendingAssignment) => (
            <div
              key={`${a.Id}|${a.MyRole}`}
              style={{
                border: "1px solid #eee",
                borderRadius: 8,
                padding: 12,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center"
              }}
            >
              <div>
                <div style={{ fontWeight: 600 }}>{a.Title}</div>
                <div style={{ fontSize: 12, color: "#666" }}>
                  Role: {a.MyRole}
                </div>
                {(a.ReviewPeriodStart || a.ReviewPeriodEnd) && (
                  <div style={{ fontSize: 12, color: "#666" }}>
                    Period: {a.ReviewPeriodStart ?? "—"} to {a.ReviewPeriodEnd ?? "—"}
                  </div>
                )}
              </div>

              <PrimaryButton text="Open" onClick={(): void => setSelected(a)} />
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
