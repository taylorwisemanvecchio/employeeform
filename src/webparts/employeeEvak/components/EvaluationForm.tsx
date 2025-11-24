import * as React from "react";
import {
  Dropdown,
  IDropdownOption,
  TextField,
  PrimaryButton,
  DefaultButton,
  Stack
} from "@fluentui/react";
import { SPFI } from "@pnp/sp";

import {
  EvaluationService,
  IEvaluationResponse,
  ReviewerType
} from "../services/EvaluationService";
import { CATEGORIES } from "../data/questions";

type Likert = 1 | 2 | 3 | 4;

interface IQuestionDef {
  key: string;
  commentKey: string;
  text: string;
}

interface ICategoryDef {
  name: string;
  questions: IQuestionDef[];
}

const ratingOptions: IDropdownOption[] = [
  { key: 1, text: "1 - Unsatisfactory" },
  { key: 2, text: "2 - Needs Development" },
  { key: 3, text: "3 - Meets Expectations" },
  { key: 4, text: "4 - Exceeds Expectations" }
];

export const EvaluationForm: React.FC<{
  sp: SPFI;
  assignmentId: number;
  reviewerType: string;
}> = ({ sp, assignmentId, reviewerType }) => {
  const svc = React.useMemo((): EvaluationService => new EvaluationService(sp), [sp]);

  const [response, setResponse] =
    React.useState<IEvaluationResponse | undefined>(undefined);
  const [responseId, setResponseId] =
    React.useState<number | undefined>(undefined);
  const [categoryIdx, setCategoryIdx] =
    React.useState<number>(0);
  const [saving, setSaving] =
    React.useState<boolean>(false);
  const [loading, setLoading] =
    React.useState<boolean>(true);
  const [err, setErr] =
    React.useState<string | undefined>(undefined);

  const initKeyRef = React.useRef<string | undefined>(undefined);

  React.useEffect((): void => {
    if (!assignmentId || isNaN(assignmentId)) {
      setResponse(undefined);
      setResponseId(undefined);
      setLoading(false);
      return;
    }

    const initKey = `${assignmentId}|${reviewerType}`;
    if (initKeyRef.current === initKey) return;
    initKeyRef.current = initKey;

    (async (): Promise<void> => {
      try {
        setLoading(true);
        setErr(undefined);

        const rt = (reviewerType as ReviewerType) || "Employee";

        const me = await svc.getCurrentUser();
        if (!me || typeof me.Id !== "number") {
          throw new Error("Unable to resolve current user. Please refresh.");
        }

        const assignment = await svc.getAssignment(assignmentId);
        if (!assignment || typeof assignment.Id !== "number") {
          throw new Error("Assignment not found. Please refresh.");
        }

        // Check if a response already exists
        const existing = await svc.getMyResponse(assignmentId, rt, me.Email);

        if (existing && typeof existing.Id === "number") {
          // Load existing response
          setResponse(existing);
          setResponseId(existing.Id);
        } else {
          // Initialize empty response - we'll create it on first save
          setResponse({} as IEvaluationResponse);
          setResponseId(undefined);
        }

        setCategoryIdx(0);
      } catch (e: unknown) {
        const msg = e instanceof Error ? e.message : "Failed to load evaluation form.";
        setErr(msg);
      } finally {
        setLoading(false);
      }
    })().catch((): void => {});
  }, [assignmentId, reviewerType, svc]);

  const updateField = (field: string, value: unknown): void => {
    setResponse((prev: IEvaluationResponse | undefined): IEvaluationResponse | undefined => {
      if (!prev) return prev;
      return { ...prev, [field]: value };
    });
  };

  const onSave = async (submit = false): Promise<void> => {
    if (!response) return;

    try {
      setSaving(true);

      const rt = (reviewerType as ReviewerType) || "Employee";

      const payload: Record<string, unknown> = {
        AssignmentIDId: assignmentId,
        ReviewerType: rt
      };

      (CATEGORIES as ICategoryDef[]).forEach((cat: ICategoryDef): void => {
        cat.questions.forEach((q: IQuestionDef): void => {
          const rating = response[q.key];
          const comment = response[q.commentKey];

          payload[q.key] =
            typeof rating === "number"
              ? rating.toString()
              : typeof rating === "string"
                ? rating
                : undefined;

          payload[q.commentKey] =
            typeof comment === "string" ? comment : "";
        });
      });

      if (submit) {
        payload.SubmittedDate = new Date().toISOString();
      }

      // If we don't have a responseId yet, create the record
      let currentResponseId = responseId;
      if (!currentResponseId) {
        const me = await svc.getCurrentUser();
        const assignment = await svc.getAssignment(assignmentId);

        const createPayload = {
          ...payload,
          Title: `${assignment.Title || "Evaluation"} - ${rt}`,
          ReviewerNameId: me.Id
        };

        const created = await svc.createResponse(createPayload);
        currentResponseId = created.Id;
        setResponseId(created.Id);
      } else {
        // Update existing response
        await svc.updateResponse(currentResponseId, payload);
      }

      if (submit) {
        await svc.markSubmitted(assignmentId, rt);
      }

      alert(submit ? "Submitted!" : "Saved");
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : "Save failed.";
      alert(msg);
    } finally {
      setSaving(false);
    }
  };

  if (loading) {
    return (
      <div style={{ padding: 16 }}>
        <h3>Loading evaluation…</h3>
      </div>
    );
  }

  if (err) {
    return (
      <div style={{ padding: 16 }}>
        <h3>Something went wrong</h3>
        <p>{err}</p>
      </div>
    );
  }

  if (!response) {
    return (
      <div style={{ padding: 16 }}>
        <h3>Preparing your evaluation…</h3>
        <p>Please wait a moment…</p>
      </div>
    );
  }

  const cats = CATEGORIES as ICategoryDef[];
  const cat = cats[categoryIdx];
  const isLastCategory = categoryIdx >= cats.length - 1;

  return (
    <div style={{ display: "flex", gap: 16 }}>
      {/* Left Nav */}
      <div style={{ width: 260 }}>
        {cats.map((c: ICategoryDef, i: number) => (
          <div
            key={c.name}
            onClick={(): void => setCategoryIdx(i)}
            style={{
              padding: "10px 12px",
              cursor: "pointer",
              background: i === categoryIdx ? "#0b6a53" : "white",
              color: i === categoryIdx ? "white" : "#333",
              borderRadius: 6,
              marginBottom: 6
            }}
          >
            {c.name}
          </div>
        ))}
      </div>

      {/* Main Content */}
      <div style={{ flex: 1 }}>
        <h2>{cat.name}</h2>

        {cat.questions.map((q: IQuestionDef) => {
          const ratingVal = response[q.key];
          const selectedKey =
            typeof ratingVal === "number" ? (ratingVal as Likert) : undefined;

          const commentVal = response[q.commentKey];
          const commentText =
            typeof commentVal === "string" ? commentVal : "";

          return (
            <div
              key={q.key}
              style={{
                padding: 12,
                border: "1px solid #eee",
                borderRadius: 8,
                marginBottom: 12
              }}
            >
              <div style={{ fontWeight: 600, color: "#0b6a53", marginBottom: 6 }}>
                {q.text}
              </div>

              <Stack horizontal tokens={{ childrenGap: 12 }}>
                <Dropdown
                  label="My Rating"
                  options={ratingOptions}
                  selectedKey={selectedKey}
                  onChange={(_, opt): void => updateField(q.key, opt?.key)}
                  styles={{ root: { width: 280 } }}
                />

                <TextField
                  label="Comments"
                  multiline
                  autoAdjustHeight
                  value={commentText}
                  onChange={(_, v): void => updateField(q.commentKey, v ?? "")}
                  styles={{ root: { flex: 1 } }}
                />
              </Stack>
            </div>
          );
        })}

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton
            text="Save"
            onClick={(): void => {
              onSave(false).catch((): void => {});
            }}
            disabled={saving}
          />

          {!isLastCategory ? (
            <PrimaryButton
              text="Next Section"
              onClick={(): void => setCategoryIdx(categoryIdx + 1)}
              disabled={saving}
            />
          ) : (
            <PrimaryButton
              text="Submit"
              onClick={(): void => {
                onSave(true).catch((): void => {});
              }}
              disabled={saving}
            />
          )}
        </Stack>
      </div>

      {/* Rating Scale Panel */}
      <div style={{ width: 260, borderLeft: "1px solid #eee", paddingLeft: 12 }}>
        <h3>Rating Scale</h3>
        <div>1 – Unsatisfactory</div>
        <div>2 – Needs Development</div>
        <div>3 – Meets Expectations</div>
        <div>4 – Exceeds Expectations</div>
      </div>
    </div>
  );
};

export default EvaluationForm;
