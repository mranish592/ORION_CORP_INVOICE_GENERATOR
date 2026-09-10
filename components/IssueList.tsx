import type { Issue } from "@/lib/types";

export function IssueList({
  title,
  issues,
  tone,
}: {
  title: string;
  issues: Issue[];
  tone: "error" | "warning";
}) {
  if (issues.length === 0) return null;

  return (
    <div className={`notice ${tone}`} role={tone === "error" ? "alert" : "status"}>
      <h3>{title}</h3>
      <ul>
        {issues.map((issue, index) => (
          <li key={`${issue.location ?? ""}-${issue.message}-${index}`}>
            {issue.location ? <code>{issue.location}</code> : null}
            {issue.location ? " — " : null}
            {issue.message}
          </li>
        ))}
      </ul>
    </div>
  );
}
