import type { CSSProperties } from "react";

export function ReportSkeleton() {
  return (
    <div className="skeleton-stack" aria-label="Loading report">
      <div className="report-meta skeleton-meta">
        <span className="skeleton-line short" />
        <strong className="skeleton-line medium" />
      </div>
      <TableSkeleton columns={6} rows={1} />
      <TableSkeleton columns={4} rows={5} />
    </div>
  );
}

export function TableSkeleton({ columns, rows }: { columns: number; rows: number }) {
  return (
    <>
      {Array.from({ length: rows }).map((_, rowIndex) => (
        <div
          className="skeleton-table-row"
          role="row"
          key={`skeleton-row-${rowIndex}`}
          style={{ "--skeleton-columns": columns } as CSSProperties}
        >
          {Array.from({ length: columns }).map((__, columnIndex) => (
            <span
              className={`skeleton-line ${columnIndex === 0 ? "wide" : columnIndex % 2 === 0 ? "medium" : "short"}`}
              key={`skeleton-cell-${rowIndex}-${columnIndex}`}
            />
          ))}
        </div>
      ))}
    </>
  );
}
