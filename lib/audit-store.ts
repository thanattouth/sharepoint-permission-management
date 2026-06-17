import type { AuditLogDraft, AuditLogRecord } from "./types";

export interface AuditStore {
  write(entry: AuditLogDraft): Promise<void>;
  list(limit?: number): Promise<AuditLogRecord[]>;
}

export class DisabledAuditStore implements AuditStore {
  async write() {
    return;
  }

  async list() {
    return [];
  }
}
