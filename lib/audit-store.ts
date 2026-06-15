import type { AuditLogDraft } from "./types";

export interface AuditStore {
  write(entry: AuditLogDraft): Promise<void>;
}

export class DisabledAuditStore implements AuditStore {
  async write() {
    return;
  }
}
