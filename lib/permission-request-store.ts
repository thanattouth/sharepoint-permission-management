import type { PermissionRequestDraft } from "./types";

export interface PermissionRequestStore {
  submit(request: PermissionRequestDraft): Promise<void>;
}

export class DisabledPermissionRequestStore implements PermissionRequestStore {
  async submit() {
    return;
  }
}
