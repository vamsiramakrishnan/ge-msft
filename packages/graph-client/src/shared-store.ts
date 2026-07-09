import type { GraphClient } from './graph-client.js';

/**
 * Structural implementation of `@ge/runtime`'s `SharedStore` port (`docfs/shared-mount.ts`) over
 * the real Microsoft Graph app folder. Kept structural — this class does not import the runtime
 * type — so `@ge/graph-client` never depends on `@ge/runtime`; TypeScript's structural typing makes
 * an instance assignable wherever a `SharedStore` is expected. The caller wires it in (e.g.
 * web-shell's `composeSession`), the same way `DocBridge`/`AuthClient` implementations live outside
 * `@ge/runtime` rather than being owned by it.
 */
export class GraphSharedStore {
  constructor(private readonly client: GraphClient) {}

  list(): Promise<{ name: string; size: number }[]> {
    return this.client.listSharedFiles();
  }

  read(path: string): Promise<string | undefined> {
    return this.client.getSharedFile(path);
  }

  write(path: string, content: string): Promise<void> {
    return this.client.putSharedFile(path, content);
  }

  remove(path: string): Promise<void> {
    return this.client.deleteSharedFile(path);
  }
}
