/**
 * The signed-in Office user's UPN, used only as an MSAL login hint for silent sign-in. It is
 * optional, so a host that never answers must not block start-up: on PowerPoint for the web
 * `Office.auth.getAuthContext()` stayed pending indefinitely and the task pane never rendered
 * (live 2026-10-01, docs/COMMAND-RELIABILITY.md).
 */
export async function getOfficeLoginHint(timeoutMs = 2000): Promise<string | undefined> {
  try {
    const officeAuth = (
      globalThis as {
        Office?: {
          auth?: {
            getAuthContext?: () => Promise<{ userPrincipalName?: string | null }>;
          };
        };
      }
    ).Office?.auth;
    const context = officeAuth?.getAuthContext?.();
    if (!context) return undefined;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timeout = new Promise<undefined>((resolve) => {
      timer = setTimeout(() => resolve(undefined), timeoutMs);
    });
    try {
      const result = await Promise.race([context, timeout]);
      return result?.userPrincipalName?.trim() || undefined;
    } finally {
      clearTimeout(timer);
    }
  } catch {
    return undefined;
  }
}
