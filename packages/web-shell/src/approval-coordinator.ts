import type { ChangeId } from '@ge/contracts';
import type { PendingWrite, PendingPlan, PendingShare } from './controller.js';

/**
 * The fail-closed approval state machine, extracted from `PanelController` (E-full) as the single
 * most tangled + correctness-critical responsibility (review Finding #6). It owns the *control* state
 * — the per-write changeId, and the two resolver promises the command loop is awaiting — while the
 * *view* state (`pendingWrite`/`pendingPlan` on `PanelState`) stays in the controller, pushed here
 * through the `showWrite`/`showPlan` callbacks. Invariants preserved exactly:
 *
 *  - **Fail-closed:** nothing resolves `true` except an explicit user approval; every abandon path
 *    (`releaseAwaiting`/`releaseAll`) resolves `false` so no write/plan actuates.
 *  - **Approve keeps the card** until the loop narrates the `write-result` (`consumeWriteResult`);
 *    reject drops it immediately.
 *  - **Superseded decisions ignored:** an approve/reject carrying a changeId that no longer matches
 *    the staged write (a late click on a replaced card) is a no-op.
 */
export class ApprovalCoordinator {
  private writeId: ChangeId | undefined;
  private resolveWrite: ((approved: boolean) => void) | undefined;
  private resolvePlan: ((approved: boolean) => void) | undefined;
  private resolveShare: ((approved: boolean) => void) | undefined;
  private writeShown = false;
  private planShown = false;
  private shareShown = false;

  constructor(
    private readonly showWrite: (write: PendingWrite | undefined) => void,
    private readonly showPlan: (plan: PendingPlan | undefined) => void,
    private readonly showShare: (share: PendingShare | undefined) => void,
  ) {}

  // ---- loop side ----------------------------------------------------------

  /** Stage a per-write decision (ADR-0004) and await it; the returned promise is what the loop awaits. */
  awaitWrite(view: PendingWrite, id: ChangeId): Promise<boolean> {
    this.settleWrite(false); // defensive: release any still-open prior decision first
    return new Promise<boolean>((resolve) => {
      this.resolveWrite = resolve;
      this.writeId = id;
      this.writeShown = true;
      this.showWrite(view);
    });
  }

  /** Stage the plan-level decision (ADR-0005) and await ONE decision over the dry-run effect-set. */
  awaitPlan(view: PendingPlan): Promise<boolean> {
    this.settlePlan(false);
    return new Promise<boolean>((resolve) => {
      this.resolvePlan = resolve;
      this.planShown = true;
      this.showPlan(view);
    });
  }

  /**
   * Stage the `share` decision (the `/shared` cross-surface handoff store is an estate-class write —
   * see `AssistSessionOptions.estateWritesEnabled`/`RunCommandsOptions.approveShare`) and await it.
   * Unlike a per-write decision, the actual Graph write happens synchronously right after this
   * resolves (there is no separate later `write-result`-style event to keep the card alive for), so
   * — like `awaitPlan` — either decision drops the card immediately.
   */
  awaitShare(view: PendingShare): Promise<boolean> {
    this.settleShare(false);
    return new Promise<boolean>((resolve) => {
      this.resolveShare = resolve;
      this.shareShown = true;
      this.showShare(view);
    });
  }

  /** The loop consumed a write decision and narrated its `write-result` → drop the staged card. */
  consumeWriteResult(): void {
    // Defense-in-depth: a `write-result` should only narrate AFTER the decision settled, so the
    // resolver must already be undefined. If one is still awaiting (an out-of-order consume that
    // should be impossible), settle it fail-closed FIRST so the invariant breaks loudly to `false`
    // rather than leaving a resolver dangling for the loop's `finally` to sweep.
    this.settleWrite(false);
    this.clearWrite();
  }

  // ---- user side ----------------------------------------------------------

  approveWrite(id?: ChangeId): void {
    if (!this.superseded(id)) this.settleWrite(true);
  }
  rejectWrite(id?: ChangeId): void {
    if (!this.superseded(id)) this.settleWrite(false);
  }
  approvePlan(): void {
    this.settlePlan(true);
  }
  rejectPlan(): void {
    this.settlePlan(false);
  }
  approveShare(): void {
    this.settleShare(true);
  }
  rejectShare(): void {
    this.settleShare(false);
  }

  // ---- terminal paths -----------------------------------------------------

  /** Release any AWAITING decision fail-closed (e.g. on cancel) — leaves no resolver hanging. */
  releaseAwaiting(): void {
    this.settleWrite(false);
    this.settlePlan(false);
    this.settleShare(false);
  }

  /**
   * Release any awaiting decision fail-closed AND drop both cards unconditionally — for the loop's
   * `finally`, where a decision may have been ALREADY consumed (an approval whose execution then
   * threw, or a write that produced no write-result) so `settle*` is a no-op but a card could linger.
   */
  releaseAll(): void {
    this.settleWrite(false);
    this.settlePlan(false);
    this.settleShare(false);
    this.clearWrite();
    this.clearPlan();
    this.clearShare();
  }

  // ---- internals ----------------------------------------------------------

  private superseded(id: ChangeId | undefined): boolean {
    return id !== undefined && id !== this.writeId;
  }

  private settleWrite(approved: boolean): void {
    const resolve = this.resolveWrite;
    if (!resolve) return;
    this.resolveWrite = undefined;
    if (!approved) this.clearWrite(); // approve keeps the card until the write-result narrates
    resolve(approved);
  }

  private settlePlan(approved: boolean): void {
    const resolve = this.resolvePlan;
    if (!resolve) return;
    this.resolvePlan = undefined;
    this.clearPlan(); // either decision drops the plan card; the per-effect outcomes narrate as steps
    resolve(approved);
  }

  private clearWrite(): void {
    this.writeId = undefined;
    if (this.writeShown) {
      this.writeShown = false;
      this.showWrite(undefined);
    }
  }

  private clearPlan(): void {
    if (this.planShown) {
      this.planShown = false;
      this.showPlan(undefined);
    }
  }

  private settleShare(approved: boolean): void {
    const resolve = this.resolveShare;
    if (!resolve) return;
    this.resolveShare = undefined;
    this.clearShare(); // either decision drops the card — the write (if any) narrates as a step
    resolve(approved);
  }

  private clearShare(): void {
    if (this.shareShown) {
      this.shareShown = false;
      this.showShare(undefined);
    }
  }
}
