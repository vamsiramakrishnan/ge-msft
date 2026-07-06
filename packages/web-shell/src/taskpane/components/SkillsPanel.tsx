import { useState } from 'react';
import type { Skill } from '../../controller.js';

export interface SkillsPanelProps {
  skills: Skill[];
  disabled?: boolean;
  /** Invoke a registered skill with bound argument values; routes through the agentic loop. */
  onInvoke: (name: string, args: Record<string, string>) => void;
  /** Render expanded inline (no self-collapse/toggle) for use inside the toolbar's icon sheet. */
  embedded?: boolean;
  /** Direction the popover opens from the persistent summary row. */
  placement?: 'below' | 'above';
}

/**
 * One registered skill (ADR-0005 `def`): its name, its `def` registration confirmation, its
 * declared params as bindable fields, and an Invoke action. Invoking does NOT actuate here — it
 * composes the skill call and routes it through the agentic command loop, so the skill's plan still
 * lands on the fail-closed plan-approval card. This component carries no gate logic; it is a
 * read-only surface over the controller's `skills` slice plus an `onInvoke` presenter.
 */
function SkillCard({
  skill,
  disabled = false,
  onInvoke,
}: {
  skill: Skill;
  disabled?: boolean;
  onInvoke: (name: string, args: Record<string, string>) => void;
}): JSX.Element {
  const [args, setArgs] = useState<Record<string, string>>(() =>
    Object.fromEntries(skill.params.map((p) => [p.name, p.example ?? ''])),
  );
  const signature = `${skill.name}(${skill.params.map((p) => p.name).join(', ')})`;
  return (
    <li className="skill" aria-label={`Skill ${skill.name}`}>
      <div className="skill-head">
        <code className="skill-sig">{signature}</code>
        {skill.registered && (
          <span className="skill-badge" title={skill.def ?? 'Registered via def'}>
            {'✓ registered'}
          </span>
        )}
      </div>
      {skill.description && <p className="skill-desc">{skill.description}</p>}
      {skill.def && (
        <pre className="cmd skill-def" aria-label="Registration, shown verbatim">
          {skill.def}
        </pre>
      )}
      {skill.params.length > 0 && (
        <div className="skill-params">
          {skill.params.map((p) => {
            const fieldId = `skill-${skill.name}-${p.name}`;
            return (
              <label key={p.name} className="skill-param" htmlFor={fieldId}>
                <span className="skill-param-name">{p.name}</span>
                <input
                  id={fieldId}
                  className="skill-param-input"
                  value={args[p.name] ?? ''}
                  placeholder={p.example}
                  disabled={disabled}
                  onChange={(e) => setArgs((prev) => ({ ...prev, [p.name]: e.target.value }))}
                />
              </label>
            );
          })}
        </div>
      )}
      <div className="act">
        <button
          type="button"
          className="btn pr"
          disabled={disabled}
          onClick={() => onInvoke(skill.name, args)}
          aria-label={`Invoke ${skill.name}`}
        >
          Invoke skill
        </button>
      </div>
    </li>
  );
}

/**
 * The in-session skills surface (ADR-0005 `def`): the composable programs registered this session,
 * each invokable into a reviewable plan. A skill call expands into the plan-approval card via the
 * agentic loop, so the headline plan-review gate is reused — this section only lists, confirms
 * registration, binds params, and triggers an invocation. Renders nothing when no skills exist.
 */
export function SkillsPanel({
  skills,
  disabled = false,
  onInvoke,
  embedded = false,
  placement = 'below',
}: SkillsPanelProps): JSX.Element | null {
  if (skills.length === 0) return null;
  return (
    <section
      className={`skills detail-hover${embedded ? ' skills--embedded' : ''}`}
      aria-label="Skills"
      aria-disabled={disabled}
      data-placement={placement}
    >
      <div className="skills-h">
        <span className="skills-mark" aria-hidden="true" />
        <span className="skills-title">Session skills</span>
        <span className="skills-summary">{skills.length} registered</span>
      </div>
      <ul className="skills-list skills-popover">
        {skills.map((skill) => (
          <SkillCard key={skill.name} skill={skill} disabled={disabled} onInvoke={onInvoke} />
        ))}
      </ul>
    </section>
  );
}
