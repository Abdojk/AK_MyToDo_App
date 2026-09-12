import { useState } from 'react';
import { PRESETS, fromDateInput, presetDate, toDateInput } from '../model/reschedule';
import type { Preset } from '../model/reschedule';

interface Props {
  /** Names the task in every control's label, because the row repeats per card. */
  label: string;
  due: Date | null;
  zone: string;
  pending: boolean;
  onComplete: () => void;
  onReschedule: (due: Date) => void;
}

/** The Done and Move to row, shared by the To Do and Planner cards. */
export function TaskActions({
  label,
  due,
  zone,
  pending,
  onComplete,
  onReschedule,
}: Props) {
  // The date field stays hidden until asked for: on nine cards at once, a date box
  // per card crowds out the subjects, and the presets cover the common moves.
  const [pickingDate, setPickingDate] = useState(false);

  const moveToDate = (value: string) => {
    const picked = fromDateInput(value, zone);
    if (!picked) return;
    setPickingDate(false);
    onReschedule(picked);
  };

  return (
    <div className="card-actions">
      <button
        type="button"
        className="done"
        onClick={onComplete}
        disabled={pending}
        aria-label={`Mark done: ${label}`}
      >
        {pending ? 'Saving…' : 'Done'}
      </button>

      <div className="move" role="group" aria-label={`Move due date: ${label}`}>
        <span className="move-label">Move to</span>
        {PRESETS.map((preset: { key: Preset; label: string }) => (
          <button
            key={preset.key}
            type="button"
            className="link-button"
            onClick={() => onReschedule(presetDate(preset.key, new Date(), zone))}
            disabled={pending}
            aria-label={`Move to ${preset.label}: ${label}`}
          >
            {preset.label}
          </button>
        ))}
        {pickingDate ? (
          <input
            type="date"
            className="move-date"
            autoFocus
            defaultValue={due ? toDateInput(due, zone) : ''}
            onChange={(e) => moveToDate(e.target.value)}
            onBlur={() => setPickingDate(false)}
            disabled={pending}
            aria-label={`Pick a due date: ${label}`}
          />
        ) : (
          <button
            type="button"
            className="link-button"
            onClick={() => setPickingDate(true)}
            disabled={pending}
            aria-label={`Pick a due date: ${label}`}
          >
            Pick a date
          </button>
        )}
      </div>
    </div>
  );
}
