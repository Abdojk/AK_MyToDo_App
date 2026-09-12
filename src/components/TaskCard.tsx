import { useState } from 'react';
import { formatDueDate } from '../model/dates';
import { PRESETS, fromDateInput, presetDate, toDateInput } from '../model/reschedule';
import type { Preset } from '../model/reschedule';
import type { Bucket, HubTask } from '../model/types';

interface Props {
  task: HubTask;
  bucket: Bucket;
  zone: string;
  pending: boolean;
  onComplete: (task: HubTask) => void;
  onReschedule: (task: HubTask, due: Date) => void;
}

export function TaskCard({
  task,
  bucket,
  zone,
  pending,
  onComplete,
  onReschedule,
}: Props) {
  // The date field stays hidden until asked for: on nine cards at once, a date box
  // per card crowds out the subjects, and the presets cover the common moves.
  const [pickingDate, setPickingDate] = useState(false);

  const movePreset = (preset: Preset) =>
    onReschedule(task, presetDate(preset, new Date(), zone));

  const moveToDate = (value: string) => {
    const due = fromDateInput(value, zone);
    if (!due) return;
    setPickingDate(false);
    onReschedule(task, due);
  };

  return (
    <article className={`card card-${bucket}${pending ? ' card-pending' : ''}`}>
      <div className="card-head">
        {task.importance === 'high' && (
          <span className="pill pill-high" title="High importance">
            High
          </span>
        )}
        <h3 className="card-title">{task.title}</h3>
      </div>

      <dl className="card-meta">
        {task.sender && (
          <>
            <dt>From</dt>
            <dd>{task.sender}</dd>
          </>
        )}
        <dt>Due</dt>
        <dd>{task.due ? formatDueDate(task.due, zone) : 'Not set'}</dd>
      </dl>

      {task.bodyPreview && <p className="card-body">{task.bodyPreview}</p>}

      {task.categories.length > 0 && (
        <div className="card-tags">
          {task.categories.map((category) => (
            <span className="pill" key={category}>
              {category}
            </span>
          ))}
        </div>
      )}

      <div className="card-actions">
        <button
          type="button"
          className="done"
          onClick={() => onComplete(task)}
          disabled={pending}
          aria-label={`Mark done: ${task.title}`}
        >
          {pending ? 'Saving…' : 'Done'}
        </button>

        <div className="move" role="group" aria-label={`Move due date: ${task.title}`}>
          <span className="move-label">Move to</span>
          {PRESETS.map((preset) => (
            <button
              key={preset.key}
              type="button"
              className="link-button"
              onClick={() => movePreset(preset.key)}
              disabled={pending}
              aria-label={`Move to ${preset.label}: ${task.title}`}
            >
              {preset.label}
            </button>
          ))}
          {pickingDate ? (
            <input
              type="date"
              className="move-date"
              autoFocus
              defaultValue={task.due ? toDateInput(task.due, zone) : ''}
              onChange={(e) => moveToDate(e.target.value)}
              onBlur={() => setPickingDate(false)}
              disabled={pending}
              aria-label={`Pick a due date: ${task.title}`}
            />
          ) : (
            <button
              type="button"
              className="link-button"
              onClick={() => setPickingDate(true)}
              disabled={pending}
              aria-label={`Pick a due date: ${task.title}`}
            >
              Pick a date
            </button>
          )}
        </div>
      </div>

      {task.outlookUrl && (
        <a
          className="card-link"
          href={task.outlookUrl}
          target="_blank"
          rel="noreferrer noopener"
        >
          Open in Outlook
        </a>
      )}
    </article>
  );
}
