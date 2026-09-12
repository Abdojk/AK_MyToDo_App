import { TaskActions } from './TaskActions';
import { formatDueDate } from '../model/dates';
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
        {task.isFlaggedEmail ? (
          task.sender && (
            <>
              <dt>From</dt>
              <dd>{task.sender}</dd>
            </>
          )
        ) : (
          <>
            <dt>List</dt>
            <dd>{task.listName}</dd>
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

      <TaskActions
        label={task.title}
        due={task.due}
        zone={zone}
        pending={pending}
        onComplete={() => onComplete(task)}
        onReschedule={(due) => onReschedule(task, due)}
      />

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
