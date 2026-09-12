import { formatDueDate } from '../model/dates';
import type { Bucket, HubTask } from '../model/types';

interface Props {
  task: HubTask;
  bucket: Bucket;
  zone: string;
}

export function TaskCard({ task, bucket, zone }: Props) {
  return (
    <article className={`card card-${bucket}`}>
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
