import { TaskCard } from './TaskCard';
import { BUCKET_LABEL } from '../model/types';
import type { Bucket, HubTask } from '../model/types';

interface Props {
  bucket: Bucket;
  tasks: HubTask[];
  zone: string;
  pendingIds: string[];
  onComplete: (task: HubTask) => void;
  onReschedule: (task: HubTask, due: Date) => void;
}

export function BucketColumn({
  bucket,
  tasks,
  zone,
  pendingIds,
  onComplete,
  onReschedule,
}: Props) {
  return (
    <section className={`column column-${bucket}`} aria-label={BUCKET_LABEL[bucket]}>
      <header className="column-head">
        <h2>{BUCKET_LABEL[bucket]}</h2>
        <span className="count" aria-label={`${tasks.length} tasks`}>
          {tasks.length}
        </span>
      </header>

      <div className="column-body">
        {tasks.length === 0 ? (
          <p className="empty" aria-hidden="true">
            —
          </p>
        ) : (
          tasks.map((task) => (
            <TaskCard
              key={task.id}
              task={task}
              bucket={bucket}
              zone={zone}
              pending={pendingIds.includes(task.id)}
              onComplete={onComplete}
              onReschedule={onReschedule}
            />
          ))
        )}
      </div>
    </section>
  );
}
