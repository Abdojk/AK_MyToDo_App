import { TaskCard } from './TaskCard';
import { EventCard } from './EventCard';
import { BUCKET_LABEL } from '../model/types';
import type { Bucket, HubItem, HubTask } from '../model/types';

interface Props {
  bucket: Bucket;
  items: HubItem[];
  zone: string;
  pendingIds: string[];
  onComplete: (task: HubTask) => void;
  onReschedule: (task: HubTask, due: Date) => void;
}

export function BucketColumn({
  bucket,
  items,
  zone,
  pendingIds,
  onComplete,
  onReschedule,
}: Props) {
  const taskCount = items.filter((i) => i.kind === 'task').length;
  const eventCount = items.length - taskCount;

  return (
    <section className={`column column-${bucket}`} aria-label={BUCKET_LABEL[bucket]}>
      <header className="column-head">
        <h2>{BUCKET_LABEL[bucket]}</h2>
        <span
          className="count"
          aria-label={`${taskCount} tasks, ${eventCount} events`}
        >
          {items.length}
        </span>
      </header>

      <div className="column-body">
        {items.length === 0 ? (
          <p className="empty" aria-hidden="true">
            —
          </p>
        ) : (
          items.map((item) =>
            item.kind === 'task' ? (
              <TaskCard
                key={`task-${item.task.id}`}
                task={item.task}
                bucket={bucket}
                zone={zone}
                pending={pendingIds.includes(item.task.id)}
                onComplete={onComplete}
                onReschedule={onReschedule}
              />
            ) : (
              <EventCard
                key={`event-${item.event.id}`}
                event={item.event}
                zone={zone}
              />
            ),
          )
        )}
      </div>
    </section>
  );
}
