import { TaskCard } from './TaskCard';
import { EventCard } from './EventCard';
import { PlannerCard } from './PlannerCard';
import { BUCKET_LABEL, itemKey } from '../model/types';
import type { Bucket, HubItem, HubPlannerTask, HubTask } from '../model/types';

interface Props {
  bucket: Bucket;
  items: HubItem[];
  zone: string;
  pendingKeys: string[];
  onComplete: (task: HubTask) => void;
  onReschedule: (task: HubTask, due: Date) => void;
  onCompletePlanner: (task: HubPlannerTask) => void;
  onReschedulePlanner: (task: HubPlannerTask, due: Date) => void;
}

export function BucketColumn({
  bucket,
  items,
  zone,
  pendingKeys,
  onComplete,
  onReschedule,
  onCompletePlanner,
  onReschedulePlanner,
}: Props) {
  const eventCount = items.filter((i) => i.kind === 'event').length;

  return (
    <section className={`column column-${bucket}`} aria-label={BUCKET_LABEL[bucket]}>
      <header className="column-head">
        <h2>{BUCKET_LABEL[bucket]}</h2>
        <span
          className="count"
          aria-label={`${items.length - eventCount} tasks, ${eventCount} events`}
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
          items.map((item) => {
            const key = itemKey(item);
            const pending = pendingKeys.includes(key);

            if (item.kind === 'task') {
              return (
                <TaskCard
                  key={key}
                  task={item.task}
                  bucket={bucket}
                  zone={zone}
                  pending={pending}
                  onComplete={onComplete}
                  onReschedule={onReschedule}
                />
              );
            }

            if (item.kind === 'planner') {
              return (
                <PlannerCard
                  key={key}
                  task={item.task}
                  zone={zone}
                  pending={pending}
                  onComplete={onCompletePlanner}
                  onReschedule={onReschedulePlanner}
                />
              );
            }

            return <EventCard key={key} event={item.event} zone={zone} />;
          })
        )}
      </div>
    </section>
  );
}
