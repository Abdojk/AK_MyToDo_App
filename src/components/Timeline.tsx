import { BucketColumn } from './BucketColumn';
import { groupByBucket } from '../model/buckets';
import { BUCKET_ORDER } from '../model/types';
import type { HubTask } from '../model/types';

interface Props {
  tasks: HubTask[];
  now: Date;
  zone: string;
  pendingIds: string[];
  onComplete: (task: HubTask) => void;
  onReschedule: (task: HubTask, due: Date) => void;
}

export function Timeline({
  tasks,
  now,
  zone,
  pendingIds,
  onComplete,
  onReschedule,
}: Props) {
  const grouped = groupByBucket(tasks, now, zone);

  return (
    <div className="timeline">
      {BUCKET_ORDER.map((bucket) => (
        <BucketColumn
          key={bucket}
          bucket={bucket}
          tasks={grouped[bucket]}
          zone={zone}
          pendingIds={pendingIds}
          onComplete={onComplete}
          onReschedule={onReschedule}
        />
      ))}
    </div>
  );
}
