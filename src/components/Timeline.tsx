import { BucketColumn } from './BucketColumn';
import { groupItems } from '../model/buckets';
import { BUCKET_ORDER } from '../model/types';
import type { HubItem, HubPlannerTask, HubTask } from '../model/types';

interface Props {
  items: HubItem[];
  now: Date;
  zone: string;
  pendingKeys: string[];
  onComplete: (task: HubTask) => void;
  onReschedule: (task: HubTask, due: Date) => void;
  onCompletePlanner: (task: HubPlannerTask) => void;
  onReschedulePlanner: (task: HubPlannerTask, due: Date) => void;
}

export function Timeline({ items, now, zone, ...handlers }: Props) {
  const grouped = groupItems(items, now, zone);

  return (
    <div className="timeline">
      {BUCKET_ORDER.map((bucket) => (
        <BucketColumn
          key={bucket}
          bucket={bucket}
          items={grouped[bucket]}
          zone={zone}
          {...handlers}
        />
      ))}
    </div>
  );
}
