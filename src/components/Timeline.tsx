import { BucketColumn } from './BucketColumn';
import { groupItems } from '../model/buckets';
import { BUCKET_ORDER } from '../model/types';
import type { HubItem, HubTask } from '../model/types';

interface Props {
  items: HubItem[];
  now: Date;
  zone: string;
  pendingIds: string[];
  onComplete: (task: HubTask) => void;
  onReschedule: (task: HubTask, due: Date) => void;
}

export function Timeline({
  items,
  now,
  zone,
  pendingIds,
  onComplete,
  onReschedule,
}: Props) {
  const grouped = groupItems(items, now, zone);

  return (
    <div className="timeline">
      {BUCKET_ORDER.map((bucket) => (
        <BucketColumn
          key={bucket}
          bucket={bucket}
          items={grouped[bucket]}
          zone={zone}
          pendingIds={pendingIds}
          onComplete={onComplete}
          onReschedule={onReschedule}
        />
      ))}
    </div>
  );
}
