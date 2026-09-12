import { BucketColumn } from './BucketColumn';
import { groupByBucket } from '../model/buckets';
import { BUCKET_ORDER } from '../model/types';
import type { HubTask } from '../model/types';

interface Props {
  tasks: HubTask[];
  now: Date;
  zone: string;
}

export function Timeline({ tasks, now, zone }: Props) {
  const grouped = groupByBucket(tasks, now, zone);

  return (
    <div className="timeline">
      {BUCKET_ORDER.map((bucket) => (
        <BucketColumn
          key={bucket}
          bucket={bucket}
          tasks={grouped[bucket]}
          zone={zone}
        />
      ))}
    </div>
  );
}
