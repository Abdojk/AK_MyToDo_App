import { TaskActions } from './TaskActions';
import { formatDueDate } from '../model/dates';
import { isPressing } from '../model/planner';
import type { HubPlannerTask } from '../model/types';

interface Props {
  task: HubPlannerTask;
  zone: string;
  pending: boolean;
  onComplete: (task: HubPlannerTask) => void;
  onReschedule: (task: HubPlannerTask, due: Date) => void;
}

const PRIORITY_LABEL: Record<string, string> = {
  urgent: 'Urgent',
  important: 'Important',
};

export function PlannerCard({
  task,
  zone,
  pending,
  onComplete,
  onReschedule,
}: Props) {
  const priority = isPressing(task.priority) ? PRIORITY_LABEL[task.priority] : null;

  return (
    <article className={`card card-planner${pending ? ' card-pending' : ''}`}>
      <div className="card-head">
        {priority && (
          <span className={`pill pill-${task.priority}`} title="Planner priority">
            {priority}
          </span>
        )}
        <h3 className="card-title">{task.title}</h3>
      </div>

      <dl className="card-meta">
        <dt>Plan</dt>
        <dd>{task.planTitle ?? 'Planner'}</dd>
        <dt>Due</dt>
        <dd>{formatDueDate(task.due, zone)}</dd>
        {task.checklistTotal > 0 && (
          <>
            <dt>Steps</dt>
            <dd>
              {task.checklistDone} of {task.checklistTotal} done
            </dd>
          </>
        )}
      </dl>

      <TaskActions
        label={task.title}
        due={task.due}
        zone={zone}
        pending={pending}
        onComplete={() => onComplete(task)}
        onReschedule={(due) => onReschedule(task, due)}
      />

      {task.url && (
        <a
          className="card-link"
          href={task.url}
          target="_blank"
          rel="noreferrer noopener"
        >
          Open in Planner
        </a>
      )}
    </article>
  );
}
