import { formatEventRange } from '../model/events';
import type { HubEvent } from '../model/types';

interface Props {
  event: HubEvent;
  zone: string;
}

/** A soft hold should not read as a commitment. */
const SOFT: Record<string, string> = {
  free: 'Free',
  tentative: 'Tentative',
  workingElsewhere: 'Elsewhere',
  oof: 'Out of office',
};

export function EventCard({ event, zone }: Props) {
  const soft = SOFT[event.showAs];
  const detail = [event.organiser, event.location].filter(Boolean).join(' · ');

  return (
    <article className="card card-event">
      <div className="event-time">
        <span className="event-clock">{formatEventRange(event, zone)}</span>
        {event.isOnlineMeeting && (
          <span className="pill pill-online" title="Online meeting">
            Online
          </span>
        )}
        {soft && <span className="pill">{soft}</span>}
        {event.isRecurring && (
          <span className="pill" title="Part of a recurring series">
            Repeats
          </span>
        )}
      </div>

      <h3 className="card-title">{event.subject}</h3>

      {detail && <p className="event-detail">{detail}</p>}

      {event.outlookUrl && (
        <a
          className="card-link"
          href={event.outlookUrl}
          target="_blank"
          rel="noreferrer noopener"
        >
          Open in Outlook
        </a>
      )}
    </article>
  );
}
