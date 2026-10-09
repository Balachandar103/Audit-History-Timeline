import { ReactElement } from "react";

import { AuditTimelinePreviewProps } from "../typings/AuditTimelineProps";

export function preview(_props: AuditTimelinePreviewProps): ReactElement {
    return (
        <div className="audit-timeline-preview">
            <strong>Audit History</strong>
            <div>Timeline preview</div>
        </div>
    );
}