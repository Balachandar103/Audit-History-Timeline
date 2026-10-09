/**
 * This file was generated from AuditTimeline.xml
 * WARNING: All changes made to this file will be overwritten
 * @author Mendix Widgets Framework Team
 */
import { ListAttributeValue, ListValue } from "mendix";
import { CSSProperties } from "react";

export interface AuditTimelineContainerProps {
    name: string;
    class: string;
    style?: CSSProperties;
    tabIndex?: number;
    dataSource: ListValue;
    action: ListAttributeValue<string>;
    user: ListAttributeValue<string>;
    changedDateTime: ListAttributeValue<Date>;
    objectName: ListAttributeValue<string>;
    attributeName: ListAttributeValue<string>;
    oldValue: ListAttributeValue<string>;
    newValue: ListAttributeValue<string>;
    description: ListAttributeValue<string>;
}

export interface AuditTimelinePreviewProps {
    /**
     * @deprecated Deprecated since version 9.18.0. Please use class property instead.
     */
    className: string;
    class: string;
    style: string;
    styleObject?: CSSProperties;
    readOnly: boolean;
    renderMode: "design" | "xray" | "structure";
    translate: (text: string) => string;
    dataSource: {} | { caption: string } | { type: string } | null;
    action: string;
    user: string;
    changedDateTime: string;
    objectName: string;
    attributeName: string;
    oldValue: string;
    newValue: string;
    description: string;
}
