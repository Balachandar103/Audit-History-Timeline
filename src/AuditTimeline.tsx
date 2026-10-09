import { ReactElement, useEffect, useState } from "react";

import { AuditTimelineContainerProps } from "../typings/AuditTimelineProps";

import "./ui/AuditTimeline.css";
import * as XLSX from "xlsx";

function formatDate(value: Date | undefined): string {
    if (!value) {
        return "";
    }

    return new Intl.DateTimeFormat("en-GB", {
        day: "2-digit",
        month: "short",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: true
    }).format(value);
}

function getActionClass(action: string): string {
    const normalizedAction = action.toLowerCase();

    if (normalizedAction.includes("create")) {
        return "audit-timeline-item-created";
    }

    if (normalizedAction.includes("delete")) {
        return "audit-timeline-item-deleted";
    }

    if (normalizedAction.includes("approve")) {
        return "audit-timeline-item-approved";
    }

    if (normalizedAction.includes("reject")) {
        return "audit-timeline-item-rejected";
    }

    return "audit-timeline-item-updated";
}

function getActionIcon(action: string): string {
    const normalizedAction = action.toLowerCase();

    if (normalizedAction.includes("create")) {
        return "+";
    }

    if (normalizedAction.includes("delete")) {
        return "×";
    }

    if (normalizedAction.includes("approve")) {
        return "✓";
    }

    if (normalizedAction.includes("reject")) {
        return "!";
    }

    return "↻";
}
function toLocalDateInput(value: Date | undefined): string {
    if (!value || Number.isNaN(value.getTime())) {
        return "";
    }

    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
}
export function AuditTimeline(props: AuditTimelineContainerProps): ReactElement {
    const [searchText, setSearchText] = useState("");
    const [columnFilters, setColumnFilters] = useState({
    date: "",
    object: "",
    user: "",
    action: "",
    attribute: "",
    oldValue: "",
    newValue: "",
    description: ""
});

const updateColumnFilter = (
    column: keyof typeof columnFilters,
    value: string
) => {
    setColumnFilters(previous => ({
        ...previous,
        [column]: value
    }));

    setCurrentPage(1);
};
    const [viewMode, setViewMode] = useState<"timeline" | "grid">("timeline");
    const [sortColumn, setSortColumn] = useState<
    "date" | "object" | "user" | "action" | "attribute"
>("date");

const [sortDirection, setSortDirection] = useState<
    "asc" | "desc"
>("desc");
    const [currentPage, setCurrentPage] = useState(1);
    const [pageSize, setPageSize] = useState(10);
    const [actionFilter, setActionFilter] = useState("All");
    const [userFilter, setUserFilter] = useState("All");
    const [fromDate, setFromDate] = useState("");
    const [toDate, setToDate] = useState("");
    const [expandedItems, setExpandedItems] = useState<Set<string>>(
    new Set()
);

useEffect(() => {
    const items = props.dataSource.items ?? [];

    setExpandedItems(
        new Set(items.map(item => item.id))
    );
}, [props.dataSource.items]);
    const toggleExpanded = (itemId: string) => {
    setExpandedItems(previous => {
        const next = new Set(previous);

        if (next.has(itemId)) {
            next.delete(itemId);
        } else {
            next.add(itemId);
        }

        return next;
    });
};
const handleSort = (
    column: "date" | "object" | "user" | "action" | "attribute"
) => {
    if (sortColumn === column) {
        setSortDirection(previous =>
            previous === "asc" ? "desc" : "asc"
        );
    } else {
        setSortColumn(column);
        setSortDirection("asc");
    }
};
const goToPreviousPage = () => {
    setCurrentPage(previous =>
        Math.max(previous - 1, 1)
    );
};

const goToNextPage = () => {
    setCurrentPage(previous =>
        Math.min(previous + 1, totalPages)
    );
};
    const items = [...(props.dataSource.items ?? [])].sort((a, b) => {
    const dateA = props.changedDateTime.get(a).value?.getTime() ?? 0;
    const dateB = props.changedDateTime.get(b).value?.getTime() ?? 0;

    return dateB - dateA;
});
const normalizedSearchText = searchText.trim().toLowerCase();
const actionOptions = Array.from(
    new Set(
        items
            .map(item => props.action.get(item).value ?? "")
            .filter(action => action.trim() !== "")
    )
).sort();
const userOptions = Array.from(
    new Set(
        items
            .map(item => props.user.get(item).value ?? "")
            .filter(user => user.trim() !== "")
    )
).sort();
const filteredItems = items.filter(item => {
    const action = props.action.get(item).value ?? "";

    const matchesSearch =
        !normalizedSearchText ||
        [
            action,
            props.user.get(item).value ?? "",
            props.attributeName.get(item).value ?? "",
            props.oldValue.get(item).value ?? "",
            props.newValue.get(item).value ?? "",
            props.description.get(item).value ?? ""
        ].some(value =>
            value.toLowerCase().includes(normalizedSearchText)
        );

    const objectName =
        props.objectName.get(item).value ?? "";

    const matchesObject =
    !columnFilters.object.trim() ||
    objectName
        .toLowerCase()
        .includes(columnFilters.object.trim().toLowerCase());

    const matchesAction =
        actionFilter === "All" ||
        action.toLowerCase().includes(actionFilter.toLowerCase());

    const user =
        props.user.get(item).value ?? "";

    const matchesUser =
        userFilter === "All" ||
        user.toLowerCase().includes(userFilter.toLowerCase());

    const changedDateTime =
        props.changedDateTime.get(item).value;

    const recordDate = changedDateTime
        ? new Date(changedDateTime)
        : null;

    const matchesFromDate =
        !fromDate ||
        (recordDate !== null &&
            recordDate >= new Date(`${fromDate}T00:00:00`));

    const matchesToDate =
        !toDate ||
        (recordDate !== null &&
            recordDate <= new Date(`${toDate}T23:59:59`));
    const matchesColumnDate =
    !columnFilters.date ||
    (
        recordDate !== null &&
        toLocalDateInput(recordDate) === columnFilters.date
    );

    const matchesColumnObject =
        !columnFilters.object ||
        objectName.toLowerCase().includes(
        columnFilters.object.trim().toLowerCase()
    );

    const matchesColumnUser =
        !columnFilters.user ||
        user.toLowerCase().includes(
        columnFilters.user.trim().toLowerCase()
    );

    const matchesColumnAction =
        !columnFilters.action ||
        action.toLowerCase().includes(
        columnFilters.action.trim().toLowerCase()
    );

    const attributeName =
        props.attributeName.get(item).value ?? "";

    const oldValue =
        props.oldValue.get(item).value ?? "";

    const newValue =
        props.newValue.get(item).value ?? "";

    const description =
        props.description.get(item).value ?? "";

    const matchesColumnAttribute =
        !columnFilters.attribute ||
        attributeName.toLowerCase().includes(
        columnFilters.attribute.trim().toLowerCase()
    );

    const matchesColumnOldValue =
        !columnFilters.oldValue ||
        oldValue.toLowerCase().includes(
        columnFilters.oldValue.trim().toLowerCase()
    );

    const matchesColumnNewValue =
        !columnFilters.newValue ||
        newValue.toLowerCase().includes(
        columnFilters.newValue.trim().toLowerCase()
    );

    const matchesColumnDescription =
        !columnFilters.description ||
        description.toLowerCase().includes(
        columnFilters.description.trim().toLowerCase()
    );

    return (
    matchesSearch &&
    matchesObject &&
    matchesAction &&
    matchesUser &&
    matchesFromDate &&
    matchesToDate &&
    matchesColumnDate &&
    matchesColumnObject &&
    matchesColumnUser &&
    matchesColumnAction &&
    matchesColumnAttribute &&
    matchesColumnOldValue &&
    matchesColumnNewValue &&
    matchesColumnDescription
);
});
const sortedItems = [...filteredItems].sort((a, b) => {
    let comparison = 0;

    switch (sortColumn) {
        case "date": {
            const dateA =
                props.changedDateTime.get(a).value?.getTime() ?? 0;

            const dateB =
                props.changedDateTime.get(b).value?.getTime() ?? 0;

            comparison = dateA - dateB;
            break;
        }

        case "object": {
            const valueA =
                props.objectName.get(a).value ?? "";

            const valueB =
                props.objectName.get(b).value ?? "";

            comparison = valueA.localeCompare(valueB);
            break;
        }

        case "user": {
            const valueA =
                props.user.get(a).value ?? "";

            const valueB =
                props.user.get(b).value ?? "";

            comparison = valueA.localeCompare(valueB);
            break;
        }

        case "action": {
            const valueA =
                props.action.get(a).value ?? "";

            const valueB =
                props.action.get(b).value ?? "";

            comparison = valueA.localeCompare(valueB);
            break;
        }

        case "attribute": {
            const valueA =
                props.attributeName.get(a).value ?? "";

            const valueB =
                props.attributeName.get(b).value ?? "";

            comparison = valueA.localeCompare(valueB);
            break;
        }
    }

    return sortDirection === "asc"
        ? comparison
        : -comparison;
});
const totalItems = sortedItems.length;

const totalPages = Math.max(
    1,
    Math.ceil(totalItems / pageSize)
);

const startIndex =
    (currentPage - 1) * pageSize;

const endIndex =
    Math.min(
        startIndex + pageSize,
        totalItems
    );

const paginatedItems =
    sortedItems.slice(startIndex, endIndex);
const exportToExcel = () => {
    const exportData = filteredItems.map(item => ({
        "Date & Time": formatDate(
            props.changedDateTime.get(item).value
        ),
        Object: props.objectName.get(item).value ?? "",
        User: props.user.get(item).value ?? "",
        Action: props.action.get(item).value ?? "",
        Attribute: props.attributeName.get(item).value ?? "",
        "Old Value": props.oldValue.get(item).value ?? "",
        "New Value": props.newValue.get(item).value ?? "",
        Description: props.description.get(item).value ?? ""
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        "Audit History"
    );

    const now = new Date();

const day = String(now.getDate()).padStart(2, "0");

const month = now.toLocaleString("en-GB", {
    month: "short"
});

const year = now.getFullYear();

let hours = now.getHours();
const minutes = String(now.getMinutes()).padStart(2, "0");
const period = hours >= 12 ? "PM" : "AM";

hours = hours % 12 || 12;

const fileName =
    `AuditHistory_${day}-${month}-${year}_` +
    `${String(hours).padStart(2, "0")}-${minutes}-${period}.xlsx`;

XLSX.writeFile(
    workbook,
    fileName
);
};

    if (items.length === 0) {
        return (
            <div className="audit-timeline">
                <div className="audit-timeline-title">
                    Audit History
                </div>

                <div className="audit-timeline-empty">
                    No audit history available.
                </div>
            </div>
        );
    }

    return (
        <div className="audit-timeline">
    <div className="audit-timeline-title">
        Audit History
    </div>
    <div className="audit-timeline-view-selector">
    <label htmlFor="audit-timeline-view">
        View:
    </label>

    <select
        id="audit-timeline-view"
        className="audit-timeline-view-select"
        value={viewMode}
        onChange={event =>
            setViewMode(event.target.value as "timeline" | "grid")
        }
    >
        <option value="timeline">Timeline View</option>
        <option value="grid">Grid View</option>
    </select>
</div>

<div className="audit-timeline-filters">
    <input
        type="text"
        className="audit-timeline-search"
        placeholder="Search audit history..."
        value={searchText}
        onChange={event => setSearchText(event.target.value)}
    />
    <select
    className="audit-timeline-action-filter"
    value={actionFilter}
    onChange={event => setActionFilter(event.target.value)}
>
    <option value="All">All Actions</option>

    {actionOptions.map(action => (
        <option key={action} value={action}>
            {action}
        </option>
    ))}
</select>
<select
    className="audit-timeline-user-filter"
    value={userFilter}
    onChange={event => setUserFilter(event.target.value)}
>
    <option value="All">All Users</option>

    {userOptions.map(user => (
        <option key={user} value={user}>
            {user}
        </option>
    ))}
</select>
<input
    type="date"
    className="audit-timeline-date-filter"
    value={fromDate}
    onChange={event => setFromDate(event.target.value)}
    title="From date"
/>
<input
    type="date"
    className="audit-timeline-date-filter"
    value={toDate}
    onChange={event => setToDate(event.target.value)}
    title="To date"
/>
<button
    type="button"
    className="audit-timeline-clear-filter"
    onClick={() => {
    setSearchText("");
    setColumnFilters({
        date: "",
        object: "",
        user: "",
        action: "",
        attribute: "",
        oldValue: "",
        newValue: "",
        description: ""
    });
    setActionFilter("All");
    setUserFilter("All");
    setFromDate("");
    setToDate("");
    setCurrentPage(1);
}}
>
    Clear Filters
</button>
</div>
{viewMode === "timeline" && (
    <div className="audit-timeline-result-count">
        Showing {filteredItems.length} of {items.length} audit records
    </div>
)}
<button
    type="button"
    className="audit-timeline-export-button"
    onClick={exportToExcel}
>
     Export to Excel <span className="audit-timeline-export-icon">⭳</span>
</button>
{viewMode === "grid" && (
    <div className="audit-timeline-pagination">
        <div className="audit-timeline-page-size">
            <label htmlFor="audit-timeline-page-size">
                Rows per page:
            </label>

            <select
                id="audit-timeline-page-size"
                value={pageSize}
                onChange={event => {
                    setPageSize(
                        Number(event.target.value)
                    );
                    setCurrentPage(1);
                }}
            >
                <option value={10}>10</option>
                <option value={20}>20</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
            </select>
        </div>

        <div className="audit-timeline-page-info">
            {totalItems === 0
                ? "0–0 of 0"
                : `${startIndex + 1}–${endIndex} of ${totalItems}`}
        </div>

        <div className="audit-timeline-page-controls">
            <button
                type="button"
                className="audit-timeline-page-button"
                onClick={goToPreviousPage}
                disabled={currentPage === 1}
                aria-label="Previous page"
            >
                ‹
            </button>

            <button
                type="button"
                className="audit-timeline-page-button"
                onClick={goToNextPage}
                disabled={currentPage >= totalPages}
                aria-label="Next page"
            >
                ›
            </button>
        </div>
    </div>
)}

    {viewMode === "grid" ? (
        <div className="audit-timeline-grid">
        <div className="audit-timeline-grid-header">
    <div
        onClick={() => handleSort("date")}
        className="audit-timeline-grid-sortable"
    >
        Date & Time
        {sortColumn === "date" && (
            <span>
                {sortDirection === "asc" ? " ↑" : " ↓"}
            </span>
        )}
    </div>

    <div
        onClick={() => handleSort("object")}
        className="audit-timeline-grid-sortable"
    >
        Object
        {sortColumn === "object" && (
            <span>
                {sortDirection === "asc" ? " ↑" : " ↓"}
            </span>
        )}
    </div>

    <div
        onClick={() => handleSort("user")}
        className="audit-timeline-grid-sortable"
    >
        User
        {sortColumn === "user" && (
            <span>
                {sortDirection === "asc" ? " ↑" : " ↓"}
            </span>
        )}
    </div>

    <div
        onClick={() => handleSort("action")}
        className="audit-timeline-grid-sortable"
    >
        Action
        {sortColumn === "action" && (
            <span>
                {sortDirection === "asc" ? " ↑" : " ↓"}
            </span>
        )}
    </div>

    <div
        onClick={() => handleSort("attribute")}
        className="audit-timeline-grid-sortable"
    >
        Attribute
        {sortColumn === "attribute" && (
            <span>
                {sortDirection === "asc" ? " ↑" : " ↓"}
            </span>
        )}
    </div>

    <div>Old Value</div>
    <div>New Value</div>
    <div>Description</div>
</div>


<div className="audit-timeline-grid-search-row">
    <div>
        <input
            type="date"
            className="audit-timeline-grid-search-input"
            aria-label="Filter by Date and Time"
            title="Filter by Date & Time"
            value={columnFilters.date}
            onChange={event =>
                updateColumnFilter("date", event.target.value)
            }
        />
    </div>

    <div>
        <input
            type="text"
            className="audit-timeline-grid-search-input"
            placeholder="Search Object..."
            aria-label="Filter by Object"
            value={columnFilters.object}
            onChange={event =>
                updateColumnFilter("object", event.target.value)
            }
        />
    </div>

    <div>
        <input
            type="text"
            className="audit-timeline-grid-search-input"
            placeholder="Search User..."
            aria-label="Filter by User"
            value={columnFilters.user}
            onChange={event =>
                updateColumnFilter("user", event.target.value)
            }
        />
    </div>

    <div>
        <input
            type="text"
            className="audit-timeline-grid-search-input"
            placeholder="Search Action..."
            aria-label="Filter by Action"
            value={columnFilters.action}
            onChange={event =>
                updateColumnFilter("action", event.target.value)
            }
        />
    </div>

    <div>
        <input
            type="text"
            className="audit-timeline-grid-search-input"
            placeholder="Search Attribute..."
            aria-label="Filter by Attribute"
            value={columnFilters.attribute}
            onChange={event =>
                updateColumnFilter("attribute", event.target.value)
            }
        />
    </div>

    <div>
        <input
            type="text"
            className="audit-timeline-grid-search-input"
            placeholder="Search Old Value..."
            aria-label="Filter by Old Value"
            value={columnFilters.oldValue}
            onChange={event =>
                updateColumnFilter("oldValue", event.target.value)
            }
        />
    </div>

    <div>
        <input
            type="text"
            className="audit-timeline-grid-search-input"
            placeholder="Search New Value..."
            aria-label="Filter by New Value"
            value={columnFilters.newValue}
            onChange={event =>
                updateColumnFilter("newValue", event.target.value)
            }
        />
    </div>

    <div>
        <input
            type="text"
            className="audit-timeline-grid-search-input"
            placeholder="Search Description..."
            aria-label="Filter by Description"
            value={columnFilters.description}
            onChange={event =>
                updateColumnFilter("description", event.target.value)
            }
        />
    </div>
</div>

        {paginatedItems.length === 0 ? (
            <div className="audit-timeline-grid-empty">
                No matching audit records found.
            </div>
        ) : (
            paginatedItems.map(item => {
            const action =
                props.action.get(item).value ?? "";

            const user =
                props.user.get(item).value ?? "";

            const changedDateTime =
                props.changedDateTime.get(item).value;

            const objectName =
                props.objectName.get(item).value ?? "";

            const attributeName =
                props.attributeName.get(item).value ?? "";

            const oldValue =
                props.oldValue.get(item).value ?? "";

            const newValue =
                props.newValue.get(item).value ?? "";

            const description =
                props.description.get(item).value ?? "";

            return (
                <div
                    key={item.id}
                    className="audit-timeline-grid-row"
                >
                    <div>
                        {formatDate(changedDateTime)}
                    </div>

                    <div>
                        {objectName || "—"}
                    </div>

                    <div>
                        {user || "—"}
                    </div>

                    <div>
                        {action || "—"}
                    </div>

                    <div>
                        {attributeName || "—"}
                    </div>

                    <div>
                        {oldValue || "—"}
                    </div>

                    <div>
                        {newValue || "—"}
                    </div>

                    <div>
                        {description || "—"}
                    </div>
                </div>
            );
        })
    )}
    </div>
) : (
    <div className="audit-timeline-list">
        {filteredItems.length === 0 ? (
            <div className="audit-timeline-empty">
                No matching audit history found.
            </div>
        ) : (
            sortedItems.map(item => {
            const action =
                props.action.get(item).value ?? "";

            const user =
                props.user.get(item).value ?? "";

            const changedDateTime =
                props.changedDateTime.get(item).value;

            const objectName =
                props.objectName.get(item).value ?? "";

            const attributeName =
                props.attributeName.get(item).value ?? "";

            const oldValue =
                props.oldValue.get(item).value ?? "";

            const newValue =
                props.newValue.get(item).value ?? "";

            const description =
                props.description.get(item).value ?? "";

            const actionClass =
                getActionClass(action);

            const actionIcon =
                getActionIcon(action);

            const isExpanded =
                expandedItems.has(item.id);

            return (
                <div
                    key={item.id}
                    className={`audit-timeline-item ${actionClass}`}
                    onClick={() =>
                        toggleExpanded(item.id)
                    }
                >
                    <div className="audit-timeline-marker">
                        {actionIcon}
                    </div>

                    <div className="audit-timeline-content">
                        <div className="audit-timeline-date">
                            {formatDate(changedDateTime)}
                        </div>

                        <div className="audit-timeline-user">
                            {user}
                        </div>

                        <div className="audit-timeline-action">
                            <span>{action}</span>

                            <span className="audit-timeline-expand-icon">
                                {isExpanded
                                    ? "▲"
                                    : "▼"}
                            </span>
                        </div>

                        {isExpanded && (
                            <>
                                {objectName && (
                                    <div className="audit-timeline-object">
                                        <span className="audit-timeline-object-label">
                                            Object:
                                        </span>

                                        <span className="audit-timeline-object-value">
                                            {objectName}
                                        </span>
                                    </div>
                                )}

                                {attributeName && (
                                    <div className="audit-timeline-change">
                                        <div className="audit-timeline-attribute">
                                            {attributeName}
                                        </div>

                                        <div className="audit-timeline-values">
                                            <span className="audit-timeline-old-value">
                                                {oldValue || "—"}
                                            </span>

                                            <span className="audit-timeline-arrow">
                                                →
                                            </span>

                                            <span className="audit-timeline-new-value">
                                                {newValue || "—"}
                                            </span>
                                        </div>
                                    </div>
                                )}

                                {description && (
                                    <div className="audit-timeline-description">
                                        {description}
                                    </div>
                                )}
                            </>
                        )}
                     </div>
                </div>
            );
            })
        )}
    </div>
)}
</div>
);
}
