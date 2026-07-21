import React, { useState, useEffect, useRef, memo } from "react";
import ChevronRightOutlinedIcon from "@mui/icons-material/ChevronRight";

/* --- Helper functions for Avatar --- */
export const getAvatarColorClass = (name = "") => {
  const colors = ["blue", "green", "orange", "purple", "pink", "teal"];
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = name.charCodeAt(i) + ((hash << 5) - hash);
  }
  return `avatar-bg-${colors[Math.abs(hash) % colors.length]}`;
};

export const getInitials = (name = "") => {
  if (!name) return "?";
  const clean = name.replace(/,.*$/, "").trim();
  const parts = clean.split(/\s+/);
  return parts.length >= 2
    ? (parts[0][0] + parts[1][0]).toUpperCase()
    : parts[0].substring(0, 2).toUpperCase();
};

/* --- 4. Reusable JSX Components --- */

export const IOSLoading = memo(() => <div className="ios-loading-spinner" />);

export const IOSButton = memo(({
  children,
  onClick,
  variant = "primary",
  disabled = false,
  loading = false,
  style,
  ariaLabel,
  type = "button",
  className = ""
}) => (
  <button
    type={type}
    onClick={onClick}
    style={style}
    aria-label={ariaLabel}
    className={`ios-btn ios-btn-${variant} ${className}`.trim()}
    disabled={disabled || loading}
  >
    {loading ? <IOSLoading /> : children}
  </button>
));

export const IOSCard = memo(({ children, interactive = false, style, onClick, className = "" }) => (
  <div
    className={`ios-card ${interactive ? "interactive" : ""} ${className}`.trim()}
    style={style}
    onClick={onClick}
  >
    {children}
  </div>
));

export const IOSSection = memo(({ children, title, footer, className = "" }) => (
  <div className={`ios-section ${className}`.trim()}>
    {title && <div className="ios-section-header">{title}</div>}
    {children}
    {footer && <div className="ios-section-footer">{footer}</div>}
  </div>
));

export const IOSList = memo(({ children, className = "", style }) => (
  <div className={`ios-list ${className}`.trim()} style={style}>
    {children}
  </div>
));

export const IOSListRow = memo(({
  children,
  onClick,
  interactive = false,
  rightContent,
  chevron = false,
  className = ""
}) => (
  <div
    className={`ios-list-row ${interactive || onClick ? "interactive" : ""} ${className}`.trim()}
    onClick={onClick}
  >
    <div className="ios-list-row-left">{children}</div>
    <div className="ios-list-row-right">
      {rightContent}
      {chevron && <ChevronRightOutlinedIcon className="ios-chevron" />}
    </div>
  </div>
));

export const IOSInput = memo(({
  type = "text",
  value,
  onChange,
  placeholder,
  select = false,
  options = [],
  style,
  ariaLabel,
  maxLength,
  inputRef,
  className = "",
  rows
}) => {
  if (select) {
    return (
      <select
        value={value}
        onChange={onChange}
        className={`ios-input ${className}`.trim()}
        style={style}
        aria-label={ariaLabel}
      >
        {options.map(o => (
          <option key={o.value} value={o.value}>
            {o.label}
          </option>
        ))}
      </select>
    );
  }
  if (type === "textarea") {
    return (
      <textarea
        ref={inputRef}
        rows={rows || 3}
        value={value}
        onChange={onChange}
        placeholder={placeholder}
        className={`ios-input ${className}`.trim()}
        style={style}
        aria-label={ariaLabel}
        maxLength={maxLength}
      />
    );
  }
  return (
    <input
      ref={inputRef}
      type={type}
      value={value}
      onChange={onChange}
      placeholder={placeholder}
      className={`ios-input ${className}`.trim()}
      style={style}
      aria-label={ariaLabel}
      maxLength={maxLength}
    />
  );
});

export const AppleSelect = memo(({ value, onChange, options, style, ariaLabel, className = "" }) => {
  const [isOpen, setIsOpen] = useState(false);
  const ref = useRef(null);
  const selected = options.find(o => o.value === value) || options[0];

  useEffect(() => {
    const handler = (e) => {
      if (ref.current && !ref.current.contains(e.target)) setIsOpen(false);
    };
    if (isOpen) document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [isOpen]);

  return (
    <div className={`apple-select-container ${className}`.trim()} ref={ref} style={style}>
      <button
        type="button"
        className="apple-select-trigger"
        onClick={() => setIsOpen(!isOpen)}
        aria-haspopup="listbox"
        aria-expanded={isOpen}
        aria-label={ariaLabel}
      >
        <span>{selected?.label ?? ""}</span>
        <ChevronRightOutlinedIcon className={`apple-select-chevron ${isOpen ? "open" : ""}`} />
      </button>
      {isOpen && (
        <div className="apple-select-dropdown" role="listbox">
          <div style={{ maxHeight: "200px", overflowY: "auto" }}>
            {options.map(o => (
              <div
                key={o.value}
                role="option"
                aria-selected={o.value === value}
                className={`apple-select-option ${o.value === value ? "selected" : ""}`}
                onClick={() => {
                  onChange({ target: { value: o.value } });
                  setIsOpen(false);
                }}
              >
                {o.label}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
});

export const IOSSwitch = memo(({ checked, onChange, ariaLabel }) => (
  <label className="ios-switch" aria-label={ariaLabel}>
    <input type="checkbox" checked={checked} onChange={onChange} />
    <span className="ios-switch-slider" />
  </label>
));

export const IOSBadge = memo(({ status = "BELUM", children }) => {
  const map = {
    HADIR: "hadir",
    IZIN: "izin",
    SAKIT: "sakit",
    ALPHA: "alpha",
    LIBUR: "libur",
    BELUM: "belum"
  };
  const key = status ? status.toUpperCase() : "BELUM";
  return (
    <span className={`ios-badge ios-badge-${map[key] ?? "belum"}`}>
      {children || status}
    </span>
  );
});

export const IOSAvatar = memo(({ name = "", large = false, className = "" }) => {
  const colorClass = getAvatarColorClass(name);
  const initials = getInitials(name);
  return (
    <div className={`ios-avatar ${large ? "ios-avatar-large" : ""} ${colorClass} ${className}`.trim()}>
      {initials}
    </div>
  );
});

export const IOSSkeleton = memo(({ height = "20px", width = "100%", style, className = "" }) => (
  <div className={`ios-skeleton ${className}`.trim()} style={{ height, width, ...style }} />
));

export const IOSEmptyState = memo(({ icon, title, description, action }) => (
  <div className="ios-empty-state">
    {icon && <div className="ios-empty-state-icon">{icon}</div>}
    {title && <h3>{title}</h3>}
    {description && <p>{description}</p>}
    {action}
  </div>
));

export const IOSSheet = memo(({ children, isOpen, onClose, className = "" }) => {
  if (!isOpen) return null;
  return (
    <div className={`ios-sheet-overlay ${className ? className + "-overlay" : ""}`} onClick={onClose}>
      <div className={`ios-sheet ${className}`} onClick={e => e.stopPropagation()} role="dialog" tabIndex={-1}>
        <div className="ios-sheet-grabber" />
        <div style={{ flex: 1, display: "flex", flexDirection: "column", overflowY: "auto" }}>
          {children}
        </div>
      </div>
    </div>
  );
});

export const IOSAlert = memo(({ isOpen, title, description, actions = [] }) => {
  if (!isOpen) return null;
  return (
    <div className="ios-alert-overlay" role="alertdialog">
      <div className="ios-alert">
        <div className="ios-alert-content">
          <h3>{title}</h3>
          {description && <p>{description}</p>}
        </div>
        <div className="ios-alert-actions">
          {actions.map((act, i) => (
            <button
              key={i}
              onClick={act.onClick}
              className={`ios-alert-action-btn ${act.bold ? "bold" : ""} ${act.destructive ? "destructive" : ""}`}
            >
              {act.label}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
});

export const IOSSegmentedControl = memo(({ segments, selectedValue, onChange, disabled = false }) => (
  <div className="ios-segmented-control" role="radiogroup">
    {segments.map(seg => (
      <button
        key={seg.value}
        type="button"
        role="radio"
        disabled={disabled}
        aria-checked={seg.value === selectedValue}
        onClick={() => onChange(seg.value)}
        className={`ios-segmented-segment ${seg.cls || ""} ${seg.value === selectedValue ? "selected" : ""}`}
      >
        {seg.label}
      </button>
    ))}
  </div>
));

export const ToastContainer = memo(({ toasts = [] }) => (
  <div className="toast-container">
    {toasts.map(t => (
      <div key={t.id} className={`toast ${t.exiting ? "exiting" : ""}`}>
        <span>{t.message}</span>
      </div>
    ))}
  </div>
));
