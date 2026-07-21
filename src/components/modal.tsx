/**
 * FinOpz-styled modal primitives — replicate the FinOpz `BaseModal` /
 * `ConfirmationModal` design (rounded-2xl card, close top-right, lg title + sm
 * description, full-width secondary + primary-black actions). The standalone
 * package can't import FinOpz's React-Aria component, so the look is matched here.
 */
import { type CSSProperties, type ReactNode } from "react";
import { createPortal } from "react-dom";
import { Z_BASE } from "../core/z-index";
import { X } from "@untitledui/icons";

const overlay: CSSProperties = {
  position: "fixed",
  inset: 0,
  background: "rgba(16,24,40,0.5)",
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  zIndex: Z_BASE + 2000,
  padding: 16,
};
const card: CSSProperties = {
  position: "relative",
  width: "100%",
  background: "#fff",
  borderRadius: 16, // rounded-2xl
  padding: 24, // p-6
  boxShadow: "0 20px 24px -4px rgba(16,24,40,0.10), 0 8px 8px -4px rgba(16,24,40,0.04)",
  border: "1px solid #eaecf0", // ring-1 ring-border-secondary
  maxHeight: "90vh",
  overflowY: "auto",
};
const closeBtn: CSSProperties = {
  position: "absolute",
  top: 16,
  right: 16,
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  width: 32,
  height: 32,
  borderRadius: 8,
  border: "none",
  background: "transparent",
  color: "#667085",
  cursor: "pointer",
};

const MAX_WIDTH: Record<"sm" | "md" | "lg", number> = { sm: 384, md: 448, lg: 512 };

export type ButtonVariant = "primary" | "secondary";

export function Button({
  variant = "secondary",
  onClick,
  disabled,
  fullWidth,
  children,
}: {
  variant?: ButtonVariant;
  onClick?: () => void;
  disabled?: boolean;
  fullWidth?: boolean;
  children: ReactNode;
}) {
  const base: CSSProperties = {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "10px 16px",
    borderRadius: 8,
    fontSize: 14,
    fontWeight: 600,
    cursor: disabled ? "not-allowed" : "pointer",
    opacity: disabled ? 0.5 : 1,
    flex: fullWidth ? 1 : undefined,
    transition: "background-color .12s ease",
  };
  const styles: CSSProperties =
    variant === "primary"
      ? { ...base, border: "none", background: "#0a0a0a", color: "#fff" } // primary-black
      : { ...base, border: "1px solid #d0d5dd", background: "#fff", color: "#344054" };
  return (
    <button type="button" disabled={disabled} onClick={onClick} style={styles}>
      {children}
    </button>
  );
}

export function Modal({
  open,
  title,
  description,
  onClose,
  children,
  footer,
  size = "md",
  showClose = true,
}: {
  open: boolean;
  title: ReactNode;
  description?: ReactNode;
  onClose: () => void;
  children?: ReactNode;
  footer?: ReactNode;
  size?: "sm" | "md" | "lg";
  showClose?: boolean;
}) {
  if (!open) return null;
  return createPortal(
    <div style={overlay} onMouseDown={onClose}>
      <div style={{ ...card, maxWidth: MAX_WIDTH[size] }} onMouseDown={(e) => e.stopPropagation()} role="dialog" aria-modal="true">
        {showClose && (
          <button type="button" aria-label="Close" onClick={onClose} style={closeBtn}>
            <X size={20} />
          </button>
        )}
        <h2 style={{ fontSize: 18, fontWeight: 600, color: "#101828", margin: 0, paddingRight: 32 }}>{title}</h2>
        {description && <p style={{ marginTop: 8, marginBottom: 0, fontSize: 14, color: "#667085", lineHeight: 1.5 }}>{description}</p>}
        {children}
        {footer && <div style={{ marginTop: 24, display: "flex", gap: 12 }}>{footer}</div>}
      </div>
    </div>,
    document.body,
  );
}

export function ConfirmModal({
  open,
  title,
  message,
  confirmLabel = "Confirm",
  cancelLabel = "Cancel",
  onConfirm,
  onClose,
}: {
  open: boolean;
  title: string;
  message: string;
  confirmLabel?: string;
  cancelLabel?: string;
  onConfirm: () => void;
  onClose: () => void;
}) {
  return (
    <Modal
      open={open}
      title={title}
      description={message}
      onClose={onClose}
      size="md"
      footer={
        <>
          <Button variant="secondary" fullWidth onClick={onClose}>
            {cancelLabel}
          </Button>
          <Button
            variant="primary"
            fullWidth
            onClick={() => {
              onConfirm();
              onClose();
            }}
          >
            {confirmLabel}
          </Button>
        </>
      }
    />
  );
}
