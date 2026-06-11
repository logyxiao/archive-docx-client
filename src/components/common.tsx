export function PathLine({ label, value }: { label: string; value: string }) {
  return (
    <div className="path-line">
      <span>{label}</span>
      <strong title={value}>{value}</strong>
    </div>
  );
}

export function PreviewItem({ label, value }: { label: string; value: string }) {
  return (
    <div className="preview-item">
      <span>{label}</span>
      <strong>{value || "/"}</strong>
    </div>
  );
}

export function Toggle({
  checked,
  onChange,
  label,
}: {
  checked: boolean;
  onChange: (checked: boolean) => void;
  label: string;
}) {
  return (
    <label className="toggle">
      <input type="checkbox" checked={checked} onChange={(event) => onChange(event.currentTarget.checked)} />
      <span>{label}</span>
    </label>
  );
}
