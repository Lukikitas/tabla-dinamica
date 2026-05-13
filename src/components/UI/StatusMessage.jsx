function StatusMessage({ tone = 'info', children }) {
  if (!children) return null;

  return <div className={`status-message status-${tone}`}>{children}</div>;
}

export default StatusMessage;
