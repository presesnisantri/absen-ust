import React, { memo } from 'react';
import InfoOutlinedIcon from '@mui/icons-material/InfoOutlined';
import ErrorOutlineIcon from '@mui/icons-material/ErrorOutlineOutlined';
import WarningAmberOutlinedIcon from '@mui/icons-material/WarningAmberOutlined';
import CheckCircleOutlinedIcon from '@mui/icons-material/CheckCircleOutlined';

const typeToIcon = {
  info: <InfoOutlinedIcon style={{ fontSize: '1.2rem', marginRight: 'var(--space-8)' }} />, 
  success: <CheckCircleOutlinedIcon style={{ fontSize: '1.2rem', marginRight: 'var(--space-8)' }} />, 
  warning: <WarningAmberOutlinedIcon style={{ fontSize: '1.2rem', marginRight: 'var(--space-8)' }} />, 
  error: <ErrorOutlineIcon style={{ fontSize: '1.2rem', marginRight: 'var(--space-8)' }} />, 
};

const ToastContainer = memo(function ToastContainer({ toasts, onRemove }) {
  const activeToast = toasts[0];
  if (!activeToast) return null;

  return (
    <div className="toast-container" role="region" aria-live="polite">
      <div
        key={activeToast.id}
        className={`toast ${activeToast.exiting ? 'exiting' : ''}`}
        onAnimationEnd={() => {
          if (activeToast.exiting && onRemove) onRemove(activeToast.id);
        }}
      >
        {typeToIcon[activeToast.type] || typeToIcon.info}
        <span className="toast-message" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
          {activeToast.message}
        </span>
      </div>
    </div>
  );
});

export default ToastContainer;
