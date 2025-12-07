// // components/LoadingOverlay.jsx
// import './LoadingOverlay.css';

// export default function LoadingOverlay({ show, text = "Loading..." }) {
//   if (!show) return null;
//   return (
//     <div className="overlay">
//       <div className="spinner-wrapper">
//         <div className="spinner" />
//       </div>
//       <div className="loading-text">{text}</div>
//     </div>
//   );
// }
// components/LoadingOverlay.jsx
export default function LoadingOverlay({ show }) {
  if (!show) return null;
  return (
    <div style={{
      position: 'fixed', inset: 0,
      background: 'rgba(0,0,0,0.5)',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      zIndex: 999
    }}>
      <div style={{
        width: '50px', height: '50px',
        border: '5px solid #ccc',
        borderTopColor: '#3b82f6',
        borderRadius: '50%',
        animation: 'spin 1s linear infinite'
      }} />
    </div>
  );
}