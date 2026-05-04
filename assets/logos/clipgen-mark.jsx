// clipgen mark — F.3 (stroke 13)
function StackedF({ size = 100, color = "currentColor" }) {
  const T = 13;
  return (
    <svg viewBox="0 0 100 100" width={size} height={size} fill="none" stroke={color} strokeWidth={T} strokeLinecap="square" strokeLinejoin="miter">
      <path d="M 18 18 L 82 18 L 82 82" />
      <path d="M 18 40 L 60 40 L 60 82" />
      <path d="M 18 62 L 38 62 L 38 82" />
    </svg>
  );
}

function StackedFAnimated({ size = 100, color = "currentColor", playKey = 0, onComplete }) {
  const T = 13;
  const path1Len = 64 + 64;
  const path2Len = 42 + 42;
  const path3Len = 20 + 20;

  const [stage, setStage] = React.useState(0);
  React.useEffect(() => {
    setStage(0);
    const t1 = setTimeout(() => setStage(1), 50);
    const t2 = setTimeout(() => setStage(2), 500);
    const t3 = setTimeout(() => setStage(3), 900);
    const t4 = setTimeout(() => { setStage(4); onComplete && onComplete(); }, 1400);
    return () => [t1, t2, t3, t4].forEach(clearTimeout);
  }, [playKey]);

  const drawStyle = (len, active) => ({
    strokeDasharray: len,
    strokeDashoffset: active ? 0 : len,
    transition: 'stroke-dashoffset 450ms cubic-bezier(0.65, 0, 0.35, 1)',
  });

  return (
    <svg viewBox="0 0 100 100" width={size} height={size} fill="none" stroke={color} strokeWidth={T} strokeLinecap="square" strokeLinejoin="miter">
      <path d="M 18 18 L 82 18 L 82 82" style={drawStyle(path1Len, stage >= 1)} />
      <path d="M 18 40 L 60 40 L 60 82" style={drawStyle(path2Len, stage >= 2)} />
      <path d="M 18 62 L 38 62 L 38 82" style={drawStyle(path3Len, stage >= 3)} />
    </svg>
  );
}

window.StackedF = StackedF;
window.StackedFAnimated = StackedFAnimated;
