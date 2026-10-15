export default function SbsLogo({ size = 36 }: { size?: number }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 100 100"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      {/* Bottom-left diamond — cyan/blue */}
      <rect
        x="10"
        y="45"
        width="32"
        height="32"
        rx="3"
        transform="rotate(-45 10 45)"
        fill="url(#blueGrad)"
      />
      {/* Bottom-right diamond — magenta */}
      <rect
        x="38"
        y="45"
        width="32"
        height="32"
        rx="3"
        transform="rotate(-45 38 45)"
        fill="url(#magentaGrad)"
      />
      {/* Top diamond — yellow/orange */}
      <rect
        x="24"
        y="18"
        width="32"
        height="32"
        rx="3"
        transform="rotate(-45 24 18)"
        fill="url(#yellowGrad)"
      />
      <defs>
        <linearGradient id="blueGrad" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor="#ff9800" />
          <stop offset="100%" stopColor="#e91e90" />
        </linearGradient>
        <linearGradient id="magentaGrad" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor="#e91e90" />
          <stop offset="100%" stopColor="#9c27b0" />
        </linearGradient>
        <linearGradient id="yellowGrad" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor="#ffca28" />
          <stop offset="100%" stopColor="#ff9800" />
        </linearGradient>
      </defs>
    </svg>
  );
}
