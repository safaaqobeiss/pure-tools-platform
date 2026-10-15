import { Progress, Typography } from "antd";
import { useMemo } from "react";

interface Props {
  password: string;
}

export default function PasswordStrength({ password }: Props) {
  const { score, label, color } = useMemo(() => {
    if (!password) return { score: 0, label: "", color: "#d9d9d9" };

    let s = 0;
    if (password.length >= 8) s++;
    if (password.length >= 12) s++;
    if (/[A-Z]/.test(password)) s++;
    if (/[0-9]/.test(password)) s++;
    if (/[^A-Za-z0-9]/.test(password)) s++;

    const levels: Record<number, { label: string; color: string }> = {
      0: { label: "Très faible", color: "#ff4d4f" },
      1: { label: "Faible", color: "#ff4d4f" },
      2: { label: "Moyen", color: "#faad14" },
      3: { label: "Bon", color: "#faad14" },
      4: { label: "Fort", color: "#52c41a" },
      5: { label: "Très fort", color: "#389e0d" },
    };

    const level = levels[s];
    return { score: (s / 5) * 100, label: level.label, color: level.color };
  }, [password]);

  if (!password) return null;

  return (
    <div style={{ marginTop: -16, marginBottom: 8 }}>
      <Progress
        percent={score}
        showInfo={false}
        strokeColor={color}
        size="small"
      />
      <Typography.Text style={{ fontSize: 12, color }}>
        {label}
      </Typography.Text>
      <div style={{ fontSize: 11, color: "#999", marginTop: 2 }}>
        Min 8 caractères, 1 majuscule, 1 chiffre, 1 caractère spécial
      </div>
    </div>
  );
}
