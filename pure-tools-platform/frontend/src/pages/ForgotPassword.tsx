import { Typography, Card, Result } from "antd";
import { Link } from "react-router-dom";
import { InfoCircleOutlined } from "@ant-design/icons";
import { useLanguage } from "../context/LanguageContext";

export default function ForgotPassword() {
  const { t } = useLanguage();
  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        background: "linear-gradient(135deg, #c2185b 0%, #e91e90 100%)",
      }}
    >
      <Card style={{ width: 440, borderRadius: 12, boxShadow: "0 4px 20px rgba(0,0,0,0.15)" }}>
        <Result
          icon={<InfoCircleOutlined style={{ color: "#ba68c8" }} />}
          title={t("forgot", "title")}
          subTitle={t("forgot", "subtitle")}
        />
        <div style={{ textAlign: "center", marginTop: -8 }}>
          <Link to="/login">{t("forgot", "backToLogin")}</Link>
        </div>
      </Card>
    </div>
  );
}
