import { useState, useEffect } from "react";
import { Form, Input, Button, Typography, message, Row, Col } from "antd";
import {
  MailOutlined,
  LockOutlined,
  ToolOutlined,
  SafetyCertificateOutlined,
  RocketOutlined,
} from "@ant-design/icons";
import { useAuth } from "../context/AuthContext";
import { useLanguage } from "../context/LanguageContext";
import { useNavigate, Link } from "react-router-dom";
import SbsLogo from "../components/SbsLogo";

const { Title, Text, Paragraph } = Typography;

export default function Login() {
  const { login, isAuthenticated } = useAuth();
  const { t } = useLanguage();
  const navigate = useNavigate();
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (isAuthenticated) {
      navigate("/home", { replace: true });
    }
  }, [isAuthenticated, navigate]);

  const onFinish = async (values: { email: string; password: string }) => {
    setLoading(true);
    try {
      await login(values.email, values.password);
      message.success(t("login", "successMsg"));
    } catch {
      message.error(t("login", "errorMsg"));
    } finally {
      setLoading(false);
    }
  };

  return (
    <Row style={{ minHeight: "100vh" }}>
      {/* Left panel — branding */}
      <Col
        xs={0}
        md={12}
        style={{
          background: "linear-gradient(135deg, #c2185b 0%, #e91e90 50%, #ff9800 100%)",
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
          padding: "60px 50px",
          position: "relative",
          overflow: "hidden",
        }}
      >
        {/* Decorative circles */}
        <div style={{ position: "absolute", top: -80, right: -80, width: 250, height: 250, borderRadius: "50%", background: "rgba(255,255,255,0.05)" }} />
        <div style={{ position: "absolute", bottom: -60, left: -60, width: 200, height: 200, borderRadius: "50%", background: "rgba(255,255,255,0.04)" }} />

        <div style={{ position: "relative", zIndex: 1 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 40 }}>
            <SbsLogo size={50} />
            <Title level={3} style={{ color: "#fff", margin: 0 }}>Pure Tools</Title>
          </div>

          <Title level={1} style={{ color: "#fff", fontWeight: 700, lineHeight: 1.2, marginBottom: 20 }}>
            {t("login", "brandTitle").split("\n").map((line, i) => (
              <span key={i}>{line}<br /></span>
            ))}
          </Title>

          <Paragraph style={{ color: "rgba(255,255,255,0.75)", fontSize: 16, lineHeight: 1.7, maxWidth: 400 }}>
            {t("login", "brandDesc")}
          </Paragraph>

          <div style={{ marginTop: 40, display: "flex", flexDirection: "column", gap: 16 }}>
            {[
              { icon: <ToolOutlined />, text: t("login", "feature1") },
              { icon: <SafetyCertificateOutlined />, text: t("login", "feature2") },
              { icon: <RocketOutlined />, text: t("login", "feature3") },
            ].map((item, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, color: "rgba(255,255,255,0.85)" }}>
                <div style={{ width: 36, height: 36, borderRadius: 10, background: "rgba(255,255,255,0.1)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>
                  {item.icon}
                </div>
                <Text style={{ color: "rgba(255,255,255,0.85)", fontSize: 14 }}>{item.text}</Text>
              </div>
            ))}
          </div>
        </div>
      </Col>

      {/* Right panel — form */}
      <Col
        xs={24}
        md={12}
        style={{
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
          alignItems: "center",
          padding: "40px 32px",
          background: "#ffffff",
        }}
      >
        <div style={{ width: "100%", maxWidth: 380 }}>
          {/* Mobile logo */}
          <div style={{ textAlign: "center", marginBottom: 32 }}>
            <SbsLogo size={52} />
            <Title level={2} style={{ margin: "8px 0 4px", color: "#ba68c8" }}>{t("login", "title")}</Title>
            <Text type="secondary">{t("login", "subtitle")}</Text>
          </div>

          <Form layout="vertical" onFinish={onFinish} size="large">
            <Form.Item
              name="email"
              rules={[
                { required: true, message: t("login", "emailRequired") },
                { type: "email", message: t("login", "emailInvalid") },
              ]}
            >
              <Input
                prefix={<MailOutlined style={{ color: "#ba68c8" }} />}
                placeholder="Email"
                style={{ borderRadius: 10, height: 46 }}
              />
            </Form.Item>

            <Form.Item
              name="password"
              rules={[{ required: true, message: t("login", "passwordRequired") }]}
            >
              <Input.Password
                prefix={<LockOutlined style={{ color: "#ba68c8" }} />}
                placeholder="Mot de passe"
                style={{ borderRadius: 10, height: 46 }}
              />
            </Form.Item>

            <div style={{ textAlign: "right", marginTop: -16, marginBottom: 16 }}>
              <Link to="/forgot-password" style={{ fontSize: 13 }}>{t("login", "forgotPassword")}</Link>
            </div>

            <Form.Item>
              <Button
                type="primary"
                htmlType="submit"
                loading={loading}
                block
                style={{
                  height: 46,
                  borderRadius: 10,
                  fontWeight: 600,
                  fontSize: 15,
                  background: "linear-gradient(135deg, #e91e90, #ff9800)",
                  border: "none",
                }}
              >
                {t("login", "submit")}
              </Button>
            </Form.Item>
          </Form>

          <div style={{ textAlign: "center", marginTop: 16 }}>
            <Text type="secondary">
              {t("login", "noAccount")}{" "}
              <Link to="/register" style={{ fontWeight: 600, color: "#ba68c8" }}>{t("login", "createAccount")}</Link>
            </Text>
          </div>
        </div>
      </Col>
    </Row>
  );
}
