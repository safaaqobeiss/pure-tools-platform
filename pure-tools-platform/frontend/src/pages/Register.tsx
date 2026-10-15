import { useState } from "react";
import { Form, Input, Button, Typography, message, Row, Col } from "antd";
import {
  MailOutlined,
  LockOutlined,
  UserOutlined,
  CheckCircleOutlined,
} from "@ant-design/icons";
import { useAuth } from "../context/AuthContext";
import { useNavigate, Link } from "react-router-dom";
import PasswordStrength from "../components/PasswordStrength";
import SbsLogo from "../components/SbsLogo";

const { Title, Text, Paragraph } = Typography;

export default function Register() {
  const { register } = useAuth();
  const navigate = useNavigate();
  const [loading, setLoading] = useState(false);
  const [password, setPassword] = useState("");

  const onFinish = async (values: {
    email: string;
    password: string;
    full_name: string;
  }) => {
    setLoading(true);
    try {
      await register(values.email, values.password, values.full_name);
      message.success("Compte créé avec succès !");
      navigate("/login");
    } catch (err: any) {
      const detail = err.response?.data?.detail || "Erreur lors de l'inscription";
      if (Array.isArray(detail)) {
        message.error(detail.map((d: any) => d.msg).join(", "));
      } else {
        message.error(detail);
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <Row style={{ minHeight: "100vh" }}>
      {/* Left panel */}
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
            Rejoignez<br />la plateforme.
          </Title>

          <Paragraph style={{ color: "rgba(255,255,255,0.75)", fontSize: 16, lineHeight: 1.7, maxWidth: 400 }}>
            Créez votre compte et accédez immédiatement à tous les outils paramétriques de l'équipe.
          </Paragraph>

          <div style={{ marginTop: 40, display: "flex", flexDirection: "column", gap: 16 }}>
            {[
              "Accès immédiat à tous les outils",
              "Historique de vos exécutions",
              "Tableau de bord personnalisé",
              "Notifications en temps réel",
            ].map((text, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, color: "rgba(255,255,255,0.85)" }}>
                <div style={{ width: 36, height: 36, borderRadius: 10, background: "rgba(255,255,255,0.1)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>
                  <CheckCircleOutlined style={{ color: "#fff" }} />
                </div>
                <Text style={{ color: "rgba(255,255,255,0.85)", fontSize: 14 }}>{text}</Text>
              </div>
            ))}
          </div>
        </div>
      </Col>

      {/* Right panel */}
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
          <div style={{ textAlign: "center", marginBottom: 28 }}>
            <SbsLogo size={52} />
            <Title level={2} style={{ margin: "8px 0 4px", color: "#ba68c8" }}>Créer un compte</Title>
            <Text type="secondary">Remplissez les informations ci-dessous</Text>
          </div>

          <Form layout="vertical" onFinish={onFinish} size="large">
            <Form.Item
              name="full_name"
              rules={[{ required: true, message: "Veuillez saisir votre nom complet" }]}
            >
              <Input
                prefix={<UserOutlined style={{ color: "#ba68c8" }} />}
                placeholder="Nom complet"
                style={{ borderRadius: 10, height: 46 }}
              />
            </Form.Item>

            <Form.Item
              name="email"
              rules={[
                { required: true, message: "Veuillez saisir votre email" },
                { type: "email", message: "Email invalide" },
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
              rules={[
                { required: true, message: "Veuillez saisir un mot de passe" },
                {
                  validator(_, value) {
                    if (!value) return Promise.resolve();
                    if (value.length < 8) return Promise.reject(new Error("Min 8 caractères"));
                    if (!/[A-Z]/.test(value)) return Promise.reject(new Error("Au moins une majuscule"));
                    if (!/\d/.test(value)) return Promise.reject(new Error("Au moins un chiffre"));
                    if (!/[^A-Za-z0-9]/.test(value)) return Promise.reject(new Error("Au moins un caractère spécial"));
                    return Promise.resolve();
                  },
                },
              ]}
            >
              <Input.Password
                prefix={<LockOutlined style={{ color: "#ba68c8" }} />}
                placeholder="Mot de passe"
                style={{ borderRadius: 10, height: 46 }}
                onChange={(e) => setPassword(e.target.value)}
              />
            </Form.Item>

            <PasswordStrength password={password} />

            <Form.Item
              name="confirm"
              dependencies={["password"]}
              rules={[
                { required: true, message: "Veuillez confirmer le mot de passe" },
                ({ getFieldValue }) => ({
                  validator(_, value) {
                    if (!value || getFieldValue("password") === value) {
                      return Promise.resolve();
                    }
                    return Promise.reject(new Error("Les mots de passe ne correspondent pas"));
                  },
                }),
              ]}
            >
              <Input.Password
                prefix={<LockOutlined style={{ color: "#ba68c8" }} />}
                placeholder="Confirmer le mot de passe"
                style={{ borderRadius: 10, height: 46 }}
              />
            </Form.Item>

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
                S'inscrire
              </Button>
            </Form.Item>
          </Form>

          <div style={{ textAlign: "center" }}>
            <Text type="secondary">
              Déjà un compte ?{" "}
              <Link to="/login" style={{ fontWeight: 600, color: "#ba68c8" }}>Se connecter</Link>
            </Text>
          </div>
        </div>
      </Col>
    </Row>
  );
}