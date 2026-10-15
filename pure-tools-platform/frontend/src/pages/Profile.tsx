import { useState } from "react";
import { useNavigate } from "react-router-dom";
import {
  Layout,
  Typography,
  Button,
  Form,
  Input,
  Upload,
  Avatar,
  message,
  Row,
  Col,
  Tabs,
} from "antd";
import {
  ArrowLeftOutlined,
  UserOutlined,
  CameraOutlined,
  SaveOutlined,
  LockOutlined,
  MailOutlined,
  EditOutlined,
  SafetyCertificateOutlined,
  CalendarOutlined,
  ThunderboltOutlined,
} from "@ant-design/icons";
import type { UploadFile } from "antd";
import api from "../api/axios";
import { useAuth } from "../context/AuthContext";
import { useLanguage } from "../context/LanguageContext";
import Navbar from "../components/Navbar";

const { Content } = Layout;
const { Title, Text } = Typography;

export default function Profile() {
  const navigate = useNavigate();
  const { user, refreshUser } = useAuth();
  const { t } = useLanguage();
  const [nameForm] = Form.useForm();
  const [passwordForm] = Form.useForm();
  const [savingName, setSavingName] = useState(false);
  const [savingPassword, setSavingPassword] = useState(false);
  const [uploadingAvatar, setUploadingAvatar] = useState(false);
  const [avatarKey, setAvatarKey] = useState(0);

  const avatarUrl = user?.avatar_url
    ? `${api.defaults.baseURL || ""}${user.avatar_url}?t=${avatarKey}`
    : null;

  const initials = user?.full_name
    ? user.full_name
        .split(" ")
        .map((w) => w[0])
        .join("")
        .toUpperCase()
        .slice(0, 2)
    : user?.email?.[0]?.toUpperCase() || "?";

  const handleSaveName = async () => {
    try {
      const values = await nameForm.validateFields();
      setSavingName(true);
      await api.put("/api/auth/profile", { full_name: values.full_name });
      await refreshUser();
      message.success(t("profile", "nameUpdated"));
    } catch {
      message.error(t("profile", "updateError"));
    } finally {
      setSavingName(false);
    }
  };

  const handleChangePassword = async () => {
    try {
      const values = await passwordForm.validateFields();
      if (values.new_password !== values.confirm_password) {
        message.error(t("profile", "passwordMismatch"));
        return;
      }
      setSavingPassword(true);
      await api.put("/api/auth/profile", {
        current_password: values.current_password,
        new_password: values.new_password,
      });
      passwordForm.resetFields();
      message.success(t("profile", "passwordUpdated"));
    } catch (err: any) {
      message.error(
        err.response?.data?.detail || t("profile", "passwordError")
      );
    } finally {
      setSavingPassword(false);
    }
  };

  const handleAvatarUpload = async (file: UploadFile) => {
    if (!file.originFileObj) return;
    setUploadingAvatar(true);
    try {
      const formData = new FormData();
      formData.append("file", file.originFileObj);
      await api.post("/api/auth/avatar", formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      await refreshUser();
      setAvatarKey((k) => k + 1);
      message.success(t("profile", "avatarUpdated"));
    } catch (err: any) {
      message.error(err.response?.data?.detail || t("profile", "avatarError"));
    } finally {
      setUploadingAvatar(false);
    }
  };

  const memberSince = user?.created_at
    ? new Date(user.created_at).toLocaleDateString("fr-FR", { year: "numeric", month: "long" })
    : null;

  return (
    <Layout style={{ minHeight: "100vh", background: "#ffffff" }}>
      <Navbar />
      <Content style={{ padding: 0 }}>

        <div style={{ maxWidth: 900, margin: "32px auto 0", padding: "0 24px", position: "relative", zIndex: 1 }}>
          {/* Back button */}
          <Button
            icon={<ArrowLeftOutlined />}
            type="text"
            onClick={() => navigate("/home")}
            style={{ color: "#a9516a", fontWeight: 500, marginBottom: 16, padding: 0 }}
          >
            {t("common", "backToTools")}
          </Button>

          {/* Avatar + Identity card */}
          <div
            style={{
              background: "#fff",
              borderRadius: 20,
              boxShadow: "0 8px 32px rgba(183,28,28,0.12)",
              padding: "0 40px 32px",
              marginBottom: 24,
            }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: 28, padding: "32px 0 0" }}>
              {/* Avatar with upload overlay */}
              <div style={{ position: "relative", flexShrink: 0 }}>
                <Upload
                  maxCount={1}
                  showUploadList={false}
                  beforeUpload={() => false}
                  accept="image/png,image/jpeg,image/webp,image/gif"
                  onChange={({ fileList }) => {
                    if (fileList.length > 0) handleAvatarUpload(fileList[0]);
                  }}
                >
                  <div style={{ position: "relative", cursor: "pointer" }}>
                    {avatarUrl ? (
                      <Avatar
                        size={120}
                        src={avatarUrl}
                        style={{
                          border: "4px solid #fff",
                          boxShadow: "0 4px 14px rgba(0,0,0,0.15)",
                        }}
                      />
                    ) : (
                      <Avatar
                        size={120}
                        style={{
                          background: "linear-gradient(135deg, #a9516a, #ff9800)",
                          fontWeight: 700,
                          border: "4px solid #fff",
                          boxShadow: "0 4px 14px rgba(0,0,0,0.15)",
                        }}
                      >
                        {initials}
                      </Avatar>
                    )}
                    <div
                      style={{
                        position: "absolute",
                        bottom: 4,
                        right: 4,
                        width: 34,
                        height: 34,
                        borderRadius: "50%",
                        background: "linear-gradient(135deg, #d77fa3, #ff9800)",
                        boxShadow: "0 2px 8px rgba(0,0,0,0.2)",
                      }}
                    >
                      {uploadingAvatar ? (
                        <ThunderboltOutlined spin style={{ color: "#fff", fontSize: 14 }} />
                      ) : (
                        <CameraOutlined style={{ color: "#fff", fontSize: 14 }} />
                      )}
                    </div>
                  </div>
                </Upload>
              </div>

              {/* Identity info */}
              <div style={{ flex: 1, paddingBottom: 8 }}>
                <Title level={2} style={{ margin: 0, color: "#a9516a", fontWeight: 700, lineHeight: 1.2 }}>
                  {user?.full_name || t("home", "user")}
                </Title>
                <div style={{ display: "flex", alignItems: "center", gap: 20, marginTop: 8, flexWrap: "wrap" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <MailOutlined style={{ color: "#8899a6" }} />
                    <Text style={{ color: "#5a6f80", fontSize: 14 }}>{user?.email}</Text>
                  </div>
                  {user?.role && (
                    <div
                      style={{
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 5,
                        padding: "2px 12px",
                        borderRadius: 20,
                        background: user.role === "admin"
                          ? "linear-gradient(135deg, rgba(183,28,28,0.1), rgba(216,67,21,0.1))"
                          : "rgba(229,57,53,0.08)",
                        border: `1px solid ${user.role === "admin" ? "rgba(183,28,28,0.2)" : "rgba(229,57,53,0.12)"}`,
                      }}
                    >
                      <SafetyCertificateOutlined style={{ color: "#9369b8", fontSize: 12 }} />
                      <Text style={{ color: "#9369b8", fontSize: 12, fontWeight: 600, textTransform: "capitalize" }}>
                        {user.role}
                      </Text>
                    </div>
                  )}
                  {memberSince && (
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <CalendarOutlined style={{ color: "#8899a6", fontSize: 13 }} />
                      <Text style={{ color: "#8899a6", fontSize: 13 }}>{t("profile", "memberSince")} {memberSince}</Text>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>

          {/* Settings tabs */}
          <div
            style={{
              background: "#fff",
              borderRadius: 20,
              boxShadow: "0 4px 20px rgba(183,28,28,0.08)",
              padding: "8px 0 32px",
              marginBottom: 40,
            }}
          >
            <Tabs
              defaultActiveKey="info"
              centered
              style={{ padding: "0 40px" }}
              items={[
                {
                  key: "info",
                  label: (
                    <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <EditOutlined /> {t("profile", "infoTab")}
                    </span>
                  ),
                  children: (
                    <div style={{ maxWidth: 480, margin: "24px auto 0" }}>
                      <Form
                        form={nameForm}
                        layout="vertical"
                        initialValues={{ full_name: user?.full_name || "" }}
                      >
                        <Form.Item
                          name="full_name"
                          label={
                            <Text strong style={{ color: "#a9516a" }}>
                              {t("profile", "fullName")}
                            </Text>
                          }
                          rules={[{ required: true, message: t("profile", "nameRequired") }]}
                        >
                          <Input
                            prefix={<UserOutlined style={{ color: "#8899a6" }} />}
                            placeholder={t("profile", "yourName")}
                            style={{ borderRadius: 10, height: 46 }}
                          />
                        </Form.Item>

                        <Form.Item
                          label={
                            <Text strong style={{ color: "#c2185b" }}>
                              {t("profile", "emailAddress")}
                            </Text>
                          }
                        >
                          <Input
                            prefix={<MailOutlined style={{ color: "#8899a6" }} />}
                            value={user?.email || ""}
                            disabled
                            style={{
                              borderRadius: 10,
                              height: 46,
                              background: "#f8fafc",
                              color: "#5a6f80",
                            }}
                          />
                        </Form.Item>

                        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8 }}>
                          <Button
                            type="primary"
                            icon={<SaveOutlined />}
                            onClick={handleSaveName}
                            loading={savingName}
                            style={{
                              height: 42,
                              borderRadius: 10,
                              fontWeight: 600,
                              paddingInline: 28,
                              background: "linear-gradient(135deg, #c2185b, #ff9800)",
                              border: "none",
                              boxShadow: "0 4px 12px rgba(183,28,28,0.25)",
                            }}
                          >
                            {t("common", "save")}
                          </Button>
                        </div>
                      </Form>
                    </div>
                  ),
                },
                {
                  key: "security",
                  label: (
                    <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <LockOutlined /> {t("profile", "securityTab")}
                    </span>
                  ),
                  children: (
                    <div style={{ maxWidth: 480, margin: "24px auto 0" }}>
                      <div style={{ marginBottom: 24 }}>
                        <Title level={5} style={{ color: "#a9516a", margin: "0 0 4px" }}>
                          {t("profile", "changePassword")}
                        </Title>
                        <Text style={{ color: "#8899a6", fontSize: 13 }}>
                          {t("profile", "passwordHint")}
                        </Text>
                      </div>

                      <Form form={passwordForm} layout="vertical">
                        <Form.Item
                          name="current_password"
                          label={
                              <Text strong style={{ color: "#c2185b" }}>
                              {t("profile", "currentPassword")}
                            </Text>
                          }
                          rules={[{ required: true, message: t("profile", "required") }]}
                        >
                          <Input.Password
                            prefix={<LockOutlined style={{ color: "#8899a6" }} />}
                            style={{ borderRadius: 10, height: 46 }}
                          />
                        </Form.Item>

                        <Row gutter={16}>
                          <Col span={12}>
                            <Form.Item
                              name="new_password"
                              label={
                                <Text strong style={{ color: "#c2185b" }}>
                                {t("profile", "newPassword")}
                              </Text>
                            }
                            rules={[
                              { required: true, message: t("profile", "required") },
                              { min: 6, message: t("profile", "minChars") },
                              ]}
                            >
                              <Input.Password
                                prefix={<LockOutlined style={{ color: "#8899a6" }} />}
                                style={{ borderRadius: 10, height: 46 }}
                              />
                            </Form.Item>
                          </Col>
                          <Col span={12}>
                            <Form.Item
                              name="confirm_password"
                              label={
                                <Text strong style={{ color: "#c2185b" }}>
                                {t("profile", "confirmPassword")}
                              </Text>
                            }
                            rules={[{ required: true, message: t("profile", "required") }]}
                            >
                              <Input.Password
                                prefix={<LockOutlined style={{ color: "#8899a6" }} />}
                                style={{ borderRadius: 10, height: 46 }}
                              />
                            </Form.Item>
                          </Col>
                        </Row>

                        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8 }}>
                          <Button
                            type="primary"
                            icon={<SafetyCertificateOutlined />}
                            onClick={handleChangePassword}
                            loading={savingPassword}
                            style={{
                              height: 42,
                              borderRadius: 10,
                              fontWeight: 600,
                              paddingInline: 28,
                              background: "linear-gradient(135deg, #e8590c, #d9480f)",
                              border: "none",
                              boxShadow: "0 4px 12px rgba(232,89,12,0.25)",
                            }}
                          >
                            {t("profile", "updatePassword")}
                          </Button>
                        </div>
                      </Form>
                    </div>
                  ),
                },
              ]}
            />
          </div>
        </div>
      </Content>
    </Layout>
  );
}
