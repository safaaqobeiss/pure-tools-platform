import { useEffect, useState } from "react";
import {
  Button,
  Card,
  Form,
  Input,
  Layout,
  message,
  Modal,
  Popconfirm,
  Select,
  Space,
  Switch,
  Table,
  Tag,
  Typography,
  Row,
  Col,
  Avatar,
  Tooltip,
} from "antd";
import {
  DeleteOutlined,
  EditOutlined,
  SettingOutlined,
  TeamOutlined,
  UserOutlined,
  SafetyCertificateOutlined,
  LockOutlined,
  CheckCircleOutlined,
  StopOutlined,
  CrownOutlined,
  SyncOutlined,
  SearchOutlined,
} from "@ant-design/icons";
import api from "../api/axios";
import Navbar from "../components/Navbar";
import { useLanguage } from "../context/LanguageContext";

const { Content } = Layout;
const { Title, Text } = Typography;

interface AdminUser {
  id: number;
  email: string;
  full_name: string | null;
  role: string;
  is_active: boolean;
}

export default function Admin() {
  const [users, setUsers] = useState<AdminUser[]>([]);
  const [loading, setLoading] = useState(true);
  const [editUser, setEditUser] = useState<AdminUser | null>(null);
  const [editModalOpen, setEditModalOpen] = useState(false);
  const [saving, setSaving] = useState(false);
  const [search, setSearch] = useState("");
  const [editForm] = Form.useForm();
  const { t } = useLanguage();

  const fetchUsers = () => {
    setLoading(true);
    api
      .get("/api/admin/users")
      .then((res) => setUsers(res.data))
      .catch(() => message.error(t("admin", "loadError")))
      .finally(() => setLoading(false));
  };

  useEffect(() => {
    fetchUsers();
  }, []);

  const updateUser = async (userId: number, data: Partial<AdminUser>) => {
    try {
      await api.put(`/api/admin/users/${userId}`, data);
      message.success(t("admin", "userUpdated"));
      fetchUsers();
    } catch (err: unknown) {
      const detail =
        (err as { response?: { data?: { detail?: string } } })?.response?.data
          ?.detail ?? t("common", "error");
      message.error(detail);
    }
  };

  const deleteUser = async (userId: number) => {
    try {
      await api.delete(`/api/admin/users/${userId}`);
      message.success(t("admin", "userDeleted"));
      fetchUsers();
    } catch (err: unknown) {
      const detail =
        (err as { response?: { data?: { detail?: string } } })?.response?.data
          ?.detail ?? "Erreur";
      message.error(detail);
    }
  };

  const openEdit = (record: AdminUser) => {
    setEditUser(record);
    editForm.setFieldsValue({
      email: record.email,
      full_name: record.full_name || "",
      role: record.role,
      is_active: record.is_active,
      new_password: "",
    });
    setEditModalOpen(true);
  };

  const handleEditSave = async () => {
    try {
      const values = await editForm.validateFields();
      if (!editUser) return;
      setSaving(true);
      const payload: Record<string, unknown> = {
        email: values.email,
        full_name: values.full_name,
        role: values.role,
        is_active: values.is_active,
      };
      if (values.new_password) {
        payload.new_password = values.new_password;
      }
      await api.put(`/api/admin/users/${editUser.id}`, payload);
      message.success(t("admin", "userEdited"));
      setEditModalOpen(false);
      fetchUsers();
    } catch (err: unknown) {
      const detail =
        (err as { response?: { data?: { detail?: string } } })?.response?.data
          ?.detail ?? "Erreur de validation";
      message.error(detail);
    } finally {
      setSaving(false);
    }
  };

  // Stats
  const totalUsers = users.length;
  const activeUsers = users.filter((u) => u.is_active).length;
  const adminUsers = users.filter((u) => u.role === "admin").length;
  const inactiveUsers = users.filter((u) => !u.is_active).length;

  // Filtered users
  const filteredUsers = search
    ? users.filter(
        (u) =>
          u.email.toLowerCase().includes(search.toLowerCase()) ||
          u.full_name?.toLowerCase().includes(search.toLowerCase())
      )
    : users;

  const statCards = [
    {
      label: t("admin", "total"),
      value: totalUsers,
      icon: <TeamOutlined style={{ fontSize: 22 }} />,
       gradient: "linear-gradient(135deg, #e91e90, #c2185b)",
      color: "#ba68c8",
    },
    {
      label: t("admin", "active"),
      value: activeUsers,
      icon: <CheckCircleOutlined style={{ fontSize: 22 }} />,
      gradient: "linear-gradient(135deg, #ff9800, #ffca28)",
      color: "#ff9800",
    },
    {
      label: t("admin", "admins"),
      value: adminUsers,
      icon: <CrownOutlined style={{ fontSize: 22 }} />,
      gradient: "linear-gradient(135deg, #ffca28, #ff9800)",
      color: "#ffca28",
    },
    {
      label: t("admin", "inactive"),
      value: inactiveUsers,
      icon: <StopOutlined style={{ fontSize: 22 }} />,
      gradient: "linear-gradient(135deg, #c2185b, #9c27b0)",
      color: "#c2185b",
    },
  ];

  const columns = [
    {
      title: t("admin", "userCol"),
      key: "user",
      render: (_: unknown, record: AdminUser) => (
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <Avatar
            size={38}
            style={{
              background: record.role === "admin"
                ? "linear-gradient(135deg, #ffca28, #ff9800)"
                : "linear-gradient(135deg, #e91e90, #c2185b)",
              fontWeight: 600,
              fontSize: 15,
            }}
          >
            {(record.full_name || record.email)[0].toUpperCase()}
          </Avatar>
          <div>
            <Text strong style={{ color: "#333", fontSize: 14 }}>
              {record.full_name || "—"}
            </Text>
            <br />
            <Text
              copyable={{ text: record.email, tooltips: [t("admin", "copy"), t("admin", "copied")] }}
              style={{ color: "#888", fontSize: 12 }}
            >
              {record.email}
            </Text>
          </div>
        </div>
      ),
    },
    {
      title: t("admin", "role"),
      dataIndex: "role",
      key: "role",
      width: 120,
      render: (role: string) =>
        role === "admin" ? (
          <Tag
            icon={<SafetyCertificateOutlined />}
            style={{
              background: "linear-gradient(135deg, #ffca28, #ff9800)",
              color: "#fff",
              border: "none",
              borderRadius: 8,
              fontWeight: 600,
              padding: "2px 12px",
            }}
          >
            Admin
          </Tag>
        ) : (
          <Tag
            icon={<UserOutlined />}
            style={{
              background: "linear-gradient(135deg, #e91e90, #c2185b)",
              color: "#fff",
              border: "none",
              borderRadius: 8,
              fontWeight: 600,
              padding: "2px 12px",
            }}
          >
            User
          </Tag>
        ),
    },
    {
      title: t("admin", "statusCol"),
      dataIndex: "is_active",
      key: "is_active",
      width: 110,
      render: (active: boolean, record: AdminUser) => (
        <Switch
          checked={active}
          checkedChildren={t("admin", "activeSwitch")}
          unCheckedChildren={t("admin", "inactiveSwitch")}
          onChange={(checked) => updateUser(record.id, { is_active: checked })}
          style={{ background: active ? "#11998e" : "#e74c3c" }}
        />
      ),
    },
    {
      title: t("admin", "actionsCol"),
      key: "actions",
      width: 120,
      render: (_: unknown, record: AdminUser) => (
        <Space>
          <Tooltip title={t("common", "edit")}>
            <Button
              icon={<EditOutlined />}
              size="small"
              onClick={() => openEdit(record)}
              style={{
                borderRadius: 8,
                background: "linear-gradient(135deg, #e91e90, #c2185b)",
                color: "#fff",
                border: "none",
              }}
            />
          </Tooltip>
          <Popconfirm
            title={t("admin", "deleteUser")}
            description={t("admin", "deleteConfirm")}
            onConfirm={() => deleteUser(record.id)}
            okText={t("common", "delete")}
            cancelText={t("common", "cancel")}
            okButtonProps={{ danger: true }}
          >
            <Tooltip title={t("common", "delete")}>
              <Button
                danger
                icon={<DeleteOutlined />}
                size="small"
                style={{ borderRadius: 8 }}
              />
            </Tooltip>
          </Popconfirm>
        </Space>
      ),
    },
  ];

  return (
    <Layout style={{ minHeight: "100vh", background: "#ffffff" }}>
      <Navbar />
      <Content style={{ padding: 0 }}>
        {/* Hero header */}
        <div
          style={{
            background:
              "linear-gradient(135deg, #c2185b 0%, #e91e90 40%, #ff9800 100%)",
            padding: "36px 48px 32px",
            position: "relative",
            overflow: "hidden",
          }}
        >
          <div
            style={{
              position: "absolute",
              top: -60,
              right: -40,
              width: 250,
              height: 250,
              borderRadius: "50%",
              background: "rgba(255,255,255,0.05)",
            }}
          />
          <div
            style={{
              position: "absolute",
              bottom: -30,
              left: "30%",
              width: 180,
              height: 180,
              borderRadius: "50%",
              background: "rgba(255,255,255,0.03)",
            }}
          />
          <div style={{ maxWidth: 1200, margin: "0 auto" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
              <div
                style={{
                  width: 46,
                  height: 46,
                  borderRadius: 14,
                  background: "rgba(255,255,255,0.12)",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                }}
              >
                <SettingOutlined style={{ fontSize: 24, color: "#fff" }} />
              </div>
              <div>
                <Title
                  level={3}
                  style={{ color: "#fff", margin: 0, fontWeight: 700 }}
                >
                  {t("admin", "title")}
                </Title>
                <Text style={{ color: "rgba(255,255,255,0.6)", fontSize: 13 }}>
                  {t("admin", "subtitle")}
                </Text>
              </div>
            </div>
          </div>
        </div>

        <div style={{ maxWidth: 1200, margin: "0 auto", padding: "24px 48px 48px" }}>
          {/* Stat Cards - overlapping hero */}
          <Row gutter={[16, 16]} style={{ marginBottom: 24, marginTop: -48 }}>
            {statCards.map((card) => (
              <Col xs={12} md={6} key={card.label}>
                <Card
                  style={{
                    borderRadius: 16,
                    border: "none",
                    boxShadow: "0 4px 20px rgba(108,44,145,0.1)",
                    overflow: "hidden",
                  }}
                  styles={{ body: { padding: "20px 20px 16px" } }}
                >
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "flex-start",
                    }}
                  >
                    <div>
                      <Text
                        style={{
                          color: "#888",
                          fontSize: 12,
                          textTransform: "uppercase",
                          letterSpacing: 0.5,
                        }}
                      >
                        {card.label}
                      </Text>
                      <Title
                        level={2}
                        style={{ margin: "4px 0 0", color: card.color }}
                      >
                        {card.value}
                      </Title>
                    </div>
                    <div
                      style={{
                        width: 42,
                        height: 42,
                        borderRadius: 12,
                        background: card.gradient,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        color: "#fff",
                      }}
                    >
                      {card.icon}
                    </div>
                  </div>
                </Card>
              </Col>
            ))}
          </Row>

          {/* Search bar */}
          <div style={{ display: "flex", gap: 12, marginBottom: 20, flexWrap: "wrap" }}>
            <Input
              placeholder={t("admin", "searchPlaceholder")}
              prefix={<SearchOutlined style={{ color: "#ba68c8" }} />}
              allowClear
              style={{ width: 320, borderRadius: 10, height: 40 }}
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
            <Button
              onClick={fetchUsers}
              icon={<SyncOutlined />}
              style={{ borderRadius: 10, height: 40 }}
            >
              {t("common", "refresh")}
            </Button>
          </div>

          {/* Table */}
          <Card
            style={{
              borderRadius: 16,
              border: "none",
              boxShadow: "0 2px 12px rgba(108,44,145,0.06)",
            }}
          >
            <Table
              dataSource={filteredUsers}
              columns={columns}
              rowKey="id"
              loading={loading}
              pagination={{
                pageSize: 12,
                showSizeChanger: true,
                showTotal: (total) => `${total} ${t("common", "users")}${total > 1 ? "s" : ""}`,
              }}
              size="middle"
              style={{
                // @ts-expect-error Ant Design Table custom CSS
                "--ant-table-header-bg": "#f9f0ff",
              }}
            />
          </Card>

          {/* Edit modal */}
          <Modal
            title={
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div
                  style={{
                    width: 32,
                    height: 32,
                    borderRadius: 8,
                    background: "linear-gradient(135deg, #e91e90, #ff9800)",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                  }}
                >
                  <EditOutlined style={{ color: "#fff", fontSize: 15 }} />
                </div>
                <span style={{ fontWeight: 600, fontSize: 16 }}>
                  {t("admin", "editUser")}
                </span>
              </div>
            }
            open={editModalOpen}
            onCancel={() => setEditModalOpen(false)}
            onOk={handleEditSave}
            confirmLoading={saving}
            okText="Enregistrer"
            cancelText="Annuler"
            width={500}
            okButtonProps={{
              style: {
                background: "linear-gradient(135deg, #e91e90, #ff9800)",
                border: "none",
                borderRadius: 8,
                fontWeight: 600,
              },
            }}
            cancelButtonProps={{
              style: { borderRadius: 8 },
            }}
            styles={{
              header: { borderBottom: "1px solid #f0e6f6" },
              body: { paddingTop: 20 },
            }}
          >
            <Form form={editForm} layout="vertical">
              <Form.Item
                name="email"
                label={<span style={{ fontWeight: 500 }}>{t("admin", "email")}</span>}
                rules={[
                  { required: true, message: "Email requis" },
                  { type: "email", message: "Email invalide" },
                ]}
              >
                <Input
                  prefix={<UserOutlined style={{ color: "#ba68c8" }} />}
                  style={{ borderRadius: 8 }}
                />
              </Form.Item>
              <Form.Item
                name="full_name"
                label={<span style={{ fontWeight: 500 }}>{t("admin", "fullName")}</span>}
              >
                <Input
                  prefix={<UserOutlined style={{ color: "#ba68c8" }} />}
                  style={{ borderRadius: 8 }}
                />
              </Form.Item>
              <Row gutter={16}>
                <Col span={12}>
                  <Form.Item
                    name="role"
                    label={<span style={{ fontWeight: 500 }}>{t("admin", "role")}</span>}
                    rules={[{ required: true }]}
                  >
                    <Select
                      style={{ borderRadius: 8 }}
                      options={[
                        { value: "user", label: "User" },
                        { value: "admin", label: "Admin" },
                      ]}
                    />
                  </Form.Item>
                </Col>
                <Col span={12}>
                  <Form.Item
                    name="is_active"
                    label={<span style={{ fontWeight: 500 }}>{t("admin", "statusCol")}</span>}
                    valuePropName="checked"
                  >
                    <Switch checkedChildren={t("admin", "activeSwitch")} unCheckedChildren={t("admin", "inactiveSwitch")} />
                  </Form.Item>
                </Col>
              </Row>
              <Form.Item
                name="new_password"
                label={<span style={{ fontWeight: 500 }}>{t("admin", "newPassword")}</span>}
                help="Laisser vide pour ne pas changer"
              >
                <Input.Password
                  prefix={<LockOutlined style={{ color: "#ba68c8" }} />}
                  placeholder="Nouveau mot de passe"
                  style={{ borderRadius: 8 }}
                />
              </Form.Item>
            </Form>
          </Modal>
        </div>
      </Content>
    </Layout>
  );
}
