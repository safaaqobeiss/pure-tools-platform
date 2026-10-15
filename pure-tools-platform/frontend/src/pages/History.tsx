import { useEffect, useState, useMemo } from "react";
import { useNavigate } from "react-router-dom";
import {
  Layout,
  Typography,
  Table,
  Tag,
  Button,
  message,
  Space,
  Tooltip,
  Empty,
  Input,
  Select,
  Avatar,
} from "antd";
import {
  DownloadOutlined,
  CheckCircleOutlined,
  CloseCircleOutlined,
  SyncOutlined,
  ClockCircleOutlined,
  ArrowLeftOutlined,
  HistoryOutlined,
  SearchOutlined,
  UserOutlined,
} from "@ant-design/icons";
import type { ColumnsType } from "antd/es/table";
import api from "../api/axios";
import Navbar from "../components/Navbar";
import { useAuth } from "../context/AuthContext";
import { useLanguage } from "../context/LanguageContext";

const { Content } = Layout;
const { Title, Text } = Typography;

interface Execution {
  id: string;
  tool_id: string;
  tool_name: string | null;
  script_path: string | null;
  status: string;
  created_at: string;
  error_message: string | null;
  user_name: string | null;
  user_email: string | null;
}

const STATUS_MAP: Record<
  string,
  { color: string; icon: React.ReactNode; label: string }
> = {
  completed: {
    color: "success",
    icon: <CheckCircleOutlined />,
    label: "Terminé",
  },
  failed: {
    color: "error",
    icon: <CloseCircleOutlined />,
    label: "Échoué",
  },
  running: {
    color: "processing",
    icon: <SyncOutlined spin />,
    label: "En cours",
  },
  pending: {
    color: "default",
    icon: <ClockCircleOutlined />,
    label: "En attente",
  },
};

export default function History() {
  const navigate = useNavigate();
  const { user } = useAuth();
  const { t } = useLanguage();
  const isAdmin = user?.role === "admin";
  const [executions, setExecutions] = useState<Execution[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState<string | null>(null);

  const fetchExecutions = () => {
    setLoading(true);
    api
      .get("/api/executions/")
      .then((res) => setExecutions(res.data))
      .catch(() => message.error(t("history", "loadError")))
      .finally(() => setLoading(false));
  };

  useEffect(() => {
    fetchExecutions();
  }, []);

  // Filtered list
  const filtered = useMemo(() => {
    let result = executions;
    if (search) {
      const q = search.toLowerCase();
      result = result.filter(
        (e) =>
          e.tool_name?.toLowerCase().includes(q) ||
          e.error_message?.toLowerCase().includes(q) ||
          e.user_name?.toLowerCase().includes(q) ||
          e.user_email?.toLowerCase().includes(q)
      );
    }
    if (statusFilter) {
      result = result.filter((e) => e.status === statusFilter);
    }
    return result;
  }, [executions, search, statusFilter]);

  const handleDownload = async (execution: Execution) => {
    try {
      const downloadUrl =
        execution.script_path === "forms_writer"
          ? `/api/forms-generator/${execution.id}/download`
          : `/api/executions/${execution.id}/download`;

      const res = await api.get(downloadUrl, { responseType: "blob" });

      const contentDisposition = res.headers["content-disposition"];
      let filename = "result.zip";
      if (contentDisposition) {
        const match = contentDisposition.match(/filename="?([^"]+)"?/);
        if (match) filename = match[1];
      }

      const url = globalThis.URL.createObjectURL(new Blob([res.data]));
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", filename);
      document.body.appendChild(link);
      link.click();
      link.remove();
      globalThis.URL.revokeObjectURL(url);
    } catch {
      message.error(t("history", "downloadError"));
    }
  };

  const STATUS_MAP_T: Record<
    string,
    { color: string; icon: React.ReactNode; label: string }
  > = {
    completed: {
      color: "success",
      icon: <CheckCircleOutlined />,
      label: t("history", "completed"),
    },
    failed: {
      color: "error",
      icon: <CloseCircleOutlined />,
      label: t("history", "failedStatus"),
    },
    running: {
      color: "processing",
      icon: <SyncOutlined spin />,
      label: t("history", "running"),
    },
    pending: {
      color: "default",
      icon: <ClockCircleOutlined />,
      label: t("history", "pending"),
    },
  };

  const columns: ColumnsType<Execution> = [
    {
      title: t("history", "tool"),
      dataIndex: "tool_name",
      key: "tool_name",
      render: (name: string | null) => (
        <Text strong style={{ color: "#ba68c8" }}>
          {name || "—"}
        </Text>
      ),
    },
    ...(isAdmin
      ? [
          {
            title: t("history", "userCol"),
            key: "user",
            width: 200,
            render: (_: unknown, record: Execution) => (
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <Avatar
                  size={28}
                  icon={<UserOutlined />}
                  style={{
                    background: "linear-gradient(135deg, #e91e90, #c2185b)",
                    fontSize: 12,
                  }}
                >
                  {(record.user_name || record.user_email || "?")[0].toUpperCase()}
                </Avatar>
                <div style={{ lineHeight: 1.3 }}>
                  <Text style={{ fontSize: 13, fontWeight: 500 }}>
                    {record.user_name || "—"}
                  </Text>
                  <br />
                  <Text style={{ fontSize: 11, color: "#888" }}>
                    {record.user_email}
                  </Text>
                </div>
              </div>
            ),
          } as const,
        ]
      : []),
    {
      title: t("history", "status"),
      dataIndex: "status",
      key: "status",
      width: 130,
      render: (status: string) => {
        const info = STATUS_MAP_T[status] || STATUS_MAP_T.pending;
        return (
          <Tag icon={info.icon} color={info.color}>
            {info.label}
          </Tag>
        );
      },
    },
    {
      title: t("history", "date"),
      dataIndex: "created_at",
      key: "created_at",
      width: 180,
      render: (date: string) =>
        new Date(date).toLocaleString("fr-FR", {
          day: "2-digit",
          month: "2-digit",
          year: "numeric",
          hour: "2-digit",
          minute: "2-digit",
        }),
      sorter: (a, b) =>
        new Date(a.created_at).getTime() - new Date(b.created_at).getTime(),
      defaultSortOrder: "descend",
    },
    {
      title: t("history", "errorCol"),
      dataIndex: "error_message",
      key: "error_message",
      ellipsis: true,
      render: (msg: string | null) =>
        msg ? (
          <Tooltip title={msg}>
            <Text type="danger" ellipsis style={{ maxWidth: 300 }}>
              {msg}
            </Text>
          </Tooltip>
        ) : (
          <Text type="secondary">—</Text>
        ),
    },
    {
      title: t("history", "actions"),
      key: "actions",
      width: 140,
      render: (_: unknown, record: Execution) => (
        <Space>
          {record.status === "completed" && (
            <Button
              size="small"
              icon={<DownloadOutlined />}
              onClick={() => handleDownload(record)}
              style={{
                borderRadius: 8,
                background: "linear-gradient(135deg, #e91e90, #ff9800)",
                color: "#fff",
                border: "none",
                fontWeight: 500,
              }}
            >
              {t("common", "download")}
            </Button>
          )}
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
          <div style={{ maxWidth: 1200, margin: "0 auto" }}>
            <Button
              icon={<ArrowLeftOutlined />}
              type="text"
              onClick={() => navigate("/home")}
              style={{
                color: "rgba(255,255,255,0.8)",
                fontWeight: 500,
                marginBottom: 12,
                padding: 0,
              }}
            >
              {t("common", "backToTools")}
            </Button>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 14,
              }}
            >
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
                <HistoryOutlined style={{ fontSize: 24, color: "#fff" }} />
              </div>
              <div>
                <Title
                  level={3}
                  style={{ color: "#fff", margin: 0, fontWeight: 700 }}
                >
                  {t("history", "title")}
                </Title>
                <Text style={{ color: "rgba(255,255,255,0.6)", fontSize: 13 }}>
                  {executions.length} {t("history", "totalLabel")}{executions.length > 1 ? "s" : ""} {t("history", "totalSuffix")}
                </Text>
              </div>
            </div>
          </div>
        </div>

        <div style={{ maxWidth: 1200, margin: "0 auto", padding: "24px 48px 48px" }}>
          {/* Search & filter bar */}
          <div
            style={{
              display: "flex",
              gap: 12,
              marginBottom: 20,
              flexWrap: "wrap",
            }}
          >
            <Input
              placeholder={t("history", "searchPlaceholder")}
              prefix={<SearchOutlined style={{ color: "#ba68c8" }} />}
              allowClear
              style={{ width: 320, borderRadius: 10, height: 40 }}
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
            <Select
              placeholder={t("history", "allStatuses")}
              allowClear
              style={{ width: 180 }}
              value={statusFilter}
              onChange={(val) => setStatusFilter(val ?? null)}
              options={[
                { value: "completed", label: t("history", "completed") },
                { value: "failed", label: t("history", "failedStatus") },
                { value: "running", label: t("history", "running") },
                { value: "pending", label: t("history", "pending") },
              ]}
            />
            <Button
              onClick={fetchExecutions}
              icon={<SyncOutlined />}
              style={{ borderRadius: 10, height: 40 }}
            >
              {t("common", "refresh")}
            </Button>
          </div>

          {/* Table */}
          <Table
            columns={columns}
            dataSource={filtered}
            rowKey="id"
            loading={loading}
            pagination={{ pageSize: 12, showSizeChanger: true, showTotal: (total) => `${total} ${t("common", "results")}${total > 1 ? "s" : ""}` }}
            locale={{
              emptyText: (
                <Empty description={t("history", "noExecutions")} />
              ),
            }}
            style={{
              background: "white",
              borderRadius: 16,
              overflow: "hidden",
              boxShadow: "0 2px 12px rgba(108,44,145,0.06)",
            }}
          />
        </div>
      </Content>
    </Layout>
  );
}
