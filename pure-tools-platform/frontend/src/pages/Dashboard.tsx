import { useEffect, useState } from "react";
import { Button, Card, Col, Layout, Row, Spin, Typography, message } from "antd";
import Navbar from "../components/Navbar";

import {
  CheckCircleOutlined,
  FilePdfOutlined,
  PlayCircleOutlined,
  TeamOutlined,
  ToolOutlined,
  UserOutlined,
  DashboardOutlined,
  ThunderboltOutlined,
  TrophyOutlined,
  FireOutlined,
  ExperimentOutlined,
} from "@ant-design/icons";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  Legend,
  AreaChart,
  Area,
} from "recharts";
import api from "../api/axios";
import { useLanguage } from "../context/LanguageContext";

const { Content } = Layout;
const { Title, Text } = Typography;

const COLORS = ["#d77fa3", "#ff9800", "#ffca28", "#a9516a", "#9c27b0", "#f57c00", "#ffb74d"];

interface Stats {
  total_executions: number;
  completed: number;
  failed: number;
  success_rate: number;
  total_users: number;
  most_used_tool: string | null;
  my_executions: number;
  per_tool: { tool: string; count: number }[];
  per_day: { date: string; count: number }[];
  per_status: { status: string; count: number }[];
}

export default function Dashboard() {
  const [stats, setStats] = useState<Stats | null>(null);
  const [loading, setLoading] = useState(true);
  const { t } = useLanguage();

  useEffect(() => {
    api
      .get("/api/dashboard/stats")
      .then((res) => setStats(res.data))
      .finally(() => setLoading(false));
  }, []);

  if (loading) {
    return (
      <Layout style={{ minHeight: "100vh", background: "#ffffff" }}>
        <Navbar />
        <div style={{ textAlign: "center", marginTop: 160 }}>
          <Spin size="large" />
        </div>
      </Layout>
    );
  }

  if (!stats) return null;

  const handleExport = async () => {
    try {
      const res = await api.get("/api/dashboard/export", { responseType: "blob" });
      const url = globalThis.URL.createObjectURL(new Blob([res.data]));
      const link = document.createElement("a");
      link.href = url;
      const disposition = res.headers["content-disposition"];
      const filename = disposition
        ? disposition.split("filename=")[1]?.replaceAll('"', "")
        : "dashboard.pdf";
      link.setAttribute("download", filename);
      document.body.appendChild(link);
      link.click();
      link.remove();
      globalThis.URL.revokeObjectURL(url);
    } catch {
      message.error(t("dashboard", "exportError"));
    }
  };

  const statCards = [
    {
      label: t("dashboard", "totalExecutions"),
      value: stats.total_executions,
      icon: <ExperimentOutlined style={{ fontSize: 22 }} />,
      gradient: "linear-gradient(135deg, #d77fa3, #a9516a)",
      color: "#9369b8",
    },
    {
      label: t("dashboard", "successful"),
      value: stats.completed,
      icon: <TrophyOutlined style={{ fontSize: 22 }} />,
      gradient: "linear-gradient(135deg, #ff9800, #ffca28)",
      color: "#ff9800",
    },
    {
      label: t("dashboard", "failed"),
      value: stats.failed,
      icon: <FireOutlined style={{ fontSize: 22 }} />,
      gradient: "linear-gradient(135deg, #a9516a, #9c27b0)",
      color: "#a9516a",
    },
    {
      label: t("dashboard", "successRate"),
      value: `${stats.success_rate}%`,
      icon: <ThunderboltOutlined style={{ fontSize: 22 }} />,
      gradient: "linear-gradient(135deg, #ffca28, #ff9800)",
      color: "#ffca28",
    },
  ];

  const kpiCards = [
    {
      label: t("dashboard", "usersLabel"),
      value: stats.total_users,
      icon: <TeamOutlined style={{ fontSize: 20 }} />,
      gradient: "linear-gradient(135deg, #d77fa3, #ff9800)",
      color: "#9369b8",
    },
    {
      label: t("dashboard", "myExecutions"),
      value: stats.my_executions,
      icon: <UserOutlined style={{ fontSize: 20 }} />,
      gradient: "linear-gradient(135deg, #ff9800, #a9516a)",
      color: "#ff9800",
    },
    {
      label: t("dashboard", "mostUsedTool"),
      value: stats.most_used_tool ?? "—",
      icon: <ToolOutlined style={{ fontSize: 20 }} />,
      gradient: "linear-gradient(135deg, #ffca28, #ff9800)",
      color: "#ffca28",
      isText: true,
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
              "linear-gradient(135deg, #a9516a 0%, #d77fa3 40%, #ff9800 100%)",
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
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
              }}
            >
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
                  <DashboardOutlined style={{ fontSize: 24, color: "#fff" }} />
                </div>
                <div>
                  <Title
                    level={3}
                    style={{ color: "#fff", margin: 0, fontWeight: 700 }}
                  >
                    {t("dashboard", "title")}
                  </Title>
                  <Text style={{ color: "rgba(255,255,255,0.6)", fontSize: 13 }}>
                    {t("dashboard", "subtitle")}
                  </Text>
                </div>
              </div>
              <Button
                icon={<FilePdfOutlined />}
                onClick={handleExport}
                size="large"
                style={{
                  borderRadius: 12,
                  background: "rgba(255,255,255,0.15)",
                  color: "#fff",
                  border: "1px solid rgba(255,255,255,0.25)",
                  fontWeight: 600,
                  backdropFilter: "blur(4px)",
                }}
              >
                {t("dashboard", "exportPdf")}
              </Button>
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

          {/* Secondary KPI row */}
          <Row gutter={[16, 16]} style={{ marginBottom: 24 }}>
            {kpiCards.map((card) => (
              <Col xs={24} md={8} key={card.label}>
                <Card
                  style={{
                    borderRadius: 16,
                    border: "none",
                    boxShadow: "0 2px 12px rgba(108,44,145,0.06)",
                  }}
                  styles={{ body: { padding: "18px 20px" } }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                    <div
                      style={{
                        width: 40,
                        height: 40,
                        borderRadius: 10,
                        background: card.gradient,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        color: "#fff",
                        flexShrink: 0,
                      }}
                    >
                      {card.icon}
                    </div>
                    <div>
                      <Text style={{ color: "#888", fontSize: 12, textTransform: "uppercase", letterSpacing: 0.5 }}>
                        {card.label}
                      </Text>
                      <div style={{ fontSize: card.isText ? 16 : 22, fontWeight: 700, color: card.color, lineHeight: 1.2 }}>
                        {card.value}
                      </div>
                    </div>
                  </div>
                </Card>
              </Col>
            ))}
          </Row>

          {/* Charts */}
          <Row gutter={[16, 16]}>
            <Col xs={24} lg={14}>
              <Card
                title={
                  <span style={{ fontWeight: 600, color: "#333" }}>
                    <PlayCircleOutlined style={{ marginRight: 8, color: "#ba68c8" }} />
                    {t("dashboard", "executionsPerDay")}
                  </span>
                }
                style={{
                  borderRadius: 16,
                  border: "none",
                  boxShadow: "0 2px 12px rgba(108,44,145,0.06)",
                }}
              >
                {stats.per_day.length > 0 ? (
                  <ResponsiveContainer width="100%" height={300}>
                    <AreaChart data={stats.per_day}>
                      <defs>
                        <linearGradient id="colorCount" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor="#e91e90" stopOpacity={0.3} />
                          <stop offset="95%" stopColor="#e91e90" stopOpacity={0.02} />
                        </linearGradient>
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0e6f6" />
                      <XAxis
                        dataKey="date"
                        tickFormatter={(v: string) => {
                          const d = new Date(v);
                          return `${d.getDate()}/${d.getMonth() + 1}`;
                        }}
                        tick={{ fill: "#888", fontSize: 12 }}
                      />
                      <YAxis allowDecimals={false} tick={{ fill: "#888", fontSize: 12 }} />
                      <Tooltip
                        contentStyle={{
                          borderRadius: 10,
                          border: "none",
                          boxShadow: "0 4px 16px rgba(0,0,0,0.1)",
                        }}
                      />
                      <Area
                        type="monotone"
                        dataKey="count"
                        stroke="#e91e90"
                        strokeWidth={2.5}
                        fill="url(#colorCount)"
                        name={t("dashboard", "executionsLabel")}
                      />
                    </AreaChart>
                  </ResponsiveContainer>
                ) : (
                  <div style={{ textAlign: "center", padding: 40, color: "#999" }}>
                    {t("common", "noData")}
                  </div>
                )}
              </Card>
            </Col>
            <Col xs={24} lg={10}>
              <Card
                title={
                  <span style={{ fontWeight: 600, color: "#333" }}>
                    <ToolOutlined style={{ marginRight: 8, color: "#ff9800" }} />
                    {t("dashboard", "toolDistribution")}
                  </span>
                }
                style={{
                  borderRadius: 16,
                  border: "none",
                  boxShadow: "0 2px 12px rgba(108,44,145,0.06)",
                }}
              >
                {stats.per_tool.length > 0 ? (
                  <ResponsiveContainer width="100%" height={300}>
                    <PieChart>
                      <Pie
                        data={stats.per_tool}
                        dataKey="count"
                        nameKey="tool"
                        cx="50%"
                        cy="50%"
                        outerRadius={100}
                        innerRadius={55}
                        paddingAngle={3}
                        label={({ tool, count }: { tool: string; count: number }) =>
                          `${tool} (${count})`
                        }
                      >
                        {stats.per_tool.map((_, i) => (
                          <Cell key={i} fill={COLORS[i % COLORS.length]} />
                        ))}
                      </Pie>
                      <Tooltip
                        contentStyle={{
                          borderRadius: 10,
                          border: "none",
                          boxShadow: "0 4px 16px rgba(0,0,0,0.1)",
                        }}
                      />
                      <Legend />
                    </PieChart>
                  </ResponsiveContainer>
                ) : (
                  <div style={{ textAlign: "center", padding: 40, color: "#999" }}>
                    {t("common", "noData")}
                  </div>
                )}
              </Card>
            </Col>
          </Row>

          {/* Status breakdown */}
          <Row gutter={[16, 16]} style={{ marginTop: 16 }}>
            <Col xs={24}>
              <Card
                title={
                  <span style={{ fontWeight: 600, color: "#333" }}>
                    <CheckCircleOutlined style={{ marginRight: 8, color: "#ba68c8" }} />
                    {t("dashboard", "statusDistribution")}
                  </span>
                }
                style={{
                  borderRadius: 16,
                  border: "none",
                  boxShadow: "0 2px 12px rgba(108,44,145,0.06)",
                }}
              >
                <ResponsiveContainer width="100%" height={200}>
                  <BarChart data={stats.per_status} layout="vertical">
                    <CartesianGrid strokeDasharray="3 3" stroke="#f0e6f6" />
                    <XAxis type="number" allowDecimals={false} tick={{ fill: "#888", fontSize: 12 }} />
                    <YAxis
                      type="category"
                      dataKey="status"
                      tick={{ fill: "#888", fontSize: 12 }}
                      width={80}
                      tickFormatter={(v: string) => {
                        const labels: Record<string, string> = { completed: t("history", "completed"), failed: t("history", "failedStatus"), other: t("common", "noData") };
                        return labels[v] || v;
                      }}
                    />
                    <Tooltip
                      contentStyle={{
                        borderRadius: 10,
                        border: "none",
                        boxShadow: "0 4px 16px rgba(0,0,0,0.1)",
                      }}
                    />
                    <Bar dataKey="count" name={t("dashboard", "executionsLabel")} radius={[0, 8, 8, 0]}>
                      {stats.per_status.map((entry, i) => {
                        const statusColors: Record<string, string> = {
                          completed: "#11998e",
                          failed: "#e74c3c",
                          other: "#f39c12",
                        };
                        return <Cell key={i} fill={statusColors[entry.status] || COLORS[i % COLORS.length]} />;
                      })}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
            </Col>
          </Row>
        </div>
      </Content>
    </Layout>
  );
}
