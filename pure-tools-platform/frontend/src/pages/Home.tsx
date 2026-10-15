import { useEffect, useState, useMemo } from "react";
import { useNavigate } from "react-router-dom";
import { Card, Row, Col, Typography, Spin, Empty, Tag, Layout, Input, Select, Space } from "antd";
import {
  ToolOutlined,
  SwapOutlined,
  FileExcelOutlined,
  FormOutlined,
  ProjectOutlined,
  SearchOutlined,
  ArrowRightOutlined,
  ThunderboltOutlined,
  SafetyCertificateOutlined,
  HistoryOutlined,
  DashboardOutlined,
  TeamOutlined,
} from "@ant-design/icons";
import api from "../api/axios";
import Navbar from "../components/Navbar";
import { useAuth } from "../context/AuthContext";
import { useLanguage } from "../context/LanguageContext";
import translations from "../i18n/translations";

const { Content } = Layout;

const ICON_MAP: Record<string, { icon: React.ReactNode; gradient: string }> = {
  SwapOutlined: {
    icon: <SwapOutlined style={{ fontSize: 28, color: "#fff" }} />,
    gradient: "linear-gradient(135deg, #d77fa3, #ff9800)",
  },
  FileExcelOutlined: {
    icon: <FileExcelOutlined style={{ fontSize: 28, color: "#fff" }} />,
    gradient: "linear-gradient(135deg, #a9516a, #9c27b0)",
  },
  FormOutlined: {
    icon: <FormOutlined style={{ fontSize: 28, color: "#fff" }} />,
    gradient: "linear-gradient(135deg, #ffca28, #ff9800)",
  },
  ProjectOutlined: {
    icon: <ProjectOutlined style={{ fontSize: 28, color: "#fff" }} />,
    gradient: "linear-gradient(135deg, #a9516a, #d77fa3)",
  },
  SafetyCertificateOutlined: {
    icon: <SafetyCertificateOutlined style={{ fontSize: 28, color: "#fff" }} />,
    gradient: "linear-gradient(135deg, #ff9800, #ffca28)",
  },
};

const DEFAULT_TOOL_STYLE = {
  icon: <ToolOutlined style={{ fontSize: 28, color: "#fff" }} />,
  gradient: "linear-gradient(135deg, #d77fa3, #a9516a)",
};

interface Tool {
  id: string;
  name: string;
  description: string | null;
  icon: string | null;
  category: string | null;
  script_path: string;
  is_active: boolean;
}

export default function Home() {
  const [tools, setTools] = useState<Tool[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState<string | null>(null);
  const navigate = useNavigate();
  const { user } = useAuth();
  const { t, lang } = useLanguage();

  // Helper: get translated tool field, fallback to DB value
  const toolText = (tool: Tool, field: "name" | "description" | "category"): string => {
    const entry = (translations.tools as Record<string, Record<string, { fr: string; en: string }>>)[tool.script_path];
    if (entry?.[field]) return entry[field][lang];
    if (field === "name") return tool.name;
    if (field === "description") return tool.description || "";
    return tool.category || "";
  };

  useEffect(() => {
    api
      .get("/api/tools/")
      .then((res) => setTools(res.data))
      .catch(() => setTools([]))
      .finally(() => setLoading(false));
  }, []);

  const categories = useMemo(
    () => [...new Set(tools.map((tool) => toolText(tool, "category")).filter(Boolean))] as string[],
    [tools, lang],
  );

  const filtered = useMemo(() => {
    let result = tools;
    if (search) {
      const q = search.toLowerCase();
      result = result.filter(
        (tool) =>
          toolText(tool, "name").toLowerCase().includes(q) ||
          toolText(tool, "description")?.toLowerCase().includes(q),
      );
    }
    if (categoryFilter) {
      result = result.filter((tool) => toolText(tool, "category") === categoryFilter);
    }
    // Sort to place ESPF to RDS converter first
    result = result.sort((a, b) => {
      if (a.script_path === "convert_espf_to_rds") return -1;
      if (b.script_path === "convert_espf_to_rds") return 1;
      return 0;
    });
    return result;
  }, [tools, search, categoryFilter, lang]);

  return (
    <Layout style={{ minHeight: "100vh", background: "#ffffff" }}>
      <Navbar />
      <Content style={{ padding: "0" }}>
        {/* Hero banner */}
        <div
          style={{
            background: "linear-gradient(135deg, #a9516a 0%, #d77fa3 40%, #ff9800 100%)",
            padding: "32px 48px 28px",
            color: "#fff",
          }}
        >
          <div style={{ maxWidth: 1200, margin: "0 auto" }}>
            <Space size={8} align="center" style={{ marginBottom: 4 }}>
              <ThunderboltOutlined style={{ fontSize: 20, color: "#ff9800" }} />
              <Typography.Text style={{ color: "rgba(255,255,255,0.6)", fontSize: 13, letterSpacing: 1, textTransform: "uppercase" }}>
                {t("home", "platformLabel")}
              </Typography.Text>
            </Space>
            <Typography.Title level={3} style={{ color: "#fff", margin: "8px 0 4px", fontWeight: 700 }}>
              {t("home", "greeting")}, {user?.full_name?.split(" ")[0] || t("home", "user")} !
            </Typography.Title>
            <Typography.Text style={{ color: "rgba(255,255,255,0.7)", fontSize: 14 }}>
              {tools.length} {t("home", "toolsAvailable")}{tools.length > 1 ? "s" : ""} — {t("home", "selectToStart")}
            </Typography.Text>

            {/* Search bar inside hero */}
            <div style={{ marginTop: 24, display: "flex", gap: 12, flexWrap: "wrap" }}>
              <Input
                placeholder={t("home", "searchPlaceholder")}
                prefix={<SearchOutlined style={{ color: "rgba(255,255,255,0.4)" }} />}
                allowClear
                style={{
                  width: 340,
                  background: "rgba(255,255,255,0.12)",
                  border: "1px solid rgba(255,255,255,0.2)",
                  borderRadius: 10,
                  color: "#fff",
                }}
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
              <Select
                placeholder={t("home", "allCategories")}
                allowClear
                style={{ width: 200 }}
                value={categoryFilter}
                onChange={(val) => setCategoryFilter(val ?? null)}
                options={categories.map((c) => ({ value: c, label: c }))}
              />
            </div>
          </div>
        </div>

        {/* A propos section */}
          <div style={{ background: "#ffffff", padding: "40px 48px" }}>
          <div style={{ maxWidth: 1200, margin: "0 auto" }}>
            <Typography.Title level={3} style={{ color: "#9369b8", marginBottom: 8 }}>
              {t("home", "aboutTitle")}
            </Typography.Title>
            <Typography.Paragraph style={{ color: "#555", fontSize: 15, maxWidth: 700, marginBottom: 28 }}>
              {t("home", "aboutDesc")}
            </Typography.Paragraph>
            <Row gutter={[24, 16]}>
              {[
                { icon: <ToolOutlined style={{ fontSize: 32, color: "#ba68c8" }} />, title: t("home", "featureIntegrated"), desc: t("home", "featureIntegratedDesc") },
                { icon: <HistoryOutlined style={{ fontSize: 32, color: "#ba68c8" }} />, title: t("home", "featureHistory"), desc: t("home", "featureHistoryDesc") },
                { icon: <DashboardOutlined style={{ fontSize: 32, color: "#ba68c8" }} />, title: t("home", "featureDashboard"), desc: t("home", "featureDashboardDesc") },
                { icon: <SafetyCertificateOutlined style={{ fontSize: 32, color: "#ba68c8" }} />, title: t("home", "featureSecure"), desc: t("home", "featureSecureDesc") },
                { icon: <TeamOutlined style={{ fontSize: 32, color: "#ba68c8" }} />, title: t("home", "featureMultiUser"), desc: t("home", "featureMultiUserDesc") },
                { icon: <ThunderboltOutlined style={{ fontSize: 32, color: "#ba68c8" }} />, title: t("home", "featureRealtime"), desc: t("home", "featureRealtimeDesc") },
              ].map((item, i) => (
                <Col xs={24} sm={12} md={8} key={i}>
                  <div style={{ display: "flex", alignItems: "flex-start", gap: 14, padding: "12px 0" }}>
                    <div style={{ flexShrink: 0, marginTop: 2 }}>
                      {item.icon}
                    </div>
                    <div>
                      <Typography.Text strong style={{ fontSize: 14, color: "#9369b8" }}>{item.title}</Typography.Text>
                      <Typography.Paragraph style={{ margin: "4px 0 0", fontSize: 13, color: "#666" }}>{item.desc}</Typography.Paragraph>
                    </div>
                  </div>
                </Col>
              ))}
            </Row>
          </div>
        </div>

        {/* Tools grid */}
        <div style={{ maxWidth: 1200, margin: "0 auto", padding: "32px 48px 48px" }}>
          <Typography.Title level={3} style={{ color: "#9369b8", marginBottom: 20 }}>
            {t("home", "ourTools")}
          </Typography.Title>
          {loading ? (
            <div style={{ textAlign: "center", marginTop: 80 }}>
              <Spin size="large" />
            </div>
          ) : tools.length === 0 ? (
            <Empty description={t("home", "noToolsAvailable")} />
          ) : filtered.length === 0 ? (
            <Empty description={t("home", "noToolsMatch")} />
          ) : (
            <Row gutter={[20, 20]}>
              {filtered.map((tool) => {
                const style = ICON_MAP[tool.icon || ""] || DEFAULT_TOOL_STYLE;
                return (
                  <Col xs={24} sm={12} md={8} key={tool.id}>
                    <Card
                      hoverable
                      style={{
                        borderRadius: 14,
                        height: "100%",
                        border: "none",
                        boxShadow: "0 2px 12px rgba(0,0,0,0.06)",
                        transition: "transform 0.2s, box-shadow 0.2s",
                        overflow: "hidden",
                      }}
                      styles={{ body: { padding: 0 } }}
                      onClick={() =>
                        tool.script_path === "init_project"
                          ? navigate("/tool/init-project")
                          : tool.script_path === "forms_writer"
                          ? navigate("/tool/forms-generator")
                          : tool.script_path === "check_submission"
                          ? navigate("/tool/check-submission")
                          : navigate(`/tool/${tool.id}`)
                      }
                      onMouseEnter={(e) => {
                        e.currentTarget.style.transform = "translateY(-4px)";
                        e.currentTarget.style.boxShadow = "0 8px 30px rgba(0,0,0,0.12)";
                      }}
                      onMouseLeave={(e) => {
                        e.currentTarget.style.transform = "translateY(0)";
                        e.currentTarget.style.boxShadow = "0 2px 12px rgba(0,0,0,0.06)";
                      }}
                    >
                      {/* Gradient header */}
                      <div style={{ background: style.gradient, padding: "20px 20px 16px", display: "flex", alignItems: "center", gap: 14 }}>
                        <div style={{ width: 48, height: 48, borderRadius: 12, background: "rgba(255,255,255,0.2)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                          {style.icon}
                        </div>
                        <div>
                          <Typography.Text strong style={{ color: "#fff", fontSize: 16, display: "block" }}>
                            {toolText(tool, "name")}
                          </Typography.Text>
                          {tool.category && (
                            <Tag style={{ marginTop: 4, border: "none", background: "rgba(255,255,255,0.2)", color: "#fff", borderRadius: 6, fontSize: 11 }}>
                              {toolText(tool, "category")}
                            </Tag>
                          )}
                        </div>
                      </div>
                      {/* Body */}
                      <div style={{ padding: "16px 20px 20px" }}>
                        <Typography.Paragraph
                          style={{ color: "#666", margin: 0, fontSize: 13, minHeight: 40 }}
                          ellipsis={{ rows: 2 }}
                        >
                          {toolText(tool, "description") || t("common", "noDescription")}
                        </Typography.Paragraph>
                        <div style={{ marginTop: 14, display: "flex", justifyContent: "flex-end" }}>
                          <Typography.Text style={{ color: "#1677ff", fontSize: 13, fontWeight: 500 }}>
                            {t("common", "open")} <ArrowRightOutlined style={{ fontSize: 11 }} />
                          </Typography.Text>
                        </div>
                      </div>
                    </Card>
                  </Col>
                );
              })}
            </Row>
          )}
        </div>
      </Content>
    </Layout>
  );
}
