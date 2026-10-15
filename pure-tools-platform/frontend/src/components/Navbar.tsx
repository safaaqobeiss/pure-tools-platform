import { Layout, Button, Typography, Space, Avatar, Dropdown } from "antd";
import {
  LogoutOutlined,
  UserOutlined,
  HomeOutlined,
  HistoryOutlined,
  DashboardOutlined,
  SettingOutlined,
  MenuOutlined,
  GlobalOutlined,
} from "@ant-design/icons";
import { useAuth } from "../context/AuthContext";
import { useLanguage } from "../context/LanguageContext";
import { useNavigate, useLocation } from "react-router-dom";
import NotificationBell from "./NotificationBell";
import SbsLogo from "./SbsLogo";
import api from "../api/axios";

const { Header } = Layout;

export default function Navbar() {
  const { user, logout } = useAuth();
  const { lang, setLang, t } = useLanguage();
  const navigate = useNavigate();
  const location = useLocation();

  const NAV_ITEMS = [
    { key: "/home", icon: <HomeOutlined />, label: t("nav", "tools") },
    { key: "/history", icon: <HistoryOutlined />, label: t("nav", "history") },
    { key: "/dashboard", icon: <DashboardOutlined />, label: t("nav", "dashboard") },
  ];

  const handleLogout = () => {
    logout();
    navigate("/login");
  };

  const isActive = (path: string) => location.pathname === path;

  const initials = user?.full_name
    ? user.full_name
        .split(" ")
        .map((w: string) => w[0])
        .join("")
        .toUpperCase()
        .slice(0, 2)
    : "?";

  const userMenuItems = [
    {
      key: "profile",
      icon: <UserOutlined />,
      label: t("nav", "myProfile"),
      onClick: () => navigate("/profile"),
    },
    ...(user?.role === "admin"
      ? [
          {
            key: "admin",
            icon: <SettingOutlined />,
            label: t("nav", "administration"),
            onClick: () => navigate("/admin"),
          },
        ]
      : []),
    { type: "divider" as const },
    {
      key: "logout",
      icon: <LogoutOutlined />,
      label: t("nav", "logout"),
      danger: true,
      onClick: handleLogout,
    },
  ];

  return (
    <Header
      style={{
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        background: "linear-gradient(90deg, #c2185b 0%, #e91e90 50%, #ff9800 100%)",
        padding: "0 24px",
        boxShadow: "0 3px 16px rgba(194, 24, 91, 0.35)",
        height: 58,
        lineHeight: "58px",
        position: "sticky",
        top: 0,
        zIndex: 100,
        borderBottom: "2px solid rgba(255,255,255,0.08)",
      }}
    >
      {/* Logo + brand */}
      <div
        style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer" }}
        onClick={() => navigate("/home")}
      >
        <div
          style={{
            background: "rgba(255,255,255,0.12)",
            borderRadius: 12,
            padding: 4,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            backdropFilter: "blur(4px)",
          }}
        >
          <SbsLogo size={34} />
        </div>
        <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.2 }}>
          <Typography.Text
            strong
            style={{ color: "#fff", fontSize: 15, letterSpacing: 0.8 }}
          >
            Pure Tools
          </Typography.Text>
          <Typography.Text
            style={{ color: "rgba(255,255,255,0.45)", fontSize: 10, letterSpacing: 0.5 }}
          >
            by SBS
          </Typography.Text>
        </div>
      </div>

      {/* Center nav — pill-shaped buttons */}
      <div
        style={{
          display: "flex",
          gap: 4,
          background: "rgba(255,255,255,0.08)",
          borderRadius: 12,
          padding: "4px 6px",
        }}
      >
        {NAV_ITEMS.map((item) => (
          <Button
            key={item.key}
            type="text"
            icon={item.icon}
            onClick={() => navigate(item.key)}
            style={{
              color: isActive(item.key) ? "#fff" : "rgba(255,255,255,0.65)",
              background: isActive(item.key)
                ? "linear-gradient(135deg, #ffca28, #ff9800)"
                : "transparent",
              borderRadius: 10,
              fontWeight: isActive(item.key) ? 600 : 400,
              transition: "all 0.25s ease",
              border: "none",
              boxShadow: isActive(item.key) ? "0 2px 8px rgba(255, 202, 40, 0.3)" : "none",
              padding: "4px 14px",
              height: 34,
            }}
          >
            {item.label}
          </Button>
        ))}
      </div>

      {/* Right side */}
      <Space size={12} align="center">
        <Button
          type="text"
          icon={<GlobalOutlined />}
          onClick={() => setLang(lang === "fr" ? "en" : "fr")}
          title={lang === "fr" ? "Switch to English" : "Passer en français"}
          style={{
            color: "#fff",
            background: "rgba(255,255,255,0.1)",
            borderRadius: 8,
            fontWeight: 700,
            fontSize: 12,
            height: 32,
            padding: "0 10px",
            border: "1px solid rgba(255,255,255,0.15)",
          }}
        >
          {lang.toUpperCase()}
        </Button>
        <NotificationBell />
        <Dropdown menu={{ items: userMenuItems }} trigger={["click"]} placement="bottomRight">
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              cursor: "pointer",
              padding: "4px 10px",
              borderRadius: 10,
              transition: "background 0.2s",
            }}
            onMouseEnter={(e) => (e.currentTarget.style.background = "rgba(255,255,255,0.1)")}
            onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
          >
            {user?.avatar_url ? (
              <Avatar size={30} src={`${api.defaults.baseURL || ""}${user.avatar_url}`} />
            ) : (
              <Avatar
                size={30}
                style={{
                  background: "linear-gradient(135deg, #e91e90, #ff9800)",
                  fontSize: 13,
                  fontWeight: 600,
                }}
              >
                {initials}
              </Avatar>
            )}
            <span style={{ color: "rgba(255,255,255,0.9)", fontSize: 13, maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {user?.full_name || user?.email}
            </span>
            <MenuOutlined style={{ color: "rgba(255,255,255,0.5)", fontSize: 10 }} />
          </div>
        </Dropdown>
      </Space>
    </Header>
  );
}
