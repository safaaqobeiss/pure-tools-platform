import { useEffect, useRef, useState } from "react";
import { Badge, Dropdown, List, Typography, Button, Empty } from "antd";
import {
  BellOutlined,
  CheckCircleOutlined,
  CloseCircleOutlined,
} from "@ant-design/icons";
import api from "../api/axios";
import { useAuth } from "../context/AuthContext";
import { useLanguage } from "../context/LanguageContext";

interface Notification {
  id: string;
  tool_name: string | null;
  status: string;
  created_at: string;
  error_message: string | null;
}

export default function NotificationBell() {
  const { isAuthenticated } = useAuth();
  const { t, lang } = useLanguage();
  const [notifications, setNotifications] = useState<Notification[]>([]);
  const [unreadCount, setUnreadCount] = useState(0);
  const seenIdsRef = useRef<Set<string>>(new Set());
  const [open, setOpen] = useState(false);

  useEffect(() => {
    if (!isAuthenticated) return;

    const poll = async () => {
      try {
        const res = await api.get("/api/executions/notifications/recent");
        const items: Notification[] = res.data;

        // Find new ones
        const newItems = items.filter((n) => !seenIdsRef.current.has(n.id));
        if (newItems.length > 0) {
          setUnreadCount((c) => c + newItems.length);
          newItems.forEach((n) => seenIdsRef.current.add(n.id));
        }

        setNotifications(items);
      } catch {
        // silent
      }
    };

    poll();
    const interval = setInterval(poll, 10000); // poll every 10s
    return () => clearInterval(interval);
  }, [isAuthenticated]);

  const handleOpen = (visible: boolean) => {
    setOpen(visible);
    if (visible) setUnreadCount(0);
  };

  const menu = (
    <div
      style={{
        background: "white",
        borderRadius: 8,
        boxShadow: "0 6px 16px rgba(0,0,0,0.12)",
        width: 360,
        maxHeight: 400,
        overflow: "auto",
      }}
    >
      <div
        style={{
          padding: "12px 16px",
          borderBottom: "1px solid #f0f0f0",
          fontWeight: 600,
        }}
      >
        {t("notifications", "title")}
      </div>
      {notifications.length === 0 ? (
        <div style={{ padding: 24 }}>
          <Empty
            description={t("notifications", "none")}
            image={Empty.PRESENTED_IMAGE_SIMPLE}
          />
        </div>
      ) : (
        <List
          dataSource={notifications}
          renderItem={(item) => (
            <List.Item style={{ padding: "10px 16px" }}>
              <List.Item.Meta
                avatar={
                  item.status === "completed" ? (
                    <CheckCircleOutlined
                      style={{ fontSize: 20, color: "#52c41a" }}
                    />
                  ) : (
                    <CloseCircleOutlined
                      style={{ fontSize: 20, color: "#ff4d4f" }}
                    />
                  )
                }
                title={
                  <Typography.Text>
                    {item.tool_name || "Tool"} —{" "}
                    {item.status === "completed" ? t("notifications", "completed") : t("notifications", "failed")}
                  </Typography.Text>
                }
                description={
                  <Typography.Text type="secondary" style={{ fontSize: 12 }}>
                    {new Date(item.created_at).toLocaleString(lang === "fr" ? "fr-FR" : "en-US", {
                      day: "2-digit",
                      month: "2-digit",
                      year: "numeric",
                      hour: "2-digit",
                      minute: "2-digit",
                    })}
                    {item.error_message && (
                      <span style={{ color: "#ff4d4f" }}>
                        {" "}
                        — {item.error_message.slice(0, 60)}
                      </span>
                    )}
                  </Typography.Text>
                }
              />
            </List.Item>
          )}
        />
      )}
    </div>
  );

  return (
    <Dropdown
      dropdownRender={() => menu}
      trigger={["click"]}
      open={open}
      onOpenChange={handleOpen}
      placement="bottomRight"
    >
      <Button
        type="text"
        style={{ color: "white", padding: "0 8px" }}
        icon={
          <Badge count={unreadCount} size="small" offset={[2, -2]}>
            <BellOutlined style={{ fontSize: 18, color: "white" }} />
          </Badge>
        }
      />
    </Dropdown>
  );
}
