import { useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import {
  Layout,
  Typography,
  Card,
  Upload,
  Button,
  Form,
  Input,
  Select,
  message,
  Result,
  Spin,
  Space,
  Divider,
  Alert,
  Progress,
} from "antd";
import {
  UploadOutlined,
  PlayCircleOutlined,
  DownloadOutlined,
  ArrowLeftOutlined,
  ToolOutlined,
} from "@ant-design/icons";
import type { UploadFile } from "antd";
import api from "../api/axios";
import Navbar from "../components/Navbar";
import { useLanguage } from "../context/LanguageContext";

const { Content } = Layout;
const { Title, Text } = Typography;

interface Tool {
  id: string;
  name: string;
  description: string | null;
  icon: string | null;
  category: string | null;
  script_path: string;
}

interface Execution {
  id: string;
  status: string;
  error_message: string | null;
}

// Tool-specific parameter definitions
const TOOL_PARAMS: Record<
  string,
  { label: string; name: string; type: "text" | "select"; options?: string[]; defaultValue?: string; required?: boolean }[]
> = {
  convert_espf_to_rds: [
    { label: "Périodicité", name: "periodicity", type: "select", options: ["monthly", "quarterly", "annual"], defaultValue: "monthly" },
    { label: "Code institution", name: "institution_code", type: "text", defaultValue: "RUBA" },
  ],
  forms_writer: [
    { label: "Mode", name: "mode", type: "select", options: ["eba", "report"], defaultValue: "eba", required: true },
    { label: "Code formulaire", name: "form_code", type: "text" },
    { label: "Code groupe formulaire", name: "form_group_code", type: "text" },
    { label: "Version", name: "version", type: "text" },
    { label: "Date de validité", name: "validity_starting_date", type: "text" },
    { label: "Domaine lié", name: "linked_domain", type: "text" },
    { label: "Type de données", name: "datatype", type: "text", defaultValue: "euebacorep.positiveOrNullAmount" },
    { label: "Unité d'arrondi", name: "rounding_unit", type: "text" },
  ],
  init_project: [
    { label: "Mode d'import", name: "import_mode", type: "select", options: ["manual", "jira_pdf"], defaultValue: "manual", required: true },
    { label: "Code réglementaire", name: "regulatory_code", type: "text", required: true },
    { label: "Type de projet", name: "project_type", type: "select", options: ["Regulation", "Output", "Both"], defaultValue: "Both" },
    { label: "Nom du projet Git", name: "git_project_name", type: "text", required: true },
    { label: "Version", name: "version", type: "text", defaultValue: "2026.1.0" },
  ],
};

const FILE_HINTS: Record<string, string> = {
  convert_espf_to_rds: "Fichier ESPF (.sXX) — format fixe, encodage latin-1",
  forms_writer: "Fichier source (.xlsx pour EBA, .xlsm pour Report). Optionnel : fichier de règles de validation (.xlsx)",
  init_project: "Mode Jira : fichier PDF exporté d'un ticket Jira (XRAY). Mode Manuel : aucun fichier requis.",
};

function FileUploadSection({
  tool,
  form,
  fileHint,
  fileList,
  setFileList,
}: {
  tool: Tool;
  form: any;
  fileHint: string;
  fileList: UploadFile[];
  setFileList: (files: UploadFile[]) => void;
}) {
  const importMode = Form.useWatch("import_mode", form);

  const showUpload =
    tool.script_path !== "init_project" || importMode === "jira_pdf";

  if (!showUpload) return null;

  const hint =
    tool.script_path === "init_project" && importMode === "jira_pdf"
      ? "Fichier PDF exporté d'un ticket Jira (XRAY Test Set PURE_RQM-*)"
      : fileHint;

  return (
    <>
      <Divider style={{ borderColor: "#f0e6f6" }} />
      <Title level={5} style={{ color: "#a9516a" }}>
        Fichiers d'entrée
      </Title>
      {hint && (
        <Alert
          message={hint}
          type="info"
          showIcon
          style={{ marginBottom: 16, borderRadius: 10, border: "1px solid #e8d5f5" }}
        />
      )}
      <Upload.Dragger
        multiple={tool.script_path !== "init_project"}
        fileList={fileList}
        beforeUpload={() => false}
        onChange={({ fileList: newList }) => setFileList(newList)}
        accept={tool.script_path === "init_project" ? ".pdf" : undefined}
        style={{
          marginBottom: 24,
          borderRadius: 14,
          border: "2px dashed #9369b8",
          background: "#ffffff",
        }}
      >
        <p className="ant-upload-drag-icon">
          <UploadOutlined style={{ color: "#9369b8", fontSize: 36 }} />
        </p>
        <p className="ant-upload-text" style={{ color: "#a9516a", fontWeight: 500 }}>
          Cliquez ou glissez-déposez vos fichiers ici
        </p>
      </Upload.Dragger>
    </>
  );
}

export default function ToolRun() {
  const { toolId } = useParams<{ toolId: string }>();
  const navigate = useNavigate();
  const { t } = useLanguage();
  const [form] = Form.useForm();
  const [tool, setTool] = useState<Tool | null>(null);
  const [loading, setLoading] = useState(true);
  const [executing, setExecuting] = useState(false);
  const [execution, setExecution] = useState<Execution | null>(null);
  const [fileList, setFileList] = useState<UploadFile[]>([]);

  useEffect(() => {
    if (!toolId) return;
    api
      .get(`/api/tools/${toolId}`)
      .then((res) => setTool(res.data))
      .catch(() => message.error(t("toolRun", "loadError")))
      .finally(() => setLoading(false));
  }, [toolId]);

  const handleExecute = async () => {
    try {
      await form.validateFields();
    } catch {
      return;
    }

    const values = form.getFieldsValue();
    const scriptPath = tool?.script_path || "";

    const needsFile =
      scriptPath !== "init_project" ||
      values.import_mode === "jira_pdf";
    if (needsFile && fileList.length === 0) {
      message.warning(t("toolRun", "selectedFiles"));
      return;
    }

    setExecuting(true);
    setExecution(null);

    try {
      const formData = new FormData();
      formData.append("tool_id", tool!.id);
      formData.append("parameters", JSON.stringify(values));

      fileList.forEach((file) => {
        if (file.originFileObj) {
          formData.append("files", file.originFileObj);
        }
      });

      const res = await api.post("/api/executions/", formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });

      setExecution(res.data);

      if (res.data.status === "completed") {
        message.success(t("toolRun", "successTitle"));
      } else if (res.data.status === "failed") {
        message.error(t("toolRun", "errorTitle"));
      }
    } catch (err: any) {
      message.error(err.response?.data?.detail || t("toolRun", "executeError"));
    } finally {
      setExecuting(false);
    }
  };

  const handleDownload = async () => {
    if (!execution) return;
    try {
      const res = await api.get(`/api/executions/${execution.id}/download`, {
        responseType: "blob",
      });

      const contentDisposition = res.headers["content-disposition"];
      let filename = "result";
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
      message.error(t("toolRun", "downloadError"));
    }
  };

  if (loading) {
    return (
      <Layout style={{ minHeight: "100vh", background: "#fff8f1" }}>
        <Navbar />
        <Content style={{ display: "flex", justifyContent: "center", alignItems: "center" }}>
          <Spin size="large" />
        </Content>
      </Layout>
    );
  }

  if (!tool) {
    return (
      <Layout style={{ minHeight: "100vh", background: "#fff8f1" }}>
        <Navbar />
        <Content style={{ padding: "32px 48px" }}>
          <Result status="404" title={t("toolRun", "toolNotFound")} subTitle={t("toolRun", "toolNotFound")} extra={<Button onClick={() => navigate("/home")}>{t("common", "backToTools")}</Button>} />
        </Content>
      </Layout>
    );
  }

  const params = TOOL_PARAMS[tool.script_path] || [];
  const fileHint = FILE_HINTS[tool.script_path] || "";

  return (
    <Layout style={{ minHeight: "100vh", background: "#fff8f1" }}>
      <Navbar />
      <Content style={{ padding: 0 }}>
        {/* Hero header */}
        <div
          style={{
            background: "linear-gradient(135deg, #a9516a 0%, #d77fa3 40%, #ff9800 100%)",
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
          <div style={{ maxWidth: 900, margin: "0 auto" }}>
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
                <ToolOutlined style={{ fontSize: 24, color: "#fff" }} />
              </div>
              <div>
                <Title level={3} style={{ color: "#fff", margin: 0, fontWeight: 700 }}>
                  {tool.name}
                </Title>
                {tool.description && (
                  <Text style={{ color: "rgba(255,255,255,0.6)", fontSize: 13 }}>
                    {tool.description}
                  </Text>
                )}
              </div>
            </div>
          </div>
        </div>

        <div style={{ maxWidth: 900, margin: "0 auto", padding: "32px 48px 48px" }}>
          <Card
            style={{
              borderRadius: 16,
              border: "none",
              boxShadow: "0 4px 24px rgba(108,44,145,0.08)",
              marginTop: -32,
            }}
          >
            {/* Parameters form */}
            {params.length > 0 && (
              <>
                <Title level={5} style={{ color: "#c2185b" }}>
                  Paramètres
                </Title>
                <Form form={form} layout="vertical" onValuesChange={() => setExecution(null)}>
                  {params.map((param) => (
                    <Form.Item
                      key={param.name}
                      name={param.name}
                      label={<span style={{ fontWeight: 500 }}>{param.label}</span>}
                      initialValue={param.defaultValue || ""}
                      rules={param.required ? [{ required: true, message: `${param.label} est requis` }] : []}
                    >
                      {param.type === "select" && param.options ? (
                        <Select style={{ borderRadius: 8 }}>
                          {param.options.map((opt) => (
                            <Select.Option key={opt} value={opt}>
                              {opt}
                            </Select.Option>
                          ))}
                        </Select>
                      ) : (
                        <Input style={{ borderRadius: 8 }} />
                      )}
                    </Form.Item>
                  ))}
                </Form>
              </>
            )}

            {/* File upload section */}
            <FileUploadSection
              tool={tool}
              form={form}
              fileHint={fileHint}
              fileList={fileList}
              setFileList={setFileList}
            />

            <Divider style={{ borderColor: "#f0e6f6" }} />

            {/* Execute button */}
            <Space direction="vertical" style={{ width: "100%" }}>
              <Button
                type="primary"
                size="large"
                icon={<PlayCircleOutlined />}
                onClick={handleExecute}
                loading={executing}
                block
                style={{
                  height: 48,
                  borderRadius: 12,
                  background: "linear-gradient(135deg, #a9516a, #ff9800)",
                  border: "none",
                  fontWeight: 600,
                  fontSize: 15,
                  boxShadow: "0 4px 16px rgba(108,44,145,0.25)",
                }}
              >
                {executing ? t("toolRun", "executing") : t("toolRun", "executeTool")}
              </Button>

              {executing && (
                <Progress
                  percent={99}
                  status="active"
                  showInfo={false}
                  strokeColor={{ from: "#a9516a", to: "#ff9800" }}
                />
              )}

              {/* Result section */}
              {execution?.status === "completed" && (
                <Result
                  status="success"
                  title={t("toolRun", "successTitle")}
                  style={{ padding: "24px 0" }}
                  extra={
                    <Button
                      type="primary"
                      icon={<DownloadOutlined />}
                      size="large"
                      onClick={handleDownload}
                      style={{
                        borderRadius: 12,
                        background: "linear-gradient(135deg, #c2185b, #ff9800)",
                        border: "none",
                        fontWeight: 600,
                        boxShadow: "0 4px 16px rgba(108,44,145,0.25)",
                      }}
                    >
                      {t("toolRun", "downloadResult")}
                    </Button>
                  }
                />
              )}

              {execution?.status === "failed" && (
                <Result
                  status="error"
                  title={t("toolRun", "errorTitle")}
                  subTitle={execution.error_message || t("common", "error")}
                  style={{ padding: "24px 0" }}
                />
              )}
            </Space>
          </Card>
        </div>
      </Content>
    </Layout>
  );
}
