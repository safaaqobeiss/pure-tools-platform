import { useEffect, useRef, useState } from "react";
import { useNavigate } from "react-router-dom";
import {
  Layout,
  Typography,
  Card,
  Button,
  Form,
  Input,
  Steps,
  Upload,
  Checkbox,
  Descriptions,
  Result,
  Spin,
  Space,
  Divider,
  Alert,
  message,
  Progress,
} from "antd";
import {
  ArrowLeftOutlined,
  UploadOutlined,
  FileExcelOutlined,
  BookOutlined,
  SettingOutlined,
  RocketOutlined,
  DownloadOutlined,
  CheckCircleOutlined,
} from "@ant-design/icons";
import type { UploadFile } from "antd";
import api from "../api/axios";
import Navbar from "../components/Navbar";
import { useLanguage } from "../context/LanguageContext";

const { Content } = Layout;
const { Title, Text, Paragraph } = Typography;

type StepKey = "upload_dpm" | "upload_dict" | "parameters" | "result";

interface GenerationResult {
  id: string;
  status: string;
  output_files: string[];
  stdout: string;
  error_message?: string;
}

export default function FormsGeneratorWizard() {
  const navigate = useNavigate();
  const { t } = useLanguage();
  const [form] = Form.useForm();
  const [currentStep, setCurrentStep] = useState(0);
  const [dpmFile, setDpmFile] = useState<UploadFile | null>(null);
  const [dictFile, setDictFile] = useState<UploadFile | null>(null);
  const [autoEnrich, setAutoEnrich] = useState(true);
  const [autoRowkey, setAutoRowkey] = useState(true);
  const [executing, setExecuting] = useState(false);
  const [result, setResult] = useState<GenerationResult | null>(null);
  const pollRef = useRef<number | null>(null);

  const stopPolling = () => {
    if (pollRef.current !== null) {
      window.clearInterval(pollRef.current);
      pollRef.current = null;
    }
  };

  // Arrête le suivi si on quitte la page
  useEffect(() => stopPolling, []);

  const STEPS: { key: StepKey; title: string; icon: React.ReactNode }[] = [
    { key: "upload_dpm", title: t("formsGenerator", "stepDpm"), icon: <FileExcelOutlined /> },
    { key: "upload_dict", title: t("formsGenerator", "stepDict"), icon: <BookOutlined /> },
    { key: "parameters", title: t("formsGenerator", "stepParams"), icon: <SettingOutlined /> },
    { key: "result", title: t("formsGenerator", "stepResult"), icon: <CheckCircleOutlined /> },
  ];

  const currentKey = STEPS[currentStep].key;

  const canNext = (): boolean => {
    switch (currentKey) {
      case "upload_dpm":
        return dpmFile !== null;
      case "upload_dict":
        return dictFile !== null;
      case "parameters":
        return true;
      default:
        return false;
    }
  };

  const handleNext = async () => {
    if (currentKey === "parameters") {
      try {
        await form.validateFields();
      } catch {
        return;
      }
      await handleGenerate();
      return;
    }
    setCurrentStep((s) => s + 1);
  };

  const finishResult = (data: GenerationResult) => {
    setResult(data);
    if (data.status === "completed") {
      message.success(t("formsGenerator", "generationSuccess"));
    } else {
      message.error(t("formsGenerator", "generationFailed"));
    }
  };

  // Interroge le backend toutes les 3 s jusqu'à la fin de la génération
  const startPolling = (id: string) => {
    stopPolling();
    const startedAt = Date.now();
    const maxDurationMs = 30 * 60 * 1000; // 30 minutes

    pollRef.current = window.setInterval(async () => {
      try {
        const res = await api.get(`/api/forms-generator/${id}/status`);
        const s = res.data;

        if (s.status === "running") {
          if (Date.now() - startedAt > maxDurationMs) {
            stopPolling();
            setExecuting(false);
            finishResult({
              id,
              status: "failed",
              output_files: [],
              stdout: "",
              error_message: "La génération prend trop de temps (plus de 30 minutes).",
            });
          }
          return;
        }

        stopPolling();
        setExecuting(false);
        finishResult({
          id,
          status: s.status,
          output_files: s.output_files || [],
          stdout: "",
          error_message: s.error_message || undefined,
        });
      } catch (err: any) {
        const code = err.response?.status;
        // Erreur définitive : on arrête. Sinon (réseau), on réessaie au prochain tour.
        if (code === 401 || code === 404) {
          stopPolling();
          setExecuting(false);
          finishResult({
            id,
            status: "failed",
            output_files: [],
            stdout: "",
            error_message:
              err.response?.data?.detail || "Impossible de suivre la génération.",
          });
        }
      }
    }, 3000);
  };

  const handleGenerate = async () => {
    if (!dpmFile?.originFileObj || !dictFile?.originFileObj) {
      message.error(t("formsGenerator", "missingFiles"));
      return;
    }

    const values = form.getFieldsValue();
    const params = {
      form_code: values.form_code || "",
      form_group_code: values.form_group_code || "",
      version: values.version || "",
      validity_starting_date: values.validity_starting_date || "",
      linked_domain: values.linked_domain || "",
      datatype: values.datatype || "",
      rounding_unit: values.rounding_unit || "",
      auto_enrich: autoEnrich,
      auto_rowkey: autoRowkey,
    };

    stopPolling();
    setResult(null);
    setExecuting(true);
    setCurrentStep(3);

    let polling = false;
    try {
      const formData = new FormData();
      formData.append("parameters", JSON.stringify(params));
      formData.append("dpm_file", dpmFile.originFileObj);
      formData.append("dict_file", dictFile.originFileObj);

      const res = await api.post("/api/forms-generator/generate", formData, {
        headers: { "Content-Type": "multipart/form-data" },
        timeout: 600000,
      });

      const data: GenerationResult = res.data;
      if (data.status === "running" && data.id) {
        // La génération tourne en arrière-plan : on attend la fin
        polling = true;
        startPolling(data.id);
      } else {
        finishResult(data);
      }
    } catch (err: any) {
      message.error(
        err.response?.data?.detail || t("formsGenerator", "executeError")
      );
      setResult({
        id: "",
        status: "failed",
        output_files: [],
        stdout: "",
        error_message: err.response?.data?.detail || String(err),
      });
    } finally {
      if (!polling) setExecuting(false);
    }
  };

  const handleDownload = async () => {
    if (!result?.id) return;
    try {
      const res = await api.get(
        `/api/forms-generator/${result.id}/download`,
        { responseType: "blob" }
      );

      const contentDisposition = res.headers["content-disposition"];
      let filename = `pure-forms-${result.id.substring(0, 8)}.zip`;
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
      message.error(t("formsGenerator", "downloadError"));
    }
  };

  const handleReset = () => {
    stopPolling();
    setExecuting(false);
    setCurrentStep(0);
    setDpmFile(null);
    setDictFile(null);
    setAutoEnrich(true);
    setAutoRowkey(true);
    setResult(null);
    form.resetFields();
  };

  // ── Step renderers ──────────────────────────────────────────────

  const renderUploadDpm = () => (
    <div>
      <Title level={5} style={{ color: "#9369b8" }}>
        <FileExcelOutlined style={{ marginRight: 8 }} />
        {t("formsGenerator", "dpmSourceFile")}
      </Title>
      <Paragraph type="secondary">
        {t("formsGenerator", "dpmSourceDesc")}
      </Paragraph>
      <Upload.Dragger
        maxCount={1}
        fileList={dpmFile ? [dpmFile] : []}
        beforeUpload={() => false}
        onChange={({ fileList }) =>
          setDpmFile(fileList.length > 0 ? fileList[0] : null)
        }
        accept=".xlsx,.xls,.xlsm"
        style={{
          marginTop: 16,
          borderRadius: 14,
          border: "2px dashed #e3f2fd",
          background: "#ffffff",
        }}
      >
        <p className="ant-upload-drag-icon">
          <UploadOutlined style={{ color: "#9369b8", fontSize: 36 }} />
        </p>
        <p className="ant-upload-text" style={{ color: "#9369b8", fontWeight: 500 }}>
          {t("formsGenerator", "dragDpmHere")}
        </p>
        <p className="ant-upload-hint" style={{ color: "#999" }}>
          {t("formsGenerator", "dpmAcceptedFormat")}
        </p>
      </Upload.Dragger>
    </div>
  );

  const renderUploadDict = () => (
    <div>
      <Title level={5} style={{ color: "#ba68c8" }}>
        <BookOutlined style={{ marginRight: 8 }} />
        {t("formsGenerator", "dictFile")}
      </Title>
      <Paragraph type="secondary">
        {t("formsGenerator", "dictFileDesc")}
      </Paragraph>
      <Upload.Dragger
        maxCount={1}
        fileList={dictFile ? [dictFile] : []}
        beforeUpload={() => false}
        onChange={({ fileList }) =>
          setDictFile(fileList.length > 0 ? fileList[0] : null)
        }
        accept=".yml,.yaml"
        style={{
          marginTop: 16,
          borderRadius: 14,
          border: "2px dashed #f8bbd0",
          background: "#ffffff",
        }}
      >
        <p className="ant-upload-drag-icon">
          <UploadOutlined style={{ color: "#ba68c8", fontSize: 36 }} />
        </p>
        <p className="ant-upload-text" style={{ color: "#ba68c8", fontWeight: 500 }}>
          {t("formsGenerator", "dragDictHere")}
        </p>
        <p className="ant-upload-hint" style={{ color: "#999" }}>
          {t("formsGenerator", "dictAcceptedFormat")}
        </p>
      </Upload.Dragger>
    </div>
  );

  const renderParameters = () => (
    <div>
      <Title level={5} style={{ color: "#ba68c8" }}>
        <SettingOutlined style={{ marginRight: 8 }} />
        {t("formsGenerator", "genParams")}
      </Title>

      {dpmFile && dictFile && (
        <Alert
          type="info"
          showIcon
          style={{ marginBottom: 16, borderRadius: 10, border: "1px solid #f8bbd0" }}
          message={t("formsGenerator", "selectedFiles")}
          description={
            <div>
              <div>
                <strong>DPM :</strong> {dpmFile.name}
              </div>
              <div>
                <strong>Dictionnaire :</strong> {dictFile.name}
              </div>
            </div>
          }
        />
      )}

      <Form form={form} layout="vertical">
        <Form.Item
          name="form_code"
          label={<span style={{ fontWeight: 500 }}>formCode</span>}
          tooltip="Code du formulaire (ex: corep, finrep)"
        >
          <Input placeholder="ex: corep" style={{ borderRadius: 8 }} />
        </Form.Item>

        <Form.Item
          name="form_group_code"
          label={<span style={{ fontWeight: 500 }}>formGroupCode</span>}
          tooltip="Code du groupe de formulaires"
        >
          <Input placeholder="optionnel" style={{ borderRadius: 8 }} />
        </Form.Item>

        <Form.Item
          name="version"
          label={<span style={{ fontWeight: 500 }}>Version</span>}
        >
          <Input placeholder="ex: 2026.1.0" style={{ borderRadius: 8 }} />
        </Form.Item>

        <Form.Item
          name="validity_starting_date"
          label={<span style={{ fontWeight: 500 }}>{t("formsGenerator", "validityDate")}</span>}
        >
          <Input placeholder="ex: 01/01/2026" style={{ borderRadius: 8 }} />
        </Form.Item>

        <Form.Item
          name="linked_domain"
          label={<span style={{ fontWeight: 500 }}>{t("formsGenerator", "linkedDomain")}</span>}
        >
          <Input placeholder="ex: euebacorep" style={{ borderRadius: 8 }} />
        </Form.Item>

        <Form.Item
          name="datatype"
          label={<span style={{ fontWeight: 500 }}>{t("formsGenerator", "datatypeLabel")}</span>}
          tooltip="Datatype par défaut pour la feuille identification"
        >
          <Input placeholder="ex: euebacorep.positiveOrNullAmount" style={{ borderRadius: 8 }} />
        </Form.Item>

        <Form.Item
          name="rounding_unit"
          label={<span style={{ fontWeight: 500 }}>{t("formsGenerator", "roundingUnit")}</span>}
        >
          <Input placeholder="ex: thousands" style={{ borderRadius: 8 }} />
        </Form.Item>

        <Divider style={{ borderColor: "#f8bbd0" }} />

        <Title level={5} style={{ color: "#9369b8" }}>{t("formsGenerator", "pipelineOptions")}</Title>
        <Space direction="vertical">
          <Checkbox
            checked={autoEnrich}
            onChange={(e) => setAutoEnrich(e.target.checked)}
          >
            {t("formsGenerator", "autoEnrich")}
          </Checkbox>
          <Checkbox
            checked={autoRowkey}
            onChange={(e) => setAutoRowkey(e.target.checked)}
          >
            {t("formsGenerator", "autoRowkey")}
          </Checkbox>
        </Space>
      </Form>
    </div>
  );

  const renderResult = () => {
    if (executing) {
      return (
        <div style={{ textAlign: "center", padding: "60px 0" }}>
          <Spin size="large" />
          <Title level={4} style={{ marginTop: 24, color: "#9369b8" }}>
            {t("formsGenerator", "generating")}
          </Title>
          <Paragraph type="secondary">
            {t("formsGenerator", "subtitle")}
          </Paragraph>
          <Progress
            percent={99}
            status="active"
            showInfo={false}
            strokeColor={{ from: "#d77fa3", to: "#ff9800" }}
          />
        </div>
      );
    }

    if (!result) return null;

    if (result.status === "completed") {
      return (
        <Result
          status="success"
          title={t("formsGenerator", "generationSuccess")}
          subTitle={`${result.output_files.length} ${t("formsGenerator", "filesGenerated")}`}
          extra={[
            <Button
              key="download"
              type="primary"
              size="large"
              icon={<DownloadOutlined />}
              onClick={handleDownload}
              style={{
                borderRadius: 12,
                background: "linear-gradient(135deg, #d77fa3, #ff9800)",
                border: "none",
                fontWeight: 600,
                boxShadow: "0 4px 16px rgba(229, 57, 53, 0.25)",
              }}
            >
              {t("formsGenerator", "downloadResult")}
            </Button>,
            <Button
              key="reset"
              onClick={handleReset}
              style={{ borderRadius: 10 }}
            >
              {t("common", "reset")}
            </Button>,
          ]}
        >
          {result.output_files.length > 0 && (
            <Descriptions
              title={t("formsGenerator", "generatedFiles")}
              column={1}
              bordered
              size="small"
              style={{ marginTop: 16 }}
            >
              {result.output_files.map((f, i) => (
                <Descriptions.Item
                  key={i}
                  label={<FileExcelOutlined style={{ color: "#9369b8" }} />}
                >
                  {f}
                </Descriptions.Item>
              ))}
            </Descriptions>
          )}

          {result.stdout && (
            <div style={{ marginTop: 16 }}>
              <Divider>{t("formsGenerator", "logs")}</Divider>
              <pre
                style={{
                  background: "#fce4ec",
                  padding: 12,
                  borderRadius: 10,
                  maxHeight: 300,
                  overflow: "auto",
                  fontSize: 12,
                  whiteSpace: "pre-wrap",
                  border: "1px solid #f8bbd0",
                }}
              >
                {result.stdout}
              </pre>
            </div>
          )}
        </Result>
      );
    }

    return (
      <Result
        status="error"
        title={t("formsGenerator", "generationFailed")}
        subTitle={t("common", "error")}
        extra={[
          <Button
            key="retry"
            type="primary"
            onClick={handleReset}
            style={{
              borderRadius: 10,
              background: "linear-gradient(135deg, #e91e90, #ff9800)",
              border: "none",
              fontWeight: 600,
            }}
          >
            {t("common", "reset")}
          </Button>,
        ]}
      >
        {(result.error_message || result.stdout) && (
          <pre
            style={{
              background: "#ffffff",
              padding: 12,
              borderRadius: 10,
              maxHeight: 300,
              overflow: "auto",
              fontSize: 12,
              whiteSpace: "pre-wrap",
              textAlign: "left",
            }}
          >
            {result.error_message}
            {result.error_message && result.stdout ? "\n\n" : ""}
            {result.stdout}
          </pre>
        )}
      </Result>
    );
  };

  // ── Main render ─────────────────────────────────────────────────

  const stepRenderers: Record<StepKey, () => React.ReactNode> = {
    upload_dpm: renderUploadDpm,
    upload_dict: renderUploadDict,
    parameters: renderParameters,
    result: renderResult,
  };

  return (
    <Layout style={{ minHeight: "100vh", background: "#ffffff" }}>
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
          <div style={{ maxWidth: 1000, margin: "0 auto" }}>
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
                <FileExcelOutlined style={{ fontSize: 24, color: "#fff" }} />
              </div>
              <div>
                <Title level={3} style={{ color: "#fff", margin: 0, fontWeight: 700 }}>
                  {t("formsGenerator", "title")}
                </Title>
                <Text style={{ color: "rgba(255,255,255,0.6)", fontSize: 13 }}>
                  {t("formsGenerator", "subtitle")}
                </Text>
              </div>
            </div>
          </div>
        </div>

        <div style={{ maxWidth: 1000, margin: "0 auto", padding: "32px 48px 48px" }}>
          <Card
            style={{
              borderRadius: 16,
              border: "none",
              boxShadow: "0 4px 24px rgba(229, 57, 53, 0.08)",
              marginTop: -32,
            }}
          >
            <Steps
              current={currentStep}
              style={{ marginBottom: 32 }}
              items={STEPS.map((s) => ({
                title: <span style={{ fontWeight: 500 }}>{s.title}</span>,
                icon: (
                  <div
                    style={{
                      width: 32,
                      height: 32,
                      borderRadius: 10,
                      background:
                        currentStep >= STEPS.indexOf(s)
                          ? "linear-gradient(135deg, #d77fa3, #ff9800)"
                          : "#f8bbd0",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      color: currentStep >= STEPS.indexOf(s) ? "#fff" : "#d77fa3",
                      fontSize: 14,
                    }}
                  >
                    {s.icon}
                  </div>
                ),
              }))}
            />

            <div style={{ minHeight: 300 }}>
              {stepRenderers[currentKey]()}
            </div>

            {/* Navigation buttons */}
            {currentKey !== "result" && (
              <>
                <Divider style={{ borderColor: "#f8bbd0" }} />
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                  }}
                >
                  <Button
                    disabled={currentStep === 0}
                    onClick={() => setCurrentStep((s) => s - 1)}
                    style={{ borderRadius: 10, height: 40 }}
                  >
                    {t("common", "previous")}
                  </Button>
                  <Button
                    type="primary"
                    icon={
                      currentKey === "parameters" ? (
                        <RocketOutlined />
                      ) : undefined
                    }
                    disabled={!canNext()}
                    onClick={handleNext}
                    loading={executing}
                    style={{
                      borderRadius: 10,
                      height: 40,
                      background: canNext()
                        ? "linear-gradient(135deg, #e91e90, #ff9800)"
                        : undefined,
                      border: "none",
                      fontWeight: 600,
                      boxShadow: canNext() ? "0 4px 12px rgba(229,57,53,0.2)" : undefined,
                    }}
                  >
                    {currentKey === "parameters"
                      ? t("formsGenerator", "generate")
                      : t("common", "next")}
                  </Button>
                </div>
              </>
            )}
          </Card>
        </div>
      </Content>
    </Layout>
  );
}