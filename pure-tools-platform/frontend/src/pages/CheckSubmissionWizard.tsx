import { useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import {
  Layout,
  Typography,
  Card,
  Button,

  ConfigProvider,
  Upload,
  Radio,
  Input,
  Steps,
  Descriptions,
  Result,
  Space,
  Divider,
  Alert,
  message,
  Progress,
  Tag,
  Table,
} from "antd";
import {
  ArrowLeftOutlined,
  SafetyCertificateOutlined,
  FileSearchOutlined,
  FolderOpenOutlined,
  UploadOutlined,
  DownloadOutlined,
  CheckCircleOutlined,
  CloseCircleOutlined,
  WarningOutlined,
  BankOutlined,
  ExperimentOutlined,
} from "@ant-design/icons";
import type { UploadFile } from "antd";
import api from "../api/axios";
import Navbar from "../components/Navbar";
import { useLanguage } from "../context/LanguageContext";

const { Content } = Layout;
const { Title, Text, Paragraph } = Typography;

const REGULATIONS = ["RUBA", "COREP", "RESOL", "FP", "P3DH", "MCONTRAN", "DBM"];

interface InstitutionInfo {
  name: string;
  modes: string[];
}

interface ExpectedFileInfo {
  institution: string;
  filename: string;
  outputCode: string;
}

interface ValidationResult {
  execution_id: string;
  status: string;
  summary: {
    OK?: number;
    DIFF?: number;
    MISSING_IN_OUTPUT?: number;
    CALCULATED_OR_EXTRA?: number;
    TOTAL?: number;
  };
  cases_summary?: Array<{
    case_id: string;
    OK: number;
    DIFF: number;
    MISSING_IN_OUTPUT: number;
    CALCULATED_OR_EXTRA: number;
    TOTAL: number;
  }>;
  output_files: string[];
  error_message?: string;
}

/** Parse institution structure from uploaded folder file list */
function detectInstitutions(files: UploadFile[]): InstitutionInfo[] {
  const map = new Map<string, Set<string>>();
  for (const f of files) {
    const rel =
      (f.originFileObj as any)?.webkitRelativePath || (f as any).webkitRelativePath || "";
    const parts = rel.split("/");
    // parts[0] = root folder, parts[1] = institution, parts[2] = csv_direct|excel
    if (parts.length >= 3) {
      const inst = parts[1];
      const mode = parts[2];
      if (mode === "csv_direct" || mode === "excel") {
        if (!map.has(inst)) map.set(inst, new Set());
        map.get(inst)!.add(mode);
      }
    }
  }
  return Array.from(map.entries())
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([name, modes]) => ({ name, modes: Array.from(modes).sort() }));
}

/** Parse expected submission files from uploaded folder */
function detectExpectedFiles(files: UploadFile[]): ExpectedFileInfo[] {
  const result: ExpectedFileInfo[] = [];
  for (const f of files) {
    const name = f.name.toLowerCase();
    if (
      name.includes("submission") &&
      !name.includes("dataset") &&
      (name.endsWith(".xbrl") || name.endsWith(".zip") || name.endsWith(".xml"))
    ) {
      const stem = f.name.replace(/\.[^.]+$/, "");
      const parts = stem.split("_");
      const inst = parts[0] || f.name;
      const outputCode = parts.length >= 2 ? parts[1] : "";
      result.push({ institution: inst, filename: f.name, outputCode });
    }
  }
  return result.sort((a, b) => a.institution.localeCompare(b.institution));
}

export default function CheckSubmissionWizard() {
  const navigate = useNavigate();
  const { t } = useLanguage();
  const [step, setStep] = useState(0);

  // Step 0: Regulation
  const [regulation, setRegulation] = useState<string>("");
  const [customRegulation, setCustomRegulation] = useState<string>("");

  // Step 1: Input directory
  const [inputFiles, setInputFiles] = useState<UploadFile[]>([]);

  // Step 2: Expected directory
  const [expectedFiles, setExpectedFiles] = useState<UploadFile[]>([]);

  // Step 3: Taxonomy
  const [taxonomyFile, setTaxonomyFile] = useState<UploadFile[]>([]);

  // Execution state
  const [validating, setValidating] = useState(false);
  const [result, setResult] = useState<ValidationResult | null>(null);

  // Derived info
  const institutions = useMemo(() => detectInstitutions(inputFiles), [inputFiles]);
  const expectedInfo = useMemo(() => detectExpectedFiles(expectedFiles), [expectedFiles]);
  const regulationName = regulation === "OTHER" ? customRegulation.trim() : regulation;

  const canLaunch =
    regulationName.length > 0 &&
    inputFiles.length > 0 &&
    expectedFiles.length > 0 &&
    taxonomyFile.length > 0;

  // ── Navigation helpers ────────────────────
  const nextDisabled: Record<number, boolean> = {
    0: !regulationName,
    1: inputFiles.length === 0,
    2: expectedFiles.length === 0,
    3: taxonomyFile.length === 0,
  };

  const handleValidate = async () => {
    if (!canLaunch) return;

    setValidating(true);
    try {
      const formData = new FormData();

      formData.append("regulation", regulationName);

      for (const f of inputFiles) {
        if (f.originFileObj) {
          const relPath =
            (f.originFileObj as any).webkitRelativePath || f.name;
          formData.append("input_files", f.originFileObj, relPath);
        }
      }

      for (const f of expectedFiles) {
        if (f.originFileObj) {
          formData.append("expected_files", f.originFileObj, f.name);
        }
      }

      if (taxonomyFile[0]?.originFileObj) {
        formData.append("taxonomy_file", taxonomyFile[0].originFileObj);
      }

      const res = await api.post("/api/check-submission/validate", formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });

      setResult(res.data);
      setStep(5);

      if (res.data.status === "completed") {
        message.success(t("checkSubmission", "validationSuccess"));
      } else {
        message.error(t("checkSubmission", "validationFailed"));
      }
    } catch (err: any) {
      message.error(
        err.response?.data?.detail || t("checkSubmission", "validationError")
      );
    } finally {
      setValidating(false);
    }
  };

  const handleDownload = async () => {
    if (!result?.execution_id) return;
    try {
      const res = await api.get(
        `/api/executions/${result.execution_id}/download`,
        { responseType: "blob" }
      );
      const contentDisposition = res.headers["content-disposition"];
      let filename = "validation_report.zip";
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
      message.error(t("checkSubmission", "downloadError"));
    }
  };

  const handleReset = () => {
    setStep(0);
    setResult(null);
    setRegulation("");
    setCustomRegulation("");
    setInputFiles([]);
    setExpectedFiles([]);
    setTaxonomyFile([]);
  };

  // ── Shared styles ─────────────────────────
  const gradientBtn: React.CSSProperties = {
    borderRadius: 10,
    background: "linear-gradient(135deg, #e91e90, #ff9800)",
    border: "none",
  };
  const disabledBtn: React.CSSProperties = { borderRadius: 10, border: "none" };
  const navRow: React.CSSProperties = { display: "flex", justifyContent: "space-between", marginTop: 8 };

  return (
    <Layout style={{ minHeight: "100vh", background: "#ffffff" }}>
      <Navbar />
      <Content style={{ padding: 0 }}>
        {/* Hero header */}
        <div
          style={{
            background: "linear-gradient(135deg, #c2185b 0%, #e91e90 40%, #ff9800 100%)",
            padding: "36px 48px 32px",
            position: "relative",
            overflow: "hidden",
          }}
        >
          <div style={{ position: "absolute", top: -60, right: -40, width: 250, height: 250, borderRadius: "50%", background: "rgba(255,255,255,0.05)" }} />
          <div style={{ maxWidth: 1000, margin: "0 auto" }}>
            <Button
              icon={<ArrowLeftOutlined />}
              type="text"
              onClick={() => navigate("/home")}
              style={{ color: "rgba(255,255,255,0.8)", fontWeight: 500, marginBottom: 12, padding: 0 }}
            >
              {t("common", "backToTools")}
            </Button>
            <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
              <div style={{ width: 46, height: 46, borderRadius: 14, background: "rgba(255,255,255,0.12)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <SafetyCertificateOutlined style={{ fontSize: 24, color: "#fff" }} />
              </div>
              <div>
                <Title level={3} style={{ color: "#fff", margin: 0, fontWeight: 700 }}>
                  {t("checkSubmission", "title")}
                </Title>
                <Text style={{ color: "rgba(255,255,255,0.6)", fontSize: 13 }}>
                  {t("checkSubmission", "subtitle")}
                </Text>
              </div>
            </div>
          </div>
        </div>

        <div style={{ maxWidth: 1000, margin: "0 auto", padding: "32px 48px 48px" }}>
          <Card style={{ borderRadius: 16, border: "none", boxShadow: "0 4px 24px rgba(108,44,145,0.08)", marginTop: -32 }}>
            <ConfigProvider theme={{ token: { colorPrimary: "#9369b8" } }}>
              <Steps
                current={step}
                items={[
                  { title: t("checkSubmission", "stepRegulation") },
                  { title: t("checkSubmission", "stepInput") },
                  { title: t("checkSubmission", "stepExpected") },
                  { title: t("checkSubmission", "stepTaxonomy") },
                  { title: t("checkSubmission", "stepSummary") },
                  { title: t("checkSubmission", "stepResult") },
                ]}
                size="small"
                style={{ marginBottom: 32 }}
              />
            </ConfigProvider>

            {/* ────────────── Step 0: Choose regulation ────────────── */}
            {step === 0 && (
              <>
                <Title level={4}>
                  <BankOutlined style={{ marginRight: 8, color: "#ba68c8" }} />
                  {t("checkSubmission", "regulationTitle")}
                </Title>
                <Paragraph type="secondary">
                  {t("checkSubmission", "regulationDesc")}
                </Paragraph>

                <Radio.Group
                  value={regulation}
                  onChange={(e) => {
                    setRegulation(e.target.value);
                    if (e.target.value !== "OTHER") setCustomRegulation("");
                  }}
                  style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 16 }}
                >
                  {REGULATIONS.map((r) => (
                    <Radio.Button
                      key={r}
                      value={r}
                      style={{
                        borderRadius: 8,
                        height: 42,
                        lineHeight: "42px",
                        fontWeight: 600,
                        minWidth: 100,
                        textAlign: "center",
                      }}
                    >
                      {r}
                    </Radio.Button>
                  ))}
                  <Radio.Button
                    value="OTHER"
                    style={{
                      borderRadius: 8,
                      height: 42,
                      lineHeight: "42px",
                      fontWeight: 600,
                      minWidth: 100,
                      textAlign: "center",
                    }}
                  >
                    {t("checkSubmission", "otherRegulation")}
                  </Radio.Button>
                </Radio.Group>

                {regulation === "OTHER" && (
                  <Input
                    placeholder={t("checkSubmission", "regulationPlaceholder")}
                    value={customRegulation}
                    onChange={(e) => setCustomRegulation(e.target.value)}
                    style={{ maxWidth: 300, marginBottom: 16, borderRadius: 8 }}
                  />
                )}

                {regulationName && (
                  <Alert
                    type="success"
                    showIcon
                    icon={<CheckCircleOutlined />}
                    message={`${t("checkSubmission", "regulationSelected")}: ${regulationName}`}
                    style={{ marginBottom: 16, borderRadius: 10 }}
                  />
                )}

                <div style={{ display: "flex", justifyContent: "flex-end" }}>
                  <Button
                    type="primary"
                    onClick={() => setStep(1)}
                    disabled={nextDisabled[0]}
                    style={!nextDisabled[0] ? gradientBtn : disabledBtn}
                  >
                    {t("common", "next")}
                  </Button>
                </div>
              </>
            )}

            {/* ────────────── Step 1: Input directory ────────────── */}
            {step === 1 && (
              <>
                <Title level={4}>
                  <FolderOpenOutlined style={{ marginRight: 8, color: "#ba68c8" }} />
                  {t("checkSubmission", "inputTitle")}
                </Title>
                <Paragraph type="secondary">
                  {t("checkSubmission", "inputDesc")}
                </Paragraph>

                <Upload.Dragger
                  directory
                  multiple
                  fileList={inputFiles}
                  beforeUpload={() => false}
                  onChange={({ fileList }) => setInputFiles(fileList)}
                  style={{ marginBottom: 24, borderRadius: 14, border: "2px dashed #ba68c8", background: "#ffffff" }}
                >
                  <p className="ant-upload-drag-icon">
                    <UploadOutlined style={{ color: "#ba68c8", fontSize: 36 }} />
                  </p>
                  <p className="ant-upload-text" style={{ color: "#ba68c8", fontWeight: 500 }}>
                    {t("checkSubmission", "dragInput")}
                  </p>
                  <p className="ant-upload-hint">
                    {t("checkSubmission", "inputHintFolder")}
                  </p>
                </Upload.Dragger>

                {institutions.length > 0 && (
                  <Alert
                    type="success"
                    showIcon
                    icon={<CheckCircleOutlined />}
                    message={`${institutions.length} ${t("checkSubmission", "institutionsFound")}`}
                    description={
                      <ul style={{ margin: "4px 0 0", paddingLeft: 20 }}>
                        {institutions.map((inst) => (
                          <li key={inst.name}>
                            <Text strong>{inst.name}</Text>{" "}
                            <Text type="secondary">→ {inst.modes.join(", ")}</Text>
                          </li>
                        ))}
                      </ul>
                    }
                    style={{ marginBottom: 16, borderRadius: 10 }}
                  />
                )}

                <div style={navRow}>
                  <Button onClick={() => setStep(0)} style={{ borderRadius: 10 }}>
                    {t("common", "previous")}
                  </Button>
                  <Button
                    type="primary"
                    onClick={() => setStep(2)}
                    disabled={nextDisabled[1]}
                    style={!nextDisabled[1] ? gradientBtn : disabledBtn}
                  >
                    {t("common", "next")}
                  </Button>
                </div>
              </>
            )}

            {/* ────────────── Step 2: Expected submissions ────────────── */}
            {step === 2 && (
              <>
                <Title level={4}>
                  <FileSearchOutlined style={{ marginRight: 8, color: "#ba68c8" }} />
                  {t("checkSubmission", "expectedTitle")}
                </Title>
                <Paragraph type="secondary">
                  {t("checkSubmission", "expectedDesc")}
                </Paragraph>

                <Upload.Dragger
                  directory
                  multiple
                  fileList={expectedFiles}
                  beforeUpload={() => false}
                  onChange={({ fileList }) => setExpectedFiles(fileList)}
                  style={{ marginBottom: 24, borderRadius: 14, border: "2px dashed #ba68c8", background: "#ffffff" }}
                >
                  <p className="ant-upload-drag-icon">
                    <FileSearchOutlined style={{ color: "#ba68c8", fontSize: 36 }} />
                  </p>
                  <p className="ant-upload-text" style={{ color: "#ba68c8", fontWeight: 500 }}>
                    {t("checkSubmission", "dragExpected")}
                  </p>
                  <p className="ant-upload-hint">
                    {t("checkSubmission", "expectedHint")}
                  </p>
                </Upload.Dragger>

                {expectedInfo.length > 0 && (
                  <Alert
                    type="success"
                    showIcon
                    icon={<CheckCircleOutlined />}
                    message={`${expectedInfo.length} ${t("checkSubmission", "expectedFilesFound")}`}
                    description={
                      <ul style={{ margin: "4px 0 0", paddingLeft: 20 }}>
                        {expectedInfo.map((ef) => (
                          <li key={ef.filename}>
                            <Text strong>{ef.institution}</Text>{" "}
                            <Text type="secondary">→ {ef.filename}</Text>
                            {ef.outputCode && (
                              <>
                                {" "}
                                <Tag color="blue" style={{ marginLeft: 4 }}>{ef.outputCode}</Tag>
                              </>
                            )}
                          </li>
                        ))}
                      </ul>
                    }
                    style={{ marginBottom: 16, borderRadius: 10 }}
                  />
                )}

                <div style={navRow}>
                  <Button onClick={() => setStep(1)} style={{ borderRadius: 10 }}>
                    {t("common", "previous")}
                  </Button>
                  <Button
                    type="primary"
                    onClick={() => setStep(3)}
                    disabled={nextDisabled[2]}
                    style={!nextDisabled[2] ? gradientBtn : disabledBtn}
                  >
                    {t("common", "next")}
                  </Button>
                </div>
              </>
            )}

            {/* ────────────── Step 3: Taxonomy ────────────── */}
            {step === 3 && (
              <>
                <Title level={4}>
                  <UploadOutlined style={{ marginRight: 8, color: "#ba68c8" }} />
                  {t("checkSubmission", "taxonomyTitle")}
                </Title>
                <Paragraph type="secondary">
                  {t("checkSubmission", "taxonomyDesc")}
                </Paragraph>

                <Upload.Dragger
                  maxCount={1}
                  fileList={taxonomyFile}
                  beforeUpload={() => false}
                  onChange={({ fileList }) => setTaxonomyFile(fileList)}
                  accept=".tar.gz,.gz"
                  style={{ marginBottom: 24, borderRadius: 14, border: "2px dashed #ba68c8", background: "#ffffff" }}
                >
                  <p className="ant-upload-drag-icon">
                    <UploadOutlined style={{ color: "#ba68c8", fontSize: 36 }} />
                  </p>
                  <p className="ant-upload-text" style={{ color: "#ba68c8", fontWeight: 500 }}>
                    {t("checkSubmission", "dragTaxonomy")}
                  </p>
                  <p className="ant-upload-hint">
                    {t("checkSubmission", "taxonomyHint")}
                  </p>
                </Upload.Dragger>

                <div style={navRow}>
                  <Button onClick={() => setStep(2)} style={{ borderRadius: 10 }}>
                    {t("common", "previous")}
                  </Button>
                  <Button
                    type="primary"
                    onClick={() => setStep(4)}
                    disabled={nextDisabled[3]}
                    style={!nextDisabled[3] ? gradientBtn : disabledBtn}
                  >
                    {t("common", "next")}
                  </Button>
                </div>
              </>
            )}

            {/* ────────────── Step 4: Summary + Launch ────────────── */}
            {step === 4 && (
              <>
                <Title level={4}>
                  <ExperimentOutlined style={{ marginRight: 8, color: "#ba68c8" }} />
                  {t("checkSubmission", "summaryLaunchTitle")}
                </Title>

                <Card
                  style={{
                    borderRadius: 12,
                    background: "#ffffff",
                    border: "1px solid #ba68c8",
                    marginBottom: 24,
                  }}
                >
                  <Descriptions column={1} size="small" bordered>
                    <Descriptions.Item label={t("checkSubmission", "regulationLabel")}>
                      <Tag color="green" style={{ fontWeight: 700, fontSize: 14 }}>{regulationName}</Tag>
                    </Descriptions.Item>
                    <Descriptions.Item label={t("checkSubmission", "inputDirLabel")}>
                      <Space direction="vertical" size={2}>
                        <Text>{inputFiles.length} {t("checkSubmission", "filesUploaded")}</Text>
                        {institutions.length > 0 && (
                          <Text type="secondary">
                            {institutions.length} institution(s): {institutions.map((i) => i.name).join(", ")}
                          </Text>
                        )}
                      </Space>
                    </Descriptions.Item>
                    <Descriptions.Item label={t("checkSubmission", "expectedDirLabel")}>
                      <Space direction="vertical" size={2}>
                        <Text>{expectedInfo.length} {t("checkSubmission", "expectedFilesLabel")}</Text>
                        {expectedInfo.map((ef) => (
                          <Text key={ef.filename} type="secondary">
                            {ef.institution} → {ef.filename}
                          </Text>
                        ))}
                      </Space>
                    </Descriptions.Item>
                    <Descriptions.Item label={t("checkSubmission", "taxonomyLabel")}>
                      <Tag>{taxonomyFile[0]?.name || "-"}</Tag>
                    </Descriptions.Item>
                  </Descriptions>
                </Card>

                {/* Tests to run */}
                {institutions.length > 0 && expectedInfo.length > 0 && (
                  <Alert
                    type="info"
                    showIcon
                    icon={<ExperimentOutlined />}
                    message={t("checkSubmission", "testsToRun")}
                    description={
                      <ul style={{ margin: "4px 0 0", paddingLeft: 20 }}>
                        {institutions.flatMap((inst) =>
                          inst.modes.map((mode) => {
                            const matched = expectedInfo.find(
                              (ef) => ef.institution.toUpperCase() === inst.name.toUpperCase()
                            );
                            return (
                              <li key={`${inst.name}-${mode}`}>
                                <Text strong>{inst.name}</Text>{" "}
                                <Tag>{mode}</Tag>{" "}
                                <Text type="secondary">
                                  → {matched ? matched.filename : "?"}
                                </Text>
                              </li>
                            );
                          })
                        )}
                      </ul>
                    }
                    style={{ marginBottom: 24, borderRadius: 10 }}
                  />
                )}

                <div style={navRow}>
                  <Button onClick={() => setStep(3)} style={{ borderRadius: 10 }}>
                    {t("common", "previous")}
                  </Button>
                  <Button
                    type="primary"
                    size="large"
                    icon={<SafetyCertificateOutlined />}
                    onClick={handleValidate}
                    loading={validating}
                    disabled={!canLaunch}
                    style={{
                      height: 46,
                      borderRadius: 12,
                      background: canLaunch ? "linear-gradient(135deg, #e91e90, #ff9800)" : undefined,
                      border: "none",
                      fontWeight: 600,
                      boxShadow: canLaunch ? "0 4px 16px rgba(108,44,145,0.25)" : undefined,
                    }}
                  >
                    {validating ? t("checkSubmission", "validating") : t("checkSubmission", "launchValidation")}
                  </Button>
                </div>
                {validating && (
                  <Progress
                    percent={99}
                    status="active"
                    showInfo={false}
                    style={{ marginTop: 16 }}
                    strokeColor={{ from: "#d77fa3", to: "#ff9800" }}
                  />
                )}
              </>
            )}

            {/* ────────────── Step 5: Result ────────────── */}
            {step === 5 && result && (
              <>
                {result.status === "completed" ? (
                  <>
                    <Result
                      status="success"
                      title={t("checkSubmission", "validationComplete")}
                      subTitle={`${regulationName} — ${result.output_files.length} ${t("checkSubmission", "filesGenerated")}`}
                      extra={[
                        <Button
                          key="download"
                          type="primary"
                          icon={<DownloadOutlined />}
                          size="large"
                          onClick={handleDownload}
                          style={{ background: "linear-gradient(135deg, #e91e90, #ff9800)", border: "none" }}
                        >
                          {t("checkSubmission", "downloadReport")}
                        </Button>,
                        <Button key="restart" onClick={handleReset}>
                          {t("checkSubmission", "newValidation")}
                        </Button>,
                      ]}
                    />

                    {/* Global summary stats */}
                    {result.summary && Object.keys(result.summary).length > 0 && (
                      <>
                        <Divider orientationMargin={0} style={{ borderColor: "#e8d5f5" }}>{t("checkSubmission", "globalSummary")}</Divider>
                        <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 24 }}>
                          {result.summary.DIFF !== undefined && (
                            <Card size="small" style={{ borderRadius: 10, borderLeft: "4px solid #ff4d4f", flex: "1 1 120px" }}>
                              <Space><CloseCircleOutlined style={{ color: "#ff4d4f", fontSize: 20 }} /><Text strong>DIFF</Text></Space>
                              <Title level={3} style={{ margin: "4px 0 0", color: "#ff4d4f" }}>{result.summary.DIFF}</Title>
                            </Card>
                          )}
                          {result.summary.MISSING_IN_OUTPUT !== undefined && (
                            <Card size="small" style={{ borderRadius: 10, borderLeft: "4px solid #faad14", flex: "1 1 120px" }}>
                              <Space><WarningOutlined style={{ color: "#faad14", fontSize: 20 }} /><Text strong>MISSING</Text></Space>
                              <Title level={3} style={{ margin: "4px 0 0", color: "#faad14" }}>{result.summary.MISSING_IN_OUTPUT}</Title>
                            </Card>
                          )}
                          {result.summary.CALCULATED_OR_EXTRA !== undefined && (
                            <Card size="small" style={{ borderRadius: 10, borderLeft: "4px solid #1677ff", flex: "1 1 120px" }}>
                              <Space><FileSearchOutlined style={{ color: "#1677ff", fontSize: 20 }} /><Text strong>EXTRA</Text></Space>
                              <Title level={3} style={{ margin: "4px 0 0", color: "#1677ff" }}>{result.summary.CALCULATED_OR_EXTRA}</Title>
                            </Card>
                          )}
                          {result.summary.TOTAL !== undefined && (
                            <Card size="small" style={{ borderRadius: 10, borderLeft: "4px solid #666", flex: "1 1 120px" }}>
                              <Text strong>TOTAL</Text>
                              <Title level={3} style={{ margin: "4px 0 0" }}>{result.summary.TOTAL}</Title>
                            </Card>
                          )}
                        </div>
                      </>
                    )}

                    {/* Per-case breakdown */}
                    {result.cases_summary && result.cases_summary.length > 0 && (
                      <>
                        <Divider orientationMargin={0} style={{ borderColor: "#e8d5f5" }}>{t("checkSubmission", "caseBreakdown")}</Divider>
                        <Table
                          dataSource={result.cases_summary}
                          rowKey="case_id"
                          pagination={false}
                          size="small"
                          style={{ marginBottom: 24 }}
                          columns={[
                            { title: "Case", dataIndex: "case_id", key: "case_id", render: (v: string) => <Text strong>{v}</Text> },
                            { title: "DIFF", dataIndex: "DIFF", key: "DIFF", render: (v: number) => <Tag color={v > 0 ? "error" : "default"}>{v}</Tag> },
                            { title: "MISSING", dataIndex: "MISSING_IN_OUTPUT", key: "MISSING", render: (v: number) => <Tag color={v > 0 ? "warning" : "default"}>{v}</Tag> },
                            { title: "EXTRA", dataIndex: "CALCULATED_OR_EXTRA", key: "EXTRA", render: (v: number) => <Tag color={v > 0 ? "processing" : "default"}>{v}</Tag> },
                            { title: "TOTAL", dataIndex: "TOTAL", key: "TOTAL", render: (v: number) => <Tag>{v}</Tag> },
                          ]}
                        />
                      </>
                    )}

                    {/* File list */}
                    {result.output_files.length > 0 && (
                      <>
                        <Divider orientationMargin={0} style={{ borderColor: "#e8d5f5" }}>{t("checkSubmission", "generatedFiles")}</Divider>
                        <ul>
                          {result.output_files.map((f) => (
                            <li key={f}><Text code>{f}</Text></li>
                          ))}
                        </ul>
                      </>
                    )}
                  </>
                ) : (
                  <Result
                    status="error"
                    title={t("checkSubmission", "validationFailed")}
                    subTitle={result.error_message}
                    extra={
                      <Button onClick={handleReset}>
                        {t("checkSubmission", "restart")}
                      </Button>
                    }
                  />
                )}
              </>
            )}
          </Card>
        </div>
      </Content>
    </Layout>
  );
}
