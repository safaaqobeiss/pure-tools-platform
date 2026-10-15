import { useState, useEffect } from "react";
import { useNavigate } from "react-router-dom";
import {
  Layout,
  Typography,
  Card,
  Button,
  ConfigProvider,
  Form,
  Input,
  Select,
  Steps,
  Upload,
  Checkbox,
  Table,
  Descriptions,
  Result,
  Spin,
  Space,
  Divider,
  Alert,
  Modal,
  message,
  Progress,
} from "antd";
import {
  ArrowLeftOutlined,
  FileSearchOutlined,
  FilePdfOutlined,
  EditOutlined,
  RocketOutlined,
  DownloadOutlined,
  CloudUploadOutlined,
} from "@ant-design/icons";
import type { UploadFile } from "antd";
import api from "../api/axios";
import Navbar from "../components/Navbar";
import { useLanguage } from "../context/LanguageContext";

const { Content } = Layout;
const { Title, Text, Paragraph } = Typography;

interface JiraParsedData {
  regulation_code: string;
  project_type: string;
  git_project_name: string;
  version: string;
  validity_date: string;
  authority_name: string;
  ticket_key: string;
  title: string;
  regulation_project_name: string;
  output_project_name: string;
  git_project_url: string;
  output_code: string;
  output_name: string;
  output_description: string;
  output_periodicity: string;
  output_type: string;
  common_collection: string;
  submission_format: string;
  container_type: string;
  regulation_identification: Record<string, string>;
  dictionary_identification: Record<string, string>;
  transformation_rules_identification: Record<string, string>;
  transformed_data_identification: Record<string, string>;
  ui_dashboard_identification: Record<string, string>;
  output_identification: Record<string, string>;
  output_structure_identification: Record<string, string>;
  file_mappings: Record<string, string>;
}

interface ReferenceFile {
  name: string;
  extension: string;
}

interface GenerateResult {
  execution_id: string;
  status: string;
  output_files: string[];
  error_message?: string;
}

// ── Mode Selection Screen ────────────────────────────────

function ModeSelection({
  onSelectMode,
}: {
  onSelectMode: (mode: "jira_pdf" | "manual") => void;
}) {
  const { t } = useLanguage();
  return (
    <div style={{ maxWidth: 600, margin: "40px auto" }}>
      <Title level={3} style={{ textAlign: "center", marginBottom: 32, color: "#a9516a" }}>
        {t("initProject", "modeTitle")}
      </Title>

      <Card
        hoverable
        style={{
          marginBottom: 16,
          borderRadius: 14,
          border: "none",
          boxShadow: "0 4px 20px rgba(183,28,28,0.1)",
          cursor: "pointer",
          transition: "transform 0.2s, box-shadow 0.2s",
        }}
        onClick={() => onSelectMode("jira_pdf")}
      >
        <Space>
          <div style={{ width: 48, height: 48, borderRadius: 12, background: "linear-gradient(135deg, #a9516a, #d77fa3)", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <FilePdfOutlined style={{ fontSize: 24, color: "#fff" }} />
          </div>
          <div>
            <Title level={5} style={{ margin: 0, color: "#a9516a" }}>
              {t("initProject", "modeJira")}
            </Title>
            <Text type="secondary">
              {t("initProject", "modeJiraDesc")}
            </Text>
          </div>
        </Space>
      </Card>

      <Card
        hoverable
        style={{
          borderRadius: 14,
          border: "none",
          boxShadow: "0 4px 20px rgba(183,28,28,0.1)",
          cursor: "pointer",
          transition: "transform 0.2s, box-shadow 0.2s",
        }}
        onClick={() => onSelectMode("manual")}
      >
        <Space>
          <div style={{ width: 48, height: 48, borderRadius: 12, background: "linear-gradient(135deg, #ff9800, #ffca28)", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <EditOutlined style={{ fontSize: 24, color: "#fff" }} />
          </div>
          <div>
            <Title level={5} style={{ margin: 0, color: "#ff9800" }}>
              {t("initProject", "modeManual")}
            </Title>
            <Text type="secondary">
              {t("initProject", "modeManualDesc")}
            </Text>
          </div>
        </Space>
      </Card>
    </div>
  );
}

// ── Jira PDF Wizard ──────────────────────────────────────

function JiraPdfWizard() {
  const { t } = useLanguage();
  const [step, setStep] = useState(0);
  const [fileList, setFileList] = useState<UploadFile[]>([]);
  const [parsing, setParsing] = useState(false);
  const [parsedData, setParsedData] = useState<JiraParsedData | null>(null);
  const [generating, setGenerating] = useState(false);
  const [result, setResult] = useState<GenerateResult | null>(null);
  const [gitModalVisible, setGitModalVisible] = useState(false);
  const [gitForm] = Form.useForm();
  const [pushingGit, setPushingGit] = useState(false);

  // Load stored token on component mount or when modal opens
  useEffect(() => {
    if (gitModalVisible) {
      const storedToken = sessionStorage.getItem("gitlab_token_jira");
      if (storedToken) {
        try {
          gitForm.setFieldValue("token", atob(storedToken));
        } catch (e) {
          // Ignore decoding errors
        }
      }
    }
  }, [gitModalVisible, gitForm]);

  const handleParsePdf = async () => {
    if (fileList.length === 0 || !fileList[0].originFileObj) {
      message.warning(t("initProject", "selectPdf"));
      return;
    }

    setParsing(true);
    try {
      const formData = new FormData();
      formData.append("file", fileList[0].originFileObj);

      const res = await api.post("/api/init-project/parse-jira", formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });

      setParsedData(res.data);
      setStep(1);
      message.success(t("initProject", "jiraParsed"));
    } catch (err: any) {
      message.error(
        err.response?.data?.detail || t("initProject", "parseError")
      );
    } finally {
      setParsing(false);
    }
  };

  const handleGenerate = async () => {
    if (!fileList[0]?.originFileObj) return;

    setGenerating(true);
    try {
      const formData = new FormData();
      formData.append("file", fileList[0].originFileObj);
      formData.append("parameters", JSON.stringify({}));

      const res = await api.post(
        "/api/init-project/generate-from-jira",
        formData,
        { headers: { "Content-Type": "multipart/form-data" } }
      );

      setResult(res.data);
      setStep(2);

      if (res.data.status === "completed") {
        message.success(t("initProject", "filesGenerated"));
      } else {
        message.error(t("initProject", "generationFailed"));
      }
    } catch (err: any) {
      message.error(
        err.response?.data?.detail || t("initProject", "generateError")
      );
    } finally {
      setGenerating(false);
    }
  };

  const handleGitPush = async () => {
    try {
      await gitForm.validateFields();
    } catch {
      return;
    }

    if (!result?.execution_id) return;

    const values = gitForm.getFieldsValue();
    setPushingGit(true);
    try {
      const res = await api.post("/api/init-project/push-gitlab", {
        execution_id: result.execution_id,
        branch: values.branch,
        token: values.token,
        regulation_project_name: values.regulation_project_name || "",
        output_project_name: values.output_project_name || "",
      });

      if (res.data.success) {
        // Store token securely in sessionStorage for current session only
        sessionStorage.setItem("gitlab_token_jira", btoa(values.token));
        message.success(t("initProject", "pushedSuccess"));
        setGitModalVisible(false);
      } else {
        message.error(`${t("initProject", "gitlabError")}: ${res.data.message}`);
      }
    } catch (err: any) {
      message.error(
        err.response?.data?.detail || t("initProject", "gitlabError")
      );
    } finally {
      setPushingGit(false);
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
      message.error("Erreur lors du telechargement");
    }
  };

  const gitProjectName = parsedData?.git_project_name || parsedData?.regulation_code?.toLowerCase() || "";
  const regProjectNameDefault = parsedData?.regulation_project_name || `pure-regulation-${gitProjectName}`;
  const outProjectNameDefault = parsedData?.output_project_name || `pure-regulatoryoutput-${gitProjectName}`;

  // Detect which project types were actually generated from output files
  const hasRegulationFiles = result?.output_files?.some((f) => f.startsWith("regulation")) ?? false;
  const hasOutputFiles = result?.output_files?.some((f) => f.startsWith("output")) ?? false;

  return (
    <div style={{ maxWidth: 800, margin: "0 auto" }}>
      <ConfigProvider theme={{ token: { colorPrimary: "#a9516a" } }}>
        <Steps
          current={step}
          items={[
            { title: t("initProject", "pdfFile") },
            { title: t("initProject", "preview") },
            { title: t("initProject", "result") },
          ]}
          style={{ marginBottom: 32 }}
        />
      </ConfigProvider>

      {/* Step 0: Upload PDF */}
      {step === 0 && (
        <Card>
          <Title level={4}>
            <FilePdfOutlined /> {t("initProject", "selectPdfFile")}
          </Title>
          <Paragraph type="secondary">
            {t("initProject", "pdfExportDesc")}
          </Paragraph>

          <Upload.Dragger
            maxCount={1}
            fileList={fileList}
            beforeUpload={() => false}
            onChange={({ fileList: newList }) => setFileList(newList)}
            accept=".pdf"
            style={{ marginBottom: 24, borderRadius: 14, border: "2px dashed #9369b8", background: "#ffffff" }}
          >
            <p className="ant-upload-drag-icon">
              <FilePdfOutlined />
            </p>
            <p className="ant-upload-text">
              {t("initProject", "dragPdf")}
            </p>
          </Upload.Dragger>

          <Button
            type="primary"
            size="large"
            icon={<FileSearchOutlined />}
            onClick={handleParsePdf}
            loading={parsing}
            disabled={fileList.length === 0}
            block
            style={{
              height: 46,
              borderRadius: 12,
              background: fileList.length > 0 ? "linear-gradient(135deg, #d77fa3, #ff9800)" : undefined,
              border: "none",
              fontWeight: 600,
              boxShadow: fileList.length > 0 ? "0 4px 16px rgba(229,57,53,0.25)" : undefined,
            }}
          >
            {parsing ? t("initProject", "analyzing") : t("initProject", "analyzePdf")}
          </Button>
        </Card>
      )}

      {/* Step 1: Review parsed data */}
      {step === 1 && parsedData && (
        <Card>
          <Title level={4}>{t("initProject", "extractedData")}</Title>

          {parsedData.ticket_key && (
            <Alert
              message={`Ticket: ${parsedData.ticket_key}`}
              description={parsedData.title}
              type="info"
              showIcon
              style={{ marginBottom: 16 }}
            />
          )}

          <Descriptions bordered column={1} size="small">
            <Descriptions.Item label={t("initProject", "regulationCode")}>
              <strong>{parsedData.regulation_code}</strong>
            </Descriptions.Item>
            <Descriptions.Item label={t("initProject", "projectType")}>
              {parsedData.project_type}
            </Descriptions.Item>
            <Descriptions.Item label={t("initProject", "gitProjectName")}>
              {parsedData.git_project_name}
            </Descriptions.Item>
            <Descriptions.Item label={t("initProject", "version")}>
              {parsedData.version}
            </Descriptions.Item>
            <Descriptions.Item label={t("initProject", "validityDate")}>
              {parsedData.validity_date}
            </Descriptions.Item>
            <Descriptions.Item label={t("initProject", "authority")}>
              {parsedData.authority_name}
            </Descriptions.Item>
            {parsedData.git_project_url && (
              <Descriptions.Item label={t("initProject", "gitProjectUrl")}>
                {parsedData.git_project_url}
              </Descriptions.Item>
            )}
            {parsedData.project_type !== "output" && parsedData.regulation_project_name && (
              <Descriptions.Item label={t("initProject", "regProjectName")}>
                {parsedData.regulation_project_name}
              </Descriptions.Item>
            )}
            {parsedData.project_type !== "regulation" && parsedData.output_project_name && (
              <Descriptions.Item label={t("initProject", "outProjectName")}>
                {parsedData.output_project_name}
              </Descriptions.Item>
            )}
          </Descriptions>

          {parsedData.project_type !== "regulation" && (parsedData.output_code || parsedData.output_name) && (
            <>
              <Divider orientation="start">{t("initProject", "outputInfo")}</Divider>
              <Descriptions bordered column={1} size="small">
                {parsedData.output_code && (
                  <Descriptions.Item label={t("initProject", "outputCode")}>
                    {parsedData.output_code}
                  </Descriptions.Item>
                )}
                {parsedData.output_name && (
                  <Descriptions.Item label={t("initProject", "outputName")}>
                    {parsedData.output_name}
                  </Descriptions.Item>
                )}
                {parsedData.output_periodicity && (
                  <Descriptions.Item label={t("initProject", "outputPeriodicity")}>
                    {parsedData.output_periodicity}
                  </Descriptions.Item>
                )}
                {parsedData.output_type && (
                  <Descriptions.Item label={t("initProject", "outputType")}>
                    {parsedData.output_type}
                  </Descriptions.Item>
                )}
                {parsedData.submission_format && (
                  <Descriptions.Item label={t("initProject", "submissionFormat")}>
                    {parsedData.submission_format}
                  </Descriptions.Item>
                )}
                {parsedData.container_type && (
                  <Descriptions.Item label={t("initProject", "containerType")}>
                    {parsedData.container_type}
                  </Descriptions.Item>
                )}
              </Descriptions>
            </>
          )}

          {parsedData.project_type !== "output" && Object.keys(parsedData.regulation_identification).length > 0 && (
            <>
              <Divider orientation="start">{t("initProject", "identificationReg")}</Divider>
              <Descriptions bordered column={1} size="small">
                {Object.entries(parsedData.regulation_identification).map(
                  ([key, value]) => (
                    <Descriptions.Item key={key} label={key}>
                      {value}
                    </Descriptions.Item>
                  )
                )}
              </Descriptions>
            </>
          )}

          {parsedData.project_type !== "output" && Object.keys(parsedData.dictionary_identification).length > 0 && (
            <>
              <Divider orientation="start">{t("initProject", "identificationDict")}</Divider>
              <Descriptions bordered column={1} size="small">
                {Object.entries(parsedData.dictionary_identification).map(
                  ([key, value]) => (
                    <Descriptions.Item key={key} label={key}>
                      {value}
                    </Descriptions.Item>
                  )
                )}
              </Descriptions>
            </>
          )}

          {Object.keys(parsedData.output_identification).length > 0 && (
            <>
              <Divider orientation="start">{t("initProject", "identificationOutput")}</Divider>
              <Descriptions bordered column={1} size="small">
                {Object.entries(parsedData.output_identification).map(
                  ([key, value]) => (
                    <Descriptions.Item key={key} label={key}>
                      {value}
                    </Descriptions.Item>
                  )
                )}
              </Descriptions>
            </>
          )}

          {Object.keys(parsedData.file_mappings).length > 0 && (
            <>
              <Divider orientation="start">{t("initProject", "mappedFiles")}</Divider>
              <Table
                dataSource={Object.entries(parsedData.file_mappings).map(
                  ([src, dst]) => ({
                    key: src,
                    source: src,
                    target: dst,
                  })
                )}
                columns={[
                  { title: "Source", dataIndex: "source" },
                  { title: t("initProject", "target"), dataIndex: "target" },
                ]}
                pagination={false}
                size="small"
              />
            </>
          )}

          <Divider />

          <Space>
            <Button onClick={() => setStep(0)}>{t("initProject", "back")}</Button>
            <Button
              type="primary"
              size="large"
              icon={<RocketOutlined />}
              onClick={handleGenerate}
              loading={generating}
              style={{
                borderRadius: 12,
                background: "linear-gradient(135deg, #d77fa3, #ff9800)",
                border: "none",
                fontWeight: 600,
                boxShadow: "0 4px 16px rgba(233,30,144,0.25)",
              }}
            >
              {generating ? t("initProject", "generatingFiles") : t("initProject", "generateFiles")}
            </Button>
          </Space>
          {generating && (
            <Progress
              percent={99}
              status="active"
              showInfo={false}
              style={{ marginTop: 16 }}
              strokeColor={{ from: "#d77fa3", to: "#ff9800" }}
            />
          )}
        </Card>
      )}

      {/* Step 2: Result */}
      {step === 2 && result && (
        <Card>
          {result.status === "completed" ? (
            <>
              <Result
                status="success"
                title={t("initProject", "generationSuccess")}
                subTitle={`${result.output_files.length} fichier(s)`}
                extra={[
                  <Button
                    key="download"
                    type="primary"
                    icon={<DownloadOutlined />}
                    size="large"
                    onClick={handleDownload}
                  >
                    {t("initProject", "downloadProject")}
                  </Button>,
                  <Button
                    key="gitlab"
                    icon={<CloudUploadOutlined />}
                    size="large"
                    onClick={() => setGitModalVisible(true)}
                  >
                    {t("initProject", "pushGitlab")}
                  </Button>,
                ]}
              />

              {result.output_files.length > 0 && (
                <>
                  <Divider orientation="start">{t("initProject", "generatedFiles")}</Divider>
                  <ul>
                    {result.output_files.map((f) => (
                      <li key={f}>
                        <Text code>{f}</Text>
                      </li>
                    ))}
                  </ul>
                </>
              )}
            </>
          ) : (
            <Result
              status="error"
              title={t("initProject", "failedTitle")}
              subTitle={result.error_message}
              extra={
                <Button onClick={() => setStep(0)}>{t("initProject", "restart")}</Button>
              }
            />
          )}
        </Card>
      )}

      {/* GitLab Push Modal */}
      <Modal
        title={t("initProject", "pushGitlab")}
        open={gitModalVisible}
        onCancel={() => setGitModalVisible(false)}
        onOk={handleGitPush}
        confirmLoading={pushingGit}
        okText={t("initProject", "push")}
        cancelText={t("initProject", "cancel")}  
      >
        <Form form={gitForm} layout="vertical">
          <Form.Item
            name="branch"
            label={t("initProject", "targetBranch")}
            initialValue="main"
            rules={[{ required: true }]}
          >
            <Input />
          </Form.Item>
          <Form.Item
            name="token"
            label="GitLab Token (PRIVATE-TOKEN)"
            rules={[
              { required: true, message: t("initProject", "tokenRequired") },
            ]}
          >
            <Input.Password />
          </Form.Item>
          <div style={{ marginBottom: 16 }}>
            <Button
              type="link"
              danger
              size="small"
              onClick={() => {
                sessionStorage.removeItem("gitlab_token_jira");
                gitForm.setFieldValue("token", "");
                message.info(t("initProject", "tokenCleared"));
              }}
            >
              {t("initProject", "clearToken")}
            </Button>
          </div>
          {hasRegulationFiles && (
            <>
              <Form.Item
                name="regulation_project_name"
                label={t("initProject", "regProjectName")}
                initialValue={regProjectNameDefault}
              >
                <Input />
              </Form.Item>
              <Alert
                message={`📍 ${t("initProject", "projectPath")}: ${regProjectNameDefault}`}
                type="info"
                showIcon={false}
                style={{ marginBottom: 16 }}
              />
            </>
          )}
          {hasOutputFiles && (
            <>
              <Form.Item
                name="output_project_name"
                label={t("initProject", "outProjectName")}
                initialValue={outProjectNameDefault}
              >
                <Input />
              </Form.Item>
              <Alert
                message={`📍 ${t("initProject", "projectPath")}: ${outProjectNameDefault}`}
                type="info"
                showIcon={false}
                style={{ marginBottom: 16 }}
              />
            </>
          )}
        </Form>
      </Modal>
    </div>
  );
}

// ── Manual Wizard ────────────────────────────────────────

function ManualWizard() {
  const { t } = useLanguage();
  const [step, setStep] = useState(0);
  const [form] = Form.useForm();
  const [, setRefFiles] = useState<{
    regulation_files: ReferenceFile[];
    output_files: ReferenceFile[];
  }>({ regulation_files: [], output_files: [] });
  const [loadingFiles, setLoadingFiles] = useState(true);

  // File selection state
  const [regFileMap, setRegFileMap] = useState<
    Record<string, { selected: boolean; targetName: string }>
  >({});
  const [outFileMap, setOutFileMap] = useState<
    Record<string, { selected: boolean; targetName: string }>
  >({});

  const [generating, setGenerating] = useState(false);
  const [result, setResult] = useState<GenerateResult | null>(null);
  const [gitModalVisible, setGitModalVisible] = useState(false);
  const [gitForm] = Form.useForm();
  const [pushingGit, setPushingGit] = useState(false);

  // Track project type with state (Form.useWatch loses value when Form.Item unmounts between steps)
  const [projectType, setProjectType] = useState("both");

  // Load reference files on mount
  useEffect(() => {
    api
      .get("/api/init-project/reference-files")
      .then((res) => {
        setRefFiles(res.data);
        // Initialize file maps
        const regMap: typeof regFileMap = {};
        res.data.regulation_files.forEach((f: ReferenceFile) => {
          regMap[f.name] = { selected: true, targetName: f.name };
        });
        setRegFileMap(regMap);

        const outMap: typeof outFileMap = {};
        res.data.output_files.forEach((f: ReferenceFile) => {
          outMap[f.name] = { selected: true, targetName: f.name };
        });
        setOutFileMap(outMap);
      })
      .catch(() => message.error(t("initProject", "loadRefError")))
      .finally(() => setLoadingFiles(false));
  }, []);

  // Load stored token on component mount or when modal opens
  useEffect(() => {
    if (gitModalVisible) {
      const storedToken = sessionStorage.getItem("gitlab_token_manual");
      if (storedToken) {
        try {
          gitForm.setFieldValue("token", atob(storedToken));
        } catch (e) {
          // Ignore decoding errors
        }
      }
    }
  }, [gitModalVisible, gitForm]);

  // Update target names when regulation code changes
  const updateFileNames = (code: string) => {
    const lowerCode = code.toLowerCase().replace(/[^a-z0-9_-]/g, "-").replace(/^[-_]+|[-_]+$/g, "");
    const tokenize = (name: string) => {
      let result = name;
      const replacements: Record<string, string> = {
        regulatorCode: lowerCode || "regulation",
        regulationCode: lowerCode || "regulation",
        regulatoryDomainCode: lowerCode || "regulation",
        regulatoryOutputCode: "REGULATORY_OUTPUT_CODE",
        outputStructureCode: "OUTPUT_STRUCTURE_CODE",
        formCode: "FORM_CODE",
        "(shortName)": "SHORT_NAME",
      };
      for (const [old, newVal] of Object.entries(replacements)) {
        result = result.replaceAll(old, newVal);
      }
      return result;
    };

    setRegFileMap((prev) => {
      const updated = { ...prev };
      for (const key of Object.keys(updated)) {
        updated[key] = { ...updated[key], targetName: tokenize(key) };
      }
      return updated;
    });

    setOutFileMap((prev) => {
      const updated = { ...prev };
      for (const key of Object.keys(updated)) {
        updated[key] = { ...updated[key], targetName: tokenize(key) };
      }
      return updated;
    });
  };

  const getVisibleSteps = () => {
    const steps = [
      { title: "Code", key: "code" },
      { title: "Type", key: "type" },
      { title: t("initProject", "gitNameTitle"), key: "git_name" },
    ];
    if (projectType !== "output") {
      steps.push({ title: t("initProject", "regFilesTitle"), key: "reg_files" });
    }
    if (projectType !== "regulation") {
      steps.push({ title: t("initProject", "outFilesTitle"), key: "out_files" });
    }
    steps.push({ title: "GitLab", key: "gitlab" });
    steps.push({ title: t("initProject", "summaryStep"), key: "summary" });
    return steps;
  };

  const visibleSteps = getVisibleSteps();

  const validateAndNext = async () => {
    const currentKey = visibleSteps[step]?.key;
    if (currentKey === "code") {
      const code = form.getFieldValue("regulatory_code");
      if (!code?.trim()) {
        message.warning(t("initProject", "codeWarning"));
        return;
      }
      updateFileNames(code.trim());
    }
    if (currentKey === "git_name") {
      const name = form.getFieldValue("git_project_name");
      const code = form.getFieldValue("regulatory_code");
      if (!name?.trim() && !code?.trim()) {
        message.warning(t("initProject", "gitNameWarning"));
        return;
      }
    }
    setStep((s) => Math.min(s + 1, visibleSteps.length - 1));
  };

  const handleGenerate = async () => {
    setGenerating(true);
    try {
      const values = form.getFieldsValue();

      // Build file maps from selections
      const regulationFileMap: Record<string, string> = {};
      for (const [src, data] of Object.entries(regFileMap)) {
        if (data.selected) {
          regulationFileMap[src] = data.targetName;
        }
      }

      const outputFileMap: Record<string, string> = {};
      for (const [src, data] of Object.entries(outFileMap)) {
        if (data.selected) {
          outputFileMap[src] = data.targetName;
        }
      }

      const payload = {
        regulatory_code: values.regulatory_code?.trim()?.toUpperCase() || "",
        project_type: projectType,
        git_project_name:
          values.git_project_name?.trim() ||
          values.regulatory_code?.trim()?.toLowerCase() ||
          "",
        version: values.version || "2026.1.0",
        regulation_file_map: regulationFileMap,
        output_file_map: outputFileMap,
        regulation_identification: {},
        dictionary_identification: {},
        transformation_rules_identification: {},
        transformed_data_identification: {},
        ui_dashboard_identification: {},
        output_identification: {},
        output_structure_identification: {},
      };

      const res = await api.post("/api/init-project/generate", payload);
      setResult(res.data);

      if (res.data.status === "completed") {
        setStep(visibleSteps.length - 1); // Go to summary/result
        message.success(t("initProject", "filesGenerated"));
      } else {
        message.error(t("initProject", "generationFailed"));
      }
    } catch (err: any) {
      message.error(
        err.response?.data?.detail || t("initProject", "generateError")
      );
    } finally {
      setGenerating(false);
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
      message.error(t("initProject", "downloadError"));
    }
  };

  const handleGitPush = async () => {
    try {
      await gitForm.validateFields();
    } catch {
      return;
    }
    if (!result?.execution_id) return;

    const values = gitForm.getFieldsValue();
    setPushingGit(true);
    try {
      const res = await api.post("/api/init-project/push-gitlab", {
        execution_id: result.execution_id,
        branch: values.branch,
        token: values.token,
        regulation_project_name: values.regulation_project_name || "",
        output_project_name: values.output_project_name || "",
      });
      if (res.data.success) {
        // Store token securely in sessionStorage for current session only
        sessionStorage.setItem("gitlab_token_manual", btoa(values.token));
        message.success(t("initProject", "pushedManualSuccess"));
        setGitModalVisible(false);
      } else {
        message.error(`${t("initProject", "gitlabError")}: ${res.data.message}`);
      }
    } catch (err: any) {
      message.error(err.response?.data?.detail || t("initProject", "gitlabError"));
    } finally {
      setPushingGit(false);
    }
  };

  if (loadingFiles) {
    return <Spin size="large" style={{ display: "block", margin: "80px auto" }} />;
  }

  const renderFileTable = (
    files: Record<string, { selected: boolean; targetName: string }>,
    setFiles: React.Dispatch<
      React.SetStateAction<
        Record<string, { selected: boolean; targetName: string }>
      >
    >,
    title: string
  ) => (
    <Card title={title}>
      <Paragraph type="secondary">
        {t("initProject", "fileSelectDesc")}
      </Paragraph>
      <Table
        dataSource={Object.entries(files).map(([name, data]) => ({
          key: name,
          name,
          ...data,
        }))}
        columns={[
          {
            title: "",
            dataIndex: "selected",
            width: 50,
            render: (val: boolean, record: { key: string }) => (
              <Checkbox
                checked={val}
                onChange={(e) =>
                  setFiles((prev) => ({
                    ...prev,
                    [record.key]: {
                      ...prev[record.key],
                      selected: e.target.checked,
                    },
                  }))
                }
              />
            ),
          },
          {
            title: t("initProject", "sourceFile"),
            dataIndex: "name",
            render: (text: string) => <Text code>{text}</Text>,
          },
          { title: "", width: 40, render: () => <span>&#8594;</span> },
          {
            title: t("initProject", "targetName"),
            dataIndex: "targetName",
            render: (val: string, record: { key: string }) => (
              <Input
                value={val}
                size="small"
                onChange={(e) =>
                  setFiles((prev) => ({
                    ...prev,
                    [record.key]: {
                      ...prev[record.key],
                      targetName: e.target.value,
                    },
                  }))
                }
              />
            ),
          },
        ]}
        pagination={false}
        size="small"
      />
    </Card>
  );

  const currentKey = visibleSteps[step]?.key;
  const gitProjectName =
    form.getFieldValue("git_project_name") ||
    form.getFieldValue("regulatory_code")?.toLowerCase() ||
    "";

  // Detect which project types were actually generated from output files
  const hasRegulationFiles = result?.output_files?.some((f) => f.startsWith("regulation")) ?? false;
  const hasOutputFiles = result?.output_files?.some((f) => f.startsWith("output")) ?? false;

  return (
    <div style={{ maxWidth: 900, margin: "0 auto" }}>
      <ConfigProvider theme={{ token: { colorPrimary: "#a9516a" } }}>
        <Steps
          current={step}
          items={visibleSteps.map((s) => ({ title: s.title }))}
          style={{ marginBottom: 32 }}
          size="small"
        />
      </ConfigProvider>

      {/* Step: Code */}
      {currentKey === "code" && (
        <Card>
          <Title level={4}>{t("initProject", "codeTitle")}</Title>
          <Paragraph type="secondary">
            {t("initProject", "codeDesc")}
          </Paragraph>
          <Form form={form} layout="vertical">
            <Form.Item
              name="regulatory_code"
              label={t("initProject", "regulationCode")}
              rules={[{ required: true, message: t("initProject", "codeRequired") }]}
            >
              <Input placeholder="COREP" style={{ maxWidth: 400 }} />
            </Form.Item>
          </Form>
        </Card>
      )}

      {/* Step: Type */}
      {currentKey === "type" && (
        <Card>
          <Title level={4}>{t("initProject", "typeTitle")}</Title>
          <Paragraph type="secondary">
            {t("initProject", "typeDesc")}
          </Paragraph>
          <Form form={form} layout="vertical">
            <Form.Item
              name="project_type"
              label="Type"
              initialValue="both"
            >
              <Select
                style={{ maxWidth: 400 }}
                onChange={(val) => setProjectType(val)}
              >
                <Select.Option value="both">
                  {t("initProject", "typeBoth")}
                </Select.Option>
                <Select.Option value="regulation">
                  {t("initProject", "typeRegOnly")}
                </Select.Option>
                <Select.Option value="output">
                  {t("initProject", "typeOutOnly")}
                </Select.Option>
              </Select>
            </Form.Item>
          </Form>
        </Card>
      )}

      {/* Step: Git Name */}
      {currentKey === "git_name" && (
        <Card>
          <Title level={4}>{t("initProject", "gitNameTitle")}</Title>
          <Paragraph type="secondary">
            {t("initProject", "gitNameDesc")}
          </Paragraph>
          <Form form={form} layout="vertical">
            <Form.Item
              name="git_project_name"
              label={t("initProject", "gitProjectName")}
            >
              <Input
                placeholder={
                  form.getFieldValue("regulatory_code")?.toLowerCase() || ""
                }
                style={{ maxWidth: 400 }}
              />
            </Form.Item>
            <Form.Item
              name="version"
              label={t("initProject", "version")}
              initialValue="2026.1.0"
            >
              <Input style={{ maxWidth: 400 }} />
            </Form.Item>
          </Form>
        </Card>
      )}

      {/* Step: Regulation Files */}
      {currentKey === "reg_files" &&
        renderFileTable(regFileMap, setRegFileMap, t("initProject", "regFilesTitle"))}

      {/* Step: Output Files */}
      {currentKey === "out_files" &&
        renderFileTable(outFileMap, setOutFileMap, t("initProject", "outFilesTitle"))}

      {/* Step: GitLab */}
      {currentKey === "gitlab" && (
        <Card>
          <Title level={4}>{t("initProject", "gitlabPublish")}</Title>
          <Paragraph type="secondary">
            {t("initProject", "gitlabPublishDesc")}
          </Paragraph>
          <Alert
            message={t("initProject", "gitlabPublishAlert")}
            type="info"
            showIcon
          />
        </Card>
      )}

      {/* Step: Summary / Result */}
      {currentKey === "summary" && (
        <Card>
          {result?.status === "completed" ? (
            <>
              <Result
                status="success"
                title={t("initProject", "generationSuccess")}
                subTitle={`${result.output_files.length} fichier(s)`}
                extra={[
                  <Button
                    key="download"
                    type="primary"
                    icon={<DownloadOutlined />}
                    size="large"
                    onClick={handleDownload}
                  >
                    {t("initProject", "downloadProject")}
                  </Button>,
                  <Button
                    key="gitlab"
                    icon={<CloudUploadOutlined />}
                    size="large"
                    onClick={() => setGitModalVisible(true)}
                  >
                    {t("initProject", "pushGitlab")}
                  </Button>,
                ]}
              />
              {result.output_files.length > 0 && (
                <>
                  <Divider orientation="start">{t("initProject", "generatedFiles")}</Divider>
                  <ul>
                    {result.output_files.map((f) => (
                      <li key={f}>
                        <Text code>{f}</Text>
                      </li>
                    ))}
                  </ul>
                </>
              )}
            </>
          ) : result?.status === "failed" ? (
            <Result
              status="error"
              title={t("initProject", "failedTitle")}
              subTitle={result.error_message}
              extra={<Button onClick={() => setStep(0)}>{t("initProject", "restart")}</Button>}
            />
          ) : (
            <>
              <Title level={4}>{t("initProject", "summaryTitle")}</Title>

              <Descriptions bordered column={1} size="small" style={{ marginBottom: 16 }}>
                <Descriptions.Item label={t("initProject", "regulationCode")}>
                  <strong>
                    {form.getFieldValue("regulatory_code")?.toUpperCase()}
                  </strong>
                </Descriptions.Item>
                <Descriptions.Item label={t("initProject", "projectType")}>
                  {projectType}
                </Descriptions.Item>
                <Descriptions.Item label={t("initProject", "gitProjectName")}>
                  {form.getFieldValue("git_project_name") ||
                    form.getFieldValue("regulatory_code")?.toLowerCase()}
                </Descriptions.Item>
                <Descriptions.Item label={t("initProject", "version")}>
                  {form.getFieldValue("version") || "2026.1.0"}
                </Descriptions.Item>
              </Descriptions>

              {projectType !== "output" && (
                <>
                  <Divider orientation="start">
                    {t("initProject", "regFilesTitle")} ({
                      Object.values(regFileMap).filter((f) => f.selected).length
                    })
                  </Divider>
                  <ul>
                    {Object.entries(regFileMap)
                      .filter(([, d]) => d.selected)
                      .map(([src, d]) => (
                        <li key={src}>
                          <Text type="secondary">{src}</Text> &#8594;{" "}
                          <Text strong>{d.targetName}</Text>
                        </li>
                      ))}
                  </ul>
                </>
              )}

              {projectType !== "regulation" && (
                <>
                  <Divider orientation="start">
                    {t("initProject", "outFilesTitle")} ({
                      Object.values(outFileMap).filter((f) => f.selected).length
                    })
                  </Divider>
                  <ul>
                    {Object.entries(outFileMap)
                      .filter(([, d]) => d.selected)
                      .map(([src, d]) => (
                        <li key={src}>
                          <Text type="secondary">{src}</Text> &#8594;{" "}
                          <Text strong>{d.targetName}</Text>
                        </li>
                      ))}
                  </ul>
                </>
              )}

              <Divider />

              <Button
                type="primary"
                size="large"
                icon={<RocketOutlined />}
                onClick={handleGenerate}
                loading={generating}
                block
                style={{
                  height: 46,
                  borderRadius: 12,
                  background: "linear-gradient(135deg, #e91e90, #ff9800)",
                  border: "none",
                  fontWeight: 600,
                  boxShadow: "0 4px 16px rgba(229,57,53,0.25)",
                }}
              >
                {generating
                  ? t("initProject", "generatingFiles")
                  : t("initProject", "generateFiles")}
              </Button>
              {generating && (
                <Progress
                  percent={99}
                  status="active"
                  showInfo={false}
                  style={{ marginTop: 16 }}
                  strokeColor={{ from: "#e91e90", to: "#ff9800" }}
                />
              )}
            </>
          )}
        </Card>
      )}

      {/* Navigation buttons */}
      {currentKey !== "summary" && (
        <div
          style={{
            marginTop: 24,
            display: "flex",
            justifyContent: "space-between",
          }}
        >
          <Button disabled={step === 0} onClick={() => setStep((s) => s - 1)} style={{ borderRadius: 10, height: 40 }}>
            {t("common", "previous")}
          </Button>
          <Button type="primary" onClick={validateAndNext} style={{ borderRadius: 10, height: 40, background: "linear-gradient(135deg, #e91e90, #ff9800)", border: "none", fontWeight: 600, boxShadow: "0 4px 12px rgba(233,30,144,0.2)" }}>
            {t("common", "next")}
          </Button>
        </div>
      )}

      {currentKey === "summary" && !result && (
        <div style={{ marginTop: 24 }}>
          <Button onClick={() => setStep((s) => s - 1)}>{t("common", "previous")}</Button>
        </div>
      )}

      {/* GitLab Push Modal */}
      <Modal
        title={t("initProject", "pushGitlab")}
        open={gitModalVisible}
        onCancel={() => setGitModalVisible(false)}
        onOk={handleGitPush}
        confirmLoading={pushingGit}
        okText={t("initProject", "push")}
        cancelText={t("initProject", "cancel")}
      >
        <Form form={gitForm} layout="vertical">
          <Form.Item
            name="branch"
            label={t("initProject", "targetBranch")}
            initialValue="main"
            rules={[{ required: true }]}
          >
            <Input />
          </Form.Item>
          <Form.Item
            name="token"
            label="GitLab Token (PRIVATE-TOKEN)"
            rules={[
              { required: true, message: t("initProject", "tokenRequired") },
            ]}
          >
            <Input.Password />
          </Form.Item>
          {hasRegulationFiles && (
            <>
              <Form.Item
                name="regulation_project_name"
                label={t("initProject", "regProjectName")}
                initialValue={`pure-regulation-${gitProjectName}`}
              >
                <Input />
              </Form.Item>
              <Alert
                message={`📍 ${t("initProject", "projectPath")}: pure-regulation-${gitProjectName}`}
                type="info"
                showIcon={false}
                style={{ marginBottom: 16 }}
              />
            </>
          )}
          {hasOutputFiles && (
            <>
              <Form.Item
                name="output_project_name"
                label={t("initProject", "outProjectName")}
                initialValue={`pure-output-${gitProjectName}`}
              >
                <Input />
              </Form.Item>
              <Alert
                message={`📍 ${t("initProject", "projectPath")}: pure-output-${gitProjectName}`}
                type="info"
                showIcon={false}
                style={{ marginBottom: 16 }}
              />
            </>
          )}
        </Form>
      </Modal>
    </div>
  );
}

// ── Main Component ──────────────────────────────────────

export default function InitProjectWizard() {
  const navigate = useNavigate();
  const { t } = useLanguage();
  const [mode, setMode] = useState<"jira_pdf" | "manual" | null>(null);

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
                <RocketOutlined style={{ fontSize: 24, color: "#fff" }} />
              </div>
              <div>
                <Title level={3} style={{ color: "#fff", margin: 0, fontWeight: 700 }}>{t("initProject", "title")}</Title>
                <Text style={{ color: "rgba(255,255,255,0.6)", fontSize: 13 }}>{t("initProject", "subtitle")}</Text>
              </div>
            </div>
          </div>
        </div>

        <div style={{ maxWidth: 1000, margin: "0 auto", padding: "32px 48px 48px" }}>
        <Card style={{ borderRadius: 16, border: "none", boxShadow: "0 4px 24px rgba(229,57,53,0.08)", marginTop: -32 }}>

          {mode && (
            <Button
              type="link"
              onClick={() => setMode(null)}
              style={{ padding: 0, marginBottom: 16 }}
            >
              {t("initProject", "changeMode")}
            </Button>
          )}

          <Divider />

          {!mode && <ModeSelection onSelectMode={setMode} />}
          {mode === "jira_pdf" && <JiraPdfWizard />}
          {mode === "manual" && <ManualWizard />}
        </Card>
        </div>
      </Content>
    </Layout>
  );
}
