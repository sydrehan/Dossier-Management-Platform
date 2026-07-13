import { useState } from "react";
import { API } from "./config/api";

function getUniqueFilename(prefix, ext) {
  return `${prefix}-${Date.now()}.${ext}`;
}

// Premium Enterprise SVG Icons
function BookOpenIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle" }}>
      <path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z" />
      <path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z" />
    </svg>
  );
}

function LayersIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle" }}>
      <path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5" />
    </svg>
  );
}

function UploadCloudIcon() {
  return (
    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle" }}>
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12" />
    </svg>
  );
}

function CheckCircleIcon() {
  return (
    <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle" }}>
      <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
      <polyline points="22 4 12 14.01 9 11.01" />
    </svg>
  );
}

function CheckIcon() {
  return (
    <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle" }}>
      <polyline points="20 6 9 17 4 12" />
    </svg>
  );
}

function XIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle" }}>
      <line x1="18" y1="6" x2="6" y2="18" />
      <line x1="6" y1="6" x2="18" y2="18" />
    </svg>
  );
}

function DownloadIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle" }}>
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3" />
    </svg>
  );
}

function ChevronUpIcon() {
  return (
    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle" }}>
      <polyline points="18 15 12 9 6 15" />
    </svg>
  );
}

function ChevronDownIcon() {
  return (
    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle" }}>
      <polyline points="6 9 12 15 18 9" />
    </svg>
  );
}

function DragHandleIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle" }}>
      <circle cx="9" cy="12" r="1" />
      <circle cx="9" cy="5" r="1" />
      <circle cx="9" cy="19" r="1" />
      <circle cx="15" cy="12" r="1" />
      <circle cx="15" cy="5" r="1" />
      <circle cx="15" cy="19" r="1" />
    </svg>
  );
}

function FlashIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle" }}>
      <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2" />
    </svg>
  );
}

function WarningIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: "inline-block", verticalAlign: "middle", flexShrink: 0 }}>
      <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
      <line x1="12" y1="9" x2="12" y2="13" />
      <line x1="12" y1="17" x2="12.01" y2="17" />
    </svg>
  );
}

export default function App() {
  const [activeTab, setActiveTab] = useState("builder"); // "builder" | "merger" | "converter"
  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState("");
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);
  const [dragOverSlot, setDragOverSlot] = useState(null);
  const [isDragOverDropzone, setIsDragOverDropzone] = useState(false);

  // ---------------------------------------------------------
  // Dossier Builder State & Configuration
  // ---------------------------------------------------------
  const [metadata, setMetadata] = useState({
    title: "AMCP Format for Formulary Submissions 5.0",
    brandName: "SIMU HIV-1 Quant System",
    genericName: "quantitative nucleic acid-based testing technology",
    manufacturer: "THE Manufacturing Company",
    version: "5.0",
    date: "April 2024",
    dossierType: "FormatB", // "FormatA" (Pipeline) or "FormatB" (Approved)
  });

  const [mappings, setMappings] = useState({
    highlights: null,
    exec_summary: null,
    product_info: null,
    clinical: null,
    economic: null,
    supporting: null,
    references: null,
  });

  const formatBSections = [
    { key: "exec_summary", num: "1.0B", label: "Executive Summary: Clinical & Economic Value" },
    { key: "product_info", num: "2.0B", label: "Product Information & Disease Description" },
    { key: "clinical", num: "3.0B", label: "Clinical Evidence" },
    { key: "economic", num: "4.0B", label: "Economic Value & Modeling Report" },
    { key: "supporting", num: "5.0B", label: "Additional Supporting Evidence" },
  ];

  const formatASections = [
    { key: "highlights", num: "1.0A", label: "Highlights and Overview" },
    { key: "product_info", num: "2.0A", label: "Product Information & Disease Description" },
    { key: "clinical", num: "3.0A", label: "Clinical Evidence" },
    { key: "economic", num: "4.0A", label: "Economic Information" },
    { key: "supporting", num: "5.0A", label: "Additional Supporting Evidence" },
  ];

  const activeSections = metadata.dossierType === "FormatB" ? formatBSections : formatASections;

  const activeUploadedCount = Object.entries(mappings)
    .filter(([key, val]) => activeSections.map(s => s.key).includes(key) && val !== null)
    .length;

  const metadataValid = 
    metadata.title.trim().length > 0 &&
    metadata.brandName.trim().length > 0 &&
    metadata.genericName.trim().length > 0 &&
    metadata.manufacturer.trim().length > 0 &&
    metadata.version.trim().length > 0 &&
    metadata.date.trim().length > 0;

  function handleSectionFile(sectionKey, file) {
    setErrorMsg(null);
    if (!file) {
      setMappings((prev) => ({ ...prev, [sectionKey]: null }));
      return;
    }
    const ext = file.name.toLowerCase().substring(file.name.lastIndexOf('.'));
    if (ext !== '.pdf' && ext !== '.docx') {
      setErrorMsg("Supported formats are PDF and DOCX only.");
      return;
    }
    setMappings((prev) => ({ ...prev, [sectionKey]: file }));
  }

  async function generateDossier() {
    setLoading(true);
    setLoadingMsg("Surgically injecting sections, normalizing Arial styles, and matching header/footer formats...");
    setResult(null);
    setErrorMsg(null);

    try {
      const formData = new FormData();
      const fileToSectionMap = {};

      Object.entries(mappings).forEach(([sectionKey, file]) => {
        if (file) {
          formData.append("files", file, file.name);
          fileToSectionMap[sectionKey] = file.name;
        }
      });

      const submitMetadata = {
        ...metadata,
        versionNumber: metadata.version,
        compilationDate: metadata.date,
        dossierType: metadata.dossierType === "FormatB" ? "format-b" : "format-a"
      };

      formData.append("metadata", JSON.stringify(submitMetadata));
      formData.append("mappings", JSON.stringify(fileToSectionMap));

      const response = await fetch(API.dossierGenerate, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error(await getErrorMessage(response, "Dossier compilation failed"));
      }

      const blob = await response.blob();
      const contentDisposition = response.headers.get("content-disposition") || "";
      const match = /filename="?([^";]+)"?/.exec(contentDisposition);
      const filename = match ? match[1] : getUniqueFilename("dossier", "docx");

      setResult({ blob, filename, type: "AMCP Submission Dossier (DOCX)" });
      triggerDownload(blob, filename);
    } catch (err) {
      console.error(err);
      setErrorMsg(err.message);
    } finally {
      setLoading(false);
    }
  }

  // ---------------------------------------------------------
  // Document Merger State & Configuration
  // ---------------------------------------------------------
  const [mergerFiles, setMergerFiles] = useState([]);
  const [mergerType, setMergerType] = useState("pdf"); // "pdf" | "docx" | "pdf-docx"

  function addMergerFiles(files) {
    setErrorMsg(null);
    const validFiles = Array.from(files).filter(file => {
      const ext = file.name.toLowerCase().substring(file.name.lastIndexOf('.'));
      return ext === '.pdf' || ext === '.docx' || ext === '.pptx' || ext === '.ppt';
    });

    if (validFiles.length !== files.length) {
      setErrorMsg("Some files were skipped. Only PDF, DOCX, and PPTX formats are supported.");
    }

    setMergerFiles(prev => [...prev, ...validFiles]);
  }

  function moveMergerItem(index, direction) {
    const newIndex = index + direction;
    if (newIndex < 0 || newIndex >= mergerFiles.length) return;
    const updated = [...mergerFiles];
    const temp = updated[index];
    updated[index] = updated[newIndex];
    updated[newIndex] = temp;
    setMergerFiles(updated);
  }

  function removeMergerItem(index) {
    setMergerFiles(prev => prev.filter((_, i) => i !== index));
  }

  async function mergeDocuments() {
    setLoading(true);
    setLoadingMsg("Converting slide formats, appending section compositions, and compiling merged package...");
    setResult(null);
    setErrorMsg(null);

    try {
      const formData = new FormData();
      mergerFiles.forEach(file => {
        formData.append("files", file, file.name);
      });
      formData.append("type", mergerType);

      const response = await fetch(API.merge, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error(await getErrorMessage(response, "Document merging failed"));
      }

      const blob = await response.blob();
      const contentDisposition = response.headers.get("content-disposition") || "";
      const match = /filename="?([^";]+)"?/.exec(contentDisposition);
      const filename = match ? match[1] : getUniqueFilename("merged", mergerType === "pdf" ? "pdf" : "docx");

      setResult({ blob, filename, type: `Merged Document (${mergerType.toUpperCase()})` });
      triggerDownload(blob, filename);
    } catch (err) {
      console.error(err);
      setErrorMsg(err.message);
    } finally {
      setLoading(false);
    }
  }

  // ---------------------------------------------------------
  // Utilities
  // ---------------------------------------------------------
  async function getErrorMessage(response, defaultMsg) {
    try {
      const data = await response.json();
      return data.error || defaultMsg;
    } catch {
      try {
        const text = await response.text();
        return text || defaultMsg;
      } catch {
        return defaultMsg;
      }
    }
  }

  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function formatBytes(bytes) {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
  }

  function resetWorkspace() {
    setResult(null);
    setErrorMsg(null);
    if (activeTab === "builder") {
      setMappings({
        highlights: null,
        exec_summary: null,
        product_info: null,
        clinical: null,
        economic: null,
        supporting: null,
        references: null,
      });
    } else if (activeTab === "merger") {
      setMergerFiles([]);
    }
  }

  return (
    <div className="dmp-layout">
      {/* Sidebar Controls */}
      <aside className="dmp-sidebar">
        <div>
          <div className="brand-section">
            <div className="brand-icon-box">D</div>
            <div className="brand-text-box">
              <h2>Dossier Management</h2>
              <span>Enterprise Hub</span>
            </div>
          </div>

          <ul className="nav-links-list">
            <li>
              <button
                className={`nav-item-btn ${activeTab === "builder" ? "active" : ""}`}
                onClick={() => {
                  setActiveTab("builder");
                  resetWorkspace();
                }}
              >
                <span className="nav-icon"><BookOpenIcon /></span>
                Dossier Builder
              </button>
            </li>
            <li>
              <button
                className={`nav-item-btn ${activeTab === "merger" ? "active" : ""}`}
                onClick={() => {
                  setActiveTab("merger");
                  resetWorkspace();
                }}
              >
                <span className="nav-icon"><LayersIcon /></span>
                Document Merger
              </button>
            </li>

          </ul>
        </div>

        <div className="sidebar-footer-status">
          <div className="status-dot-pulse"></div>
          <span className="status-lbl">Cluster: Online</span>
        </div>
      </aside>

      {/* Main workspace container */}
      <main className="dmp-workspace">
        {/* Loader Overlay */}
        {loading ? (
          <div className="workspace-loader-panel">
            <div className="saas-spinner-indicator"></div>
            <div>
              <p className="loader-main-text">Processing Request</p>
              <p className="loader-sub-text">{loadingMsg}</p>
            </div>
          </div>
        ) : result ? (
          /* Process completed Success Panel */
          <div className="workspace-success-panel">
            <div className="success-icon-badge"><CheckCircleIcon /></div>
            <h2 className="success-main-title">Operation Successful!</h2>
            <p className="success-sub-desc">
              Your document has been compiled and formatted to meet enterprise layout specifications.
            </p>

            <div className="success-output-details-box">
              <div className="success-details-heading">Job Execution Log</div>
              <div className="success-details-row">
                <span className="success-details-lbl">Job Output:</span>
                <span className="success-details-val highlight-txt">{result.filename}</span>
              </div>
              <div className="success-details-row">
                <span className="success-details-lbl">Process Type:</span>
                <span className="success-details-val">{result.type}</span>
              </div>
              <div className="success-details-row">
                <span className="success-details-lbl">Status:</span>
                <span className="success-details-val" style={{ color: "var(--success)" }}>Validated & Signed</span>
              </div>
            </div>

            <div className="success-control-buttons">
              <button className="btn-success-action" onClick={() => triggerDownload(result.blob, result.filename)}>
                <DownloadIcon /> Download Converted Document
              </button>
              <button className="btn-reset-light-style" onClick={resetWorkspace}>
                Process Another File
              </button>
            </div>
          </div>
        ) : (
          /* Regular workspaces based on tabs */
          <>
            {activeTab === "builder" && (
              <>
                <div className="workspace-header">
                  <h1>AMCP Dossier Builder</h1>
                  <p>Surgically populate the official template guidelines with product metadata and formatted sections.</p>
                </div>

                <div className="builder-split-layout">
                  {/* Settings Panel */}
                  <aside style={{ display: "flex", flexDirection: "column", gap: "20px" }}>
                    <div className="dmp-panel-card">
                      <div className="panel-section-title">
                        <h3>Dossier Settings</h3>
                        <p>Customize the Cover Page and section mappings</p>
                      </div>

                      <div className="dmp-form-group">
                        <label className="dmp-lbl">Submission Protocol</label>
                        <div className="dmp-segmented-bar">
                          <button
                            type="button"
                            className={`dmp-segment-btn ${metadata.dossierType === "FormatB" ? "active" : ""}`}
                            onClick={() => setMetadata({ ...metadata, dossierType: "FormatB", title: "AMCP Format for Formulary Submissions 5.0" })}
                          >
                            Approved (Format B)
                          </button>
                          <button
                            type="button"
                            className={`dmp-segment-btn ${metadata.dossierType === "FormatA" ? "active" : ""}`}
                            onClick={() => setMetadata({ ...metadata, dossierType: "FormatA", title: "AMCP Format for Formulary Submissions 5.0 (Pipeline)" })}
                          >
                            Pipeline (Format A)
                          </button>
                        </div>
                      </div>

                      <div className="dmp-form-group">
                        <label className="dmp-lbl">
                          Dossier Title <span className="dmp-required">*</span>
                        </label>
                        <textarea
                          className="dmp-input"
                          rows={2}
                          value={metadata.title}
                          onChange={(e) => setMetadata({ ...metadata, title: e.target.value })}
                          placeholder="Title..."
                        />
                      </div>

                      <div className="dmp-form-row">
                        <div className="dmp-form-group">
                          <label className="dmp-lbl">
                            Brand Name <span className="dmp-required">*</span>
                          </label>
                          <input
                            className="dmp-input"
                            type="text"
                            value={metadata.brandName}
                            onChange={(e) => setMetadata({ ...metadata, brandName: e.target.value })}
                            placeholder="e.g. SIMU HIV-1"
                          />
                        </div>
                        <div className="dmp-form-group">
                          <label className="dmp-lbl">
                            Generic Name <span className="dmp-required">*</span>
                          </label>
                          <input
                            className="dmp-input"
                            type="text"
                            value={metadata.genericName}
                            onChange={(e) => setMetadata({ ...metadata, genericName: e.target.value })}
                            placeholder="e.g. Assay"
                          />
                        </div>
                      </div>

                      <div className="dmp-form-row">
                        <div className="dmp-form-group">
                          <label className="dmp-lbl">
                            Manufacturer <span className="dmp-required">*</span>
                          </label>
                          <input
                            className="dmp-input"
                            type="text"
                            value={metadata.manufacturer}
                            onChange={(e) => setMetadata({ ...metadata, manufacturer: e.target.value })}
                            placeholder="e.g. KadelLabs"
                          />
                        </div>
                        <div className="dmp-form-group">
                          <label className="dmp-lbl">
                            Version <span className="dmp-required">*</span>
                          </label>
                          <input
                            className="dmp-input code-style"
                            type="text"
                            value={metadata.version}
                            onChange={(e) => setMetadata({ ...metadata, version: e.target.value })}
                            placeholder="5.0"
                          />
                        </div>
                      </div>

                      <div className="dmp-form-group">
                        <label className="dmp-lbl">
                          Compilation Date <span className="dmp-required">*</span>
                        </label>
                        <input
                          className="dmp-input code-style"
                          type="text"
                          value={metadata.date}
                          onChange={(e) => setMetadata({ ...metadata, date: e.target.value })}
                          placeholder="April 2024"
                        />
                      </div>
                    </div>

                    <div className="dmp-panel-card">
                      <div className="panel-section-title">
                        <h3>Workspace Checklist</h3>
                        <p>Verify parameters before compiling the final dossier</p>
                      </div>

                      <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
                        <div className={`check-row-lbl ${metadataValid ? "valid" : ""}`}>
                          <span className="check-chk-box">{metadataValid ? <CheckIcon /> : ""}</span>
                          <span>Cover Metadata Valid</span>
                        </div>
                        <div className={`check-row-lbl ${activeUploadedCount > 0 ? "valid" : ""}`}>
                          <span className="check-chk-box">{activeUploadedCount > 0 ? <CheckIcon /> : ""}</span>
                          <span>At Least One Section Mapped</span>
                        </div>
                      </div>

                      {errorMsg && <div className="dmp-error-banner"><WarningIcon /> {errorMsg}</div>}

                      {activeUploadedCount > 0 && (
                        <button className="btn-reset-light-style" onClick={resetWorkspace}>
                          Reset Workspace
                        </button>
                      )}
                    </div>
                  </aside>

                  {/* Section Slots list */}
                  <main className="section-slots-vertical-list">
                    <div className="slots-header-clean">
                      <div>
                        <h3>Template Content Mapping</h3>
                        <p className="slots-description">Map individual PDF or DOCX files to the correct AMCP sections</p>
                      </div>
                    </div>

                    {activeSections.map((sec) => {
                      const file = mappings[sec.key];
                      const isDragActive = dragOverSlot === sec.key;

                      return (
                        <div
                          key={sec.key}
                          className={`dossier-card-slot ${file ? "filled" : ""} ${isDragActive ? "drag-over" : ""}`}
                          onDragEnter={(e) => {
                            e.preventDefault();
                            setDragOverSlot(sec.key);
                          }}
                          onDragOver={(e) => {
                            e.preventDefault();
                          }}
                          onDragLeave={() => setDragOverSlot(null)}
                          onDrop={(e) => {
                            e.preventDefault();
                            setDragOverSlot(null);
                            if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                              handleSectionFile(sec.key, e.dataTransfer.files[0]);
                            }
                          }}
                        >
                          <div className="slot-card-meta">
                            <span className="slot-card-badge">{sec.num}</span>
                            <span className="slot-card-title-lbl">{sec.label}</span>
                          </div>

                          {file ? (
                            <div className="file-list-row">
                              <div className="file-row-details">
                                <span className={`file-badge-type ${file.name.toLowerCase().endsWith('.pdf') ? 'pdf' : 'docx'}`}>
                                  {file.name.toLowerCase().endsWith('.pdf') ? 'pdf' : 'docx'}
                                </span>
                                <span className="file-name-txt" title={file.name}>
                                  {file.name}
                                </span>
                              </div>
                              <div className="file-row-actions">
                                <span className="file-size-lbl">{formatBytes(file.size)}</span>
                                <button
                                  className="btn-file-row-clear"
                                  onClick={() => handleSectionFile(sec.key, null)}
                                >
                                  <XIcon />
                                </button>
                              </div>
                            </div>
                          ) : (
                            <div className="dropzone-saas-area">
                              <input
                                id={`file-input-${sec.key}`}
                                type="file"
                                accept=".pdf,.docx"
                                onChange={(e) => {
                                  if (e.target.files && e.target.files[0]) {
                                    handleSectionFile(sec.key, e.target.files[0]);
                                  }
                                }}
                                style={{ display: "none" }}
                              />
                              <label htmlFor={`file-input-${sec.key}`} className="dropzone-saas-lbl">
                                <span className="up-ico"><UploadCloudIcon /></span>
                                <span className="dropzone-saas-txt">
                                  Drag file here or <span className="hl-browse">browse</span>
                                </span>
                              </label>
                            </div>
                          )}
                        </div>
                      );
                    })}

                    <div style={{ marginTop: "12px" }}>
                      <button
                        className="btn-saas-primary"
                        disabled={!metadataValid || activeUploadedCount === 0}
                        onClick={generateDossier}
                      >
                        <FlashIcon /> Compile AMCP 5.0 Submission Dossier
                      </button>
                    </div>
                  </main>
                </div>
              </>
            )}

            {activeTab === "merger" && (
              <div className="merger-center-layout">
                <div className="workspace-header">
                  <h1>Multi-Format Document Merger</h1>
                  <p>Combine multiple PDF, DOCX, and PPTX documents into a single professional PDF or DOCX file.</p>
                </div>

                <div className="dmp-panel-card merger-uploads-card">
                  <div className="panel-section-title">
                    <h3>Uploaded Files ({mergerFiles.length})</h3>
                    <p>Select multiple documents. Reorder them using the controls to determine merge sequence.</p>
                  </div>

                  {/* Drag and Drop Zone */}
                  <div
                    className={`converter-dropzone-box ${isDragOverDropzone ? "active" : ""}`}
                    onDragEnter={(e) => {
                      e.preventDefault();
                      setIsDragOverDropzone(true);
                    }}
                    onDragOver={(e) => {
                      e.preventDefault();
                    }}
                    onDragLeave={() => setIsDragOverDropzone(false)}
                    onDrop={(e) => {
                      e.preventDefault();
                      setIsDragOverDropzone(false);
                      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                        addMergerFiles(e.dataTransfer.files);
                      }
                    }}
                    onClick={() => document.getElementById("merger-file-input").click()}
                  >
                    <input
                      id="merger-file-input"
                      type="file"
                      multiple
                      accept=".pdf,.docx,.pptx,.ppt"
                      onChange={(e) => {
                        if (e.target.files && e.target.files.length > 0) {
                          addMergerFiles(e.target.files);
                        }
                      }}
                      style={{ display: "none" }}
                    />
                    <div className="converter-dropzone-icon"><UploadCloudIcon /></div>
                    <div className="converter-dropzone-text">
                      <h4>Drag and drop multiple documents here</h4>
                      <p>Supports PDF, DOCX, PPT, PPTX formats</p>
                    </div>
                  </div>

                  {/* File List */}
                  {mergerFiles.length > 0 && (
                    <div className="merger-items-list">
                      {mergerFiles.map((file, index) => {
                        const ext = file.name.toLowerCase().substring(file.name.lastIndexOf('.'));
                        const extClass = ext === '.pdf' ? 'pdf' : (ext === '.docx' ? 'docx' : 'generic');
                        return (
                          <div key={index} className="merger-item-row">
                            <div className="merger-item-details">
                              <span className="drag-handle-icon"><DragHandleIcon /></span>
                              <span className={`file-badge-type ${extClass}`}>{ext.replace('.', '')}</span>
                              <span className="merger-item-name" title={file.name}>{file.name}</span>
                            </div>

                            <div className="merger-item-actions">
                              <span className="file-size-lbl">{formatBytes(file.size)}</span>
                              
                              <div style={{ display: "flex", gap: "4px" }}>
                                <button
                                  className="btn-file-row-clear"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    moveMergerItem(index, -1);
                                  }}
                                  disabled={index === 0}
                                  title="Move Up"
                                  style={{ fontSize: "0.7rem", opacity: index === 0 ? 0.3 : 1 }}
                                >
                                  <ChevronUpIcon />
                                </button>
                                <button
                                  className="btn-file-row-clear"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    moveMergerItem(index, 1);
                                  }}
                                  disabled={index === mergerFiles.length - 1}
                                  title="Move Down"
                                  style={{ fontSize: "0.7rem", opacity: index === mergerFiles.length - 1 ? 0.3 : 1 }}
                                >
                                  <ChevronDownIcon />
                                </button>
                              </div>

                              <span className="merger-item-order-badge">{index + 1}</span>

                              <button
                                className="btn-file-row-clear"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  removeMergerItem(index);
                                }}
                                title="Remove File"
                              >
                                <XIcon />
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}

                  {errorMsg && <div className="dmp-error-banner"><WarningIcon /> {errorMsg}</div>}
                </div>

                {mergerFiles.length > 0 && (
                  <div className="dmp-panel-card">
                    <div className="panel-section-title">
                      <h3>Compile Settings</h3>
                      <p>Select target file structure options</p>
                    </div>

                    <div className="dmp-form-row">
                      <div className="dmp-form-group">
                        <label className="dmp-lbl">Target Format</label>
                        <select
                          className="select-dropdown-style"
                          value={mergerType}
                          onChange={(e) => setMergerType(e.target.value)}
                        >
                          <option value="pdf">Adobe Portable Document Format (.pdf)</option>
                          <option value="docx">Microsoft Word Template (.docx)</option>
                          <option value="pdf-docx">Hybrid Compilation Package (.docx)</option>
                        </select>
                      </div>
                    </div>

                    <div style={{ marginTop: "12px" }}>
                      <button
                        className="btn-saas-primary"
                        onClick={mergeDocuments}
                        disabled={mergerFiles.length < 2}
                      >
                        <FlashIcon /> Compile and Merge {mergerFiles.length} Documents
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}


          </>
        )}
      </main>
    </div>
  );
}
