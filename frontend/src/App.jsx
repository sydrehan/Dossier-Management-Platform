import { useState } from "react";

function getUniqueFilename(prefix, ext) {
  return `${prefix}-${Date.now()}.${ext}`;
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

      const response = await fetch("http://localhost:3000/dossier/generate", {
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

      const response = await fetch("http://localhost:3000/merge", {
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
  // Document Converter State & Configuration
  // ---------------------------------------------------------
  const [converterFile, setConverterFile] = useState(null);
  const [convType, setConvType] = useState("pdf-to-docx");

  function handleConverterFile(file) {
    setErrorMsg(null);
    if (!file) {
      setConverterFile(null);
      return;
    }

    const ext = file.name.toLowerCase().substring(file.name.lastIndexOf('.'));
    if (ext === '.pdf') {
      setConverterFile(file);
      setConvType("pdf-to-docx");
    } else if (ext === '.pptx' || ext === '.ppt') {
      setConverterFile(file);
      setConvType("ppt-to-pdf");
    } else {
      setErrorMsg("Unsupported file type for conversion. Upload a PDF or PowerPoint file.");
    }
  }

  async function convertDocument() {
    setLoading(true);
    setLoadingMsg("Executing layout conversion pipelines and mapping element properties...");
    setResult(null);
    setErrorMsg(null);

    try {
      const formData = new FormData();
      formData.append("file", converterFile, converterFile.name);
      formData.append("convType", convType);

      const response = await fetch("http://localhost:3000/convert", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error(await getErrorMessage(response, "Conversion failed"));
      }

      const blob = await response.blob();
      const contentDisposition = response.headers.get("content-disposition") || "";
      const match = /filename="?([^";]+)"?/.exec(contentDisposition);
      const outExt = convType.endsWith("pdf") ? "pdf" : "docx";
      const filename = match ? match[1] : getUniqueFilename("converted", outExt);

      setResult({ blob, filename, type: `Converted Document (${outExt.toUpperCase()})` });
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
    } else if (activeTab === "converter") {
      setConverterFile(null);
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
                <span className="nav-icon">📖</span>
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
                <span className="nav-icon">🥞</span>
                Document Merger
              </button>
            </li>
            <li>
              <button
                className={`nav-item-btn ${activeTab === "converter" ? "active" : ""}`}
                onClick={() => {
                  setActiveTab("converter");
                  resetWorkspace();
                }}
              >
                <span className="nav-icon">🔁</span>
                Document Converter
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
            <div className="success-icon-badge">🎉</div>
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
                📥 Download Converted Document
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
                          <span className="check-chk-box">{metadataValid ? "✓" : ""}</span>
                          <span>Cover Metadata Valid</span>
                        </div>
                        <div className={`check-row-lbl ${activeUploadedCount > 0 ? "valid" : ""}`}>
                          <span className="check-chk-box">{activeUploadedCount > 0 ? "✓" : ""}</span>
                          <span>At Least One Section Mapped</span>
                        </div>
                      </div>

                      {errorMsg && <div className="dmp-error-banner">⚠ {errorMsg}</div>}

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
                                  ✕
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
                                <span className="up-ico">📤</span>
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
                        ⚡ Compile AMCP 5.0 Submission Dossier
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
                    <div className="converter-dropzone-icon">📤</div>
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
                              <span className="drag-handle-icon">☰</span>
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
                                  ▲
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
                                  ▼
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
                                ✕
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}

                  {errorMsg && <div className="dmp-error-banner">⚠ {errorMsg}</div>}
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
                        ⚡ Compile and Merge {mergerFiles.length} Documents
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}

            {activeTab === "converter" && (
              <div className="merger-center-layout">
                <div className="workspace-header">
                  <h1>Enterprise Document Converter</h1>
                  <p>Seamlessly convert individual files to industry-standard publication formats.</p>
                </div>

                <div className="converter-row-panel">
                  {/* File Upload zone */}
                  <div className="dmp-panel-card">
                    <div className="panel-section-title">
                      <h3>Select Document</h3>
                      <p>Drag and drop a PDF or PowerPoint file to convert.</p>
                    </div>

                    {converterFile ? (
                      <div className="file-list-row" style={{ marginTop: "12px" }}>
                        <div className="file-row-details">
                          <span className={`file-badge-type ${converterFile.name.toLowerCase().endsWith('.pdf') ? 'pdf' : 'generic'}`}>
                            {converterFile.name.split('.').pop()}
                          </span>
                          <span className="file-name-txt" title={converterFile.name}>
                            {converterFile.name}
                          </span>
                        </div>
                        <div className="file-row-actions">
                          <span className="file-size-lbl">{formatBytes(converterFile.size)}</span>
                          <button className="btn-file-row-clear" onClick={() => handleConverterFile(null)}>
                            ✕
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div
                        className="converter-dropzone-box"
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
                          if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                            handleConverterFile(e.dataTransfer.files[0]);
                          }
                        }}
                        onClick={() => document.getElementById("converter-file-input").click()}
                      >
                        <input
                          id="converter-file-input"
                          type="file"
                          accept=".pdf,.pptx,.ppt"
                          onChange={(e) => {
                            if (e.target.files && e.target.files[0]) {
                              handleConverterFile(e.target.files[0]);
                            }
                          }}
                          style={{ display: "none" }}
                        />
                        <div className="converter-dropzone-icon">📂</div>
                        <div className="converter-dropzone-text">
                          <h4>Drag & drop file here or browse</h4>
                          <p>Supports PDF, PPT, PPTX formats</p>
                        </div>
                      </div>
                    )}

                    {errorMsg && <div className="dmp-error-banner" style={{ marginTop: "12px" }}>⚠ {errorMsg}</div>}
                  </div>

                  {/* Settings / Convert Trigger */}
                  <div className="dmp-panel-card">
                    <div className="panel-section-title">
                      <h3>Conversion Pipeline</h3>
                      <p>Select the target translation module</p>
                    </div>

                    <div className="dmp-form-group">
                      <label className="dmp-lbl">Conversion Scheme</label>
                      <select
                        className="select-dropdown-style"
                        value={convType}
                        onChange={(e) => setConvType(e.target.value)}
                        disabled={!converterFile}
                      >
                        {!converterFile ? (
                          <option value="none">Upload a file first...</option>
                        ) : converterFile.name.toLowerCase().endsWith(".pdf") ? (
                          <option value="pdf-to-docx">PDF to Microsoft Word (.docx)</option>
                        ) : (
                          <>
                            <option value="ppt-to-pdf">PowerPoint to PDF (.pdf)</option>
                            <option value="ppt-to-docx">PowerPoint to Microsoft Word (.docx)</option>
                          </>
                        )}
                      </select>
                    </div>

                    <div style={{ marginTop: "auto", paddingTop: "20px" }}>
                      <button
                        className="btn-saas-primary"
                        onClick={convertDocument}
                        disabled={!converterFile}
                      >
                        ⚡ Convert Uploaded File
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </>
        )}
      </main>
    </div>
  );
}
