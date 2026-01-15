const { useMemo, useState } = React;

function App() {
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState("");
  const [loading, setLoading] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState("");

  const fileName = useMemo(() => (file ? file.name : "No file selected"), [file]);

  function handleFileChange(event) {
    const nextFile = event.target.files[0] || null;
    setFile(nextFile);
    setStatus("");
    if (downloadUrl) {
      URL.revokeObjectURL(downloadUrl);
      setDownloadUrl("");
    }
  }

  async function handleSubmit(event) {
    event.preventDefault();
    if (!file) {
      setStatus("Please choose an .xlsx report first.");
      return;
    }

    setLoading(true);
    setStatus("Cleaning report...");

    const data = new FormData();
    data.append("file", file);

    try {
      const response = await fetch("/api/clean", {
        method: "POST",
        body: data,
      });

      if (!response.ok) {
        let message = "Upload failed.";
        try {
          const error = await response.json();
          if (error && error.detail) {
            message = error.detail;
          }
        } catch (err) {
          // ignore JSON parse errors
        }
        setStatus(message);
        setLoading(false);
        return;
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
      setStatus("Report ready.");
    } catch (err) {
      setStatus("Something went wrong. Please try again.");
    } finally {
      setLoading(false);
    }
  }

  return React.createElement(
    "section",
    { className: "hero" },
    React.createElement(
      "div",
      { className: "hero-copy" },
      React.createElement("h1", null, "Asset Report Cleanup"),
      React.createElement(
        "p",
        { className: "lede" },
        "Upload your raw account asset report and download a cleaned, formatted workbook ready for review."
      )
    ),
    React.createElement(
      "div",
      { className: "panel" },
      React.createElement(
        "form",
        { className: "upload-card", onSubmit: handleSubmit },
        React.createElement(
          "label",
          { className: "file-field" },
          React.createElement("input", {
            id: "file-input",
            name: "file",
            type: "file",
            accept: ".xlsx",
            onChange: handleFileChange,
          }),
          React.createElement("span", { className: "file-label" }, "Select report file"),
          React.createElement("span", { className: "file-name" }, fileName)
        ),
        React.createElement(
          "button",
          { type: "submit", className: "primary", disabled: loading },
          loading ? "Working..." : "Clean report"
        ),
        React.createElement("p", { className: "status", role: "status" }, status),
        downloadUrl
          ? React.createElement(
              "a",
              {
                className: "download",
                href: downloadUrl,
                download: file ? `cleaned_${file.name}` : "cleaned_report.xlsx",
              },
              "Download cleaned report"
            )
          : null
      )
    )
  );
}

const root = ReactDOM.createRoot(document.getElementById("root"));
root.render(React.createElement(App));
