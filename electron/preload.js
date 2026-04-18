const { contextBridge, ipcRenderer } = require("electron");
const QRCode = require("qrcode");
const SvgRenderer = require("qrcode/lib/renderer/svg-tag");

function renderQrSvg(text = "QR") {
  const content = String(text || "QR");
  const qrData = QRCode.create(content, {
    errorCorrectionLevel: "M",
    margin: 0,
  });

  return SvgRenderer.render(qrData, {
    margin: 0,
    color: {
      dark: "#111111",
      light: "#ffffff",
    },
  }).replace(
    "<svg ",
    '<svg shape-rendering="crispEdges" preserveAspectRatio="xMidYMid meet" '
  );
}

contextBridge.exposeInMainWorld("pigeonApi", {
  request(command, params = {}) {
    return ipcRenderer.invoke("backend:request", command, params);
  },
  chooseImage() {
    return ipcRenderer.invoke("dialog:open-image");
  },
  choosePngPath() {
    return ipcRenderer.invoke("dialog:save-png");
  },
  renderQrSvg(text) {
    return renderQrSvg(text);
  },
});
