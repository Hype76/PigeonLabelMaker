const { contextBridge, ipcRenderer } = require("electron");
const QRCode = require("qrcode");
const SvgRenderer = require("qrcode/lib/renderer/svg-tag");
const JsBarcode = require("jsbarcode");

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

function renderBarcodeSvg(text = "123456789") {
  const content = String(text || "123456789");
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  JsBarcode(svg, content, {
    format: "CODE128",
    displayValue: false,
    margin: 0,
    background: "#ffffff",
    lineColor: "#111111",
    width: 2,
    height: 64,
  });
  svg.setAttribute("shape-rendering", "crispEdges");
  svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
  return svg.outerHTML;
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
  renderBarcodeSvg(text) {
    return renderBarcodeSvg(text);
  },
});
