// app.js
// - Upload ảnh QR (không cần camera) -> jsQR decode
// - Parse QR CCCD phổ biến dạng: cccd|cmnd_cu|ho_ten|ddmmyyyy|gioi_tinh|dia_chi|ddmmyyyy(ngay_cap)|...
// - Fill form input theo id trong index.html
// - Render template.docx bằng docxtemplater -> tải file DOCX

const FIELD_IDS = [
  "NGAY","NGAY_TRA_HANG","HO_TEN","SDT","CCCD","NGAY_CAP","NOI_CAP","DIA_CHI","TIEN_BANG_CHU",
  "1L_SL","1L_KL1","1L_TKL","1L_DG","1L_TT",
  "5L_SL","5L_KL1","5L_TKL","5L_DG","5L_TT",
  "1KG_SL","1KG_KL1","1KG_TKL","1KG_DG","1KG_TT",
];

function $(id) { return document.getElementById(id); }

function setVal(id, val) {
  const el = $(id);
  if (el) el.value = val ?? "";
}

function getData() {
  const data = {};
  for (const k of FIELD_IDS) data[k] = ($(k)?.value ?? "").trim();
  return data;
}

function fmtDateFromDDMMYYYY(ddmmyyyy) {
  if (!ddmmyyyy || ddmmyyyy.length !== 8) return "";
  return `${ddmmyyyy.slice(0,2)}/${ddmmyyyy.slice(2,4)}/${ddmmyyyy.slice(4)}`;
}

function fillInputsFromQrText(qrText) {
  // CCCD QR thường: cccd|cmnd_cu|ho_ten|ddmmyyyy|gioi_tinh|dia_chi|ddmmyyyy(ngay_cap)|...
  const parts = (qrText || "").split("|");
  if (parts.length < 7) throw new Error("QR không đúng định dạng CCCD phổ biến.");

  const cccd = parts[0] || "";
  const hoTen = parts[2] || "";
  const diaChi = parts[5] || "";
  const ngayCapRaw = parts[6] || "";

  setVal("HO_TEN", hoTen);
  setVal("CCCD", cccd);
  setVal("NGAY_CAP", fmtDateFromDDMMYYYY(ngayCapRaw));
  setVal("DIA_CHI", diaChi);

  // QR CCCD thường KHÔNG có SĐT và Nơi cấp:
  // setVal("SDT", "");
  // setVal("NOI_CAP", "");
}

function decodeQrFromImageFile(file) {
  return new Promise((resolve, reject) => {
    const img = new Image();

    img.onload = () => {
      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d", { willReadFrequently: true });

      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      ctx.drawImage(img, 0, 0);

      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const code = jsQR(imageData.data, imageData.width, imageData.height);

      if (!code) return reject(new Error("Không đọc được QR. Thử ảnh rõ hơn/ít rung/đủ sáng."));
      resolve(code.data);
    };

    img.onerror = () => reject(new Error("Ảnh lỗi hoặc không đọc được."));

    const url = URL.createObjectURL(file);
    img.src = url;
  });
}

function bindQrUpload() {
  const input = $("qrFile");
  const status = $("qr-status");
  const preview = $("qr-preview");

  if (!input) return;

  input.addEventListener("change", async () => {
    const file = input.files && input.files[0];
    if (!file) return;

    if (status) status.textContent = "Đang đọc ảnh…";
    if (preview) {
      preview.style.display = "block";
      preview.src = URL.createObjectURL(file);
    }

    try {
      const decodedText = await decodeQrFromImageFile(file);
      fillInputsFromQrText(decodedText);
      if (status) status.textContent = "Quét xong ✅ Đã tự điền thông tin vào form.";
    } catch (e) {
      console.error(e);
      if (status) status.textContent = "Lỗi: " + e.message;
    }
  });
}

async function fetchTemplateArrayBuffer() {
  const res = await fetch("./template.docx", { cache: "no-store" });
  if (!res.ok) {
    throw new Error("Không tải được template.docx. Kiểm tra repo có file template.docx ở cùng thư mục với index.html.");
  }
  return await res.arrayBuffer();
}

function bindGenerateDocx() {
  const btn = $("btn");
  const status = $("status");
  if (!btn) return;

  btn.addEventListener("click", async () => {
    if (status) status.textContent = "";

    try {
      if (status) status.textContent = "Đang tải template…";
      const content = await fetchTemplateArrayBuffer();

      if (status) status.textContent = "Đang render DOCX…";
      const zip = new PizZip(content);
      const doc = new window.docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      doc.setData(getData());
      doc.render();

      const out = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      saveAs(out, "phieu_chot_gia_da_dien.docx");
      if (status) status.textContent = "Xong ✅ File đã được tải xuống.";
    } catch (e) {
      console.error(e);

      // Thông báo dễ hiểu:
      // 1) template.docx không tải được
      // 2) Tag trong template bị thiếu/sai hoặc tag bị tách định dạng trong Word
      if (status) {
        status.textContent =
          "Lỗi: " + (e?.message || "Không rõ") +
          " | Thường do template.docx thiếu hoặc tag {{...}} trong Word bị sai/tách định dạng.";
      }
    }
  });
}

window.addEventListener("load", () => {
  bindQrUpload();
  bindGenerateDocx();
});
