const ids = [
  "NGAY","NGAY_TRA_HANG","HO_TEN","SDT","CCCD","NGAY_CAP","NOI_CAP","DIA_CHI","TIEN_BANG_CHU",
  "1L_SL","1L_KL1","1L_TKL","1L_DG","1L_TT",
  "5L_SL","5L_KL1","5L_TKL","5L_DG","5L_TT",
  "1KG_SL","1KG_KL1","1KG_TKL","1KG_DG","1KG_TT",
];

function getData(){
  const data = {};
  for (const k of ids) data[k] = (document.getElementById(k)?.value || "").trim();
  return data;
}

async function fetchTemplateArrayBuffer(){
  const res = await fetch("./template.docx");
  if(!res.ok) throw new Error("Không tải được template.docx (kiểm tra file có nằm đúng repo không).");
  return await res.arrayBuffer();
}

document.getElementById("btn").addEventListener("click", async () => {
  const status = document.getElementById("status");
  status.textContent = "";

  try{
    status.textContent = "Đang tải template…";
    const content = await fetchTemplateArrayBuffer();

    status.textContent = "Đang render…";
    const zip = new PizZip(content);
    const doc = new window.docxtemplater(zip, { paragraphLoop:true, linebreaks:true });
    doc.setData(getData());
    doc.render();

    const out = doc.getZip().generate({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    saveAs(out, "phieu_chot_gia_da_dien.docx");
    status.textContent = "Xong. File đã tải xuống.";
  }catch(e){
    console.error(e);
    status.textContent =
      "Lỗi: thường do tag {{...}} trong Word bị tách định dạng hoặc thiếu tag. Mở Console (F12) để xem chi tiết.";
  }
});
