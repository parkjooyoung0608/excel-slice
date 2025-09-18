const fileInput = document.getElementById("fileInput");
const extractBtn = document.getElementById("extractBtn");
const downloadBtn = document.getElementById("downloadBtn");
const transformBtn = document.getElementById("transformBtn");
const downloadTransformBtn = document.getElementById("downloadTransformBtn");
const resultDiv = document.getElementById("result");
const progressBar = document.getElementById("progressBar");

let workbookData = null;
let filteredData = null;
let transformedData = [];
let lastRenderedData = [];

// 시뮬레이션 프로그레스 함수
function simulateProgress(duration = 2000) {
  return new Promise((resolve) => {
    let progress = 0;
    const interval = 20;
    const step = 100 / (duration / interval);
    const timer = setInterval(() => {
      progress += step;
      if (progress >= 99) progress = 99;
      progressBar.style.width = progress + "%";
      progressBar.textContent = Math.floor(progress) + "%";
    }, interval);
    setTimeout(() => {
      clearInterval(timer);
      resolve();
    }, duration);
  });
}

// 파일 읽기
fileInput.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  progressBar.style.width = "0%";
  progressBar.textContent = "0%";

  await simulateProgress(1500);

  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    workbookData = XLSX.read(data, { type: "array" });

    progressBar.style.width = "100%";
    progressBar.textContent = "완료!";
    alert("엑셀 파일이 로드되었습니다!");
  };
  reader.onerror = () => alert("파일 로드 중 오류가 발생했습니다.");

  reader.readAsArrayBuffer(file);
});

// 중복 추출 및 화면 표시
extractBtn.addEventListener("click", () => {
  if (!workbookData) return alert("먼저 엑셀 파일을 업로드하세요.");

  const firstSheetName = workbookData.SheetNames[0];
  const worksheet = workbookData.Sheets[firstSheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet);

  const codeCount = {};
  jsonData.forEach((row) => {
    const code = row["회원코드"];
    if (code) codeCount[code] = (codeCount[code] || 0) + 1;
  });

  filteredData = jsonData.filter((row) => codeCount[row["회원코드"]] >= 2);

  renderTable(filteredData);

  if (filteredData.length === 0) {
    alert("중복된 회원코드 데이터를 찾을 수 없습니다.");
    downloadBtn.disabled = true;
  } else {
    alert(`${filteredData.length}개의 중복 행이 추출되었습니다.`);
    downloadBtn.disabled = false;
  }
});

// 중복 추출 다운로드
downloadBtn.addEventListener("click", () => {
  if (!lastRenderedData || lastRenderedData.length === 0) return;
  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.json_to_sheet(lastRenderedData);
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Filtered");
  XLSX.writeFile(newWorkbook, "filtered.xlsx");
});

// 가로 변환 추출 및 화면 표시
function formatExcelDate(excelDate) {
  if (typeof excelDate === "number") {
    const date = new Date((excelDate - 25569) * 86400 * 1000);
    return date.toISOString().split("T")[0];
  }
  return excelDate;
}

transformBtn.addEventListener("click", () => {
  if (!workbookData) return alert("먼저 엑셀 파일을 업로드하세요.");

  const firstSheetName = workbookData.SheetNames[0];
  const worksheet = workbookData.Sheets[firstSheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

  const grouped = {};
  jsonData.forEach((row) => {
    const code = row["회원코드"];
    if (!code) return;
    if (!grouped[code]) grouped[code] = [];
    grouped[code].push(row);
  });

  transformedData = Object.entries(grouped).map(([code, rows]) => {
    rows.sort(
      (a, b) => new Date(a["첫 주문일자"]) - new Date(b["첫 주문일자"])
    );
    const base = { 회원코드: code };
    rows.forEach((row, idx) => {
      if (idx === 0) {
        base["주문자 E-Mail"] = row["주문자 E-Mail"];
        base["주문자명"] = row["주문자명"];
        base["주문자 연락처"] = row["주문자 연락처"];
        base["첫주문_주문일자"] = formatExcelDate(row["첫 주문일자"]);
        base["첫주문_옵션"] = row["옵션정보"];
        base["첫주문_수량"] = row["수량"];
      } else {
        if (row["첫 주문일자"] || row["옵션정보"] || row["수량"]) {
          base[`재구매${idx}_주문일자`] = formatExcelDate(row["첫 주문일자"]);
          base[`재구매${idx}_옵션`] = row["옵션정보"];
          base[`재구매${idx}_수량`] = row["수량"];
        }
      }
    });
    return base;
  });

  renderTable(transformedData);

  if (transformedData.length === 0) {
    alert("데이터를 변환할 수 없습니다.");
    downloadTransformBtn.disabled = true;
  } else {
    alert(
      `${transformedData.length}개의 회원코드 데이터가 가로 형태로 변환되었습니다.`
    );
    downloadTransformBtn.disabled = false;
  }
});

// 가로 변환 다운로드
downloadTransformBtn.addEventListener("click", () => {
  if (!lastRenderedData || lastRenderedData.length === 0) return;
  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.json_to_sheet(lastRenderedData);
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Transformed");
  XLSX.writeFile(newWorkbook, "transformed_result.xlsx");
});

function renderTable(data) {
  lastRenderedData = data;
  if (!data || data.length === 0) {
    resultDiv.innerHTML = "<p>결과가 없습니다.</p>";
    return;
  }
  const columns = Object.keys(data[0]);
  let html = "<table><thead><tr>";
  columns.forEach((col) => (html += `<th>${col}</th>`));
  html += "</tr></thead><tbody>";
  data.forEach((row) => {
    html += "<tr>";
    columns.forEach((col) => (html += `<td>${row[col] ?? ""}</td>`));
    html += "</tr>";
  });
  html += "</tbody></table>";
  resultDiv.innerHTML = html;
}
