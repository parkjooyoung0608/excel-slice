const fileInput = document.getElementById("fileInput");
const idInput = document.getElementById("idInput");
const extractBtn = document.getElementById("extractBtn");
const downloadBtn = document.getElementById("downloadBtn");
const resultDiv = document.getElementById("result");
const progressBar = document.getElementById("progressBar");

let workbookData = null;
let filteredData = null;

// 시뮬레이션 프로그레스 함수
function simulateProgress(duration = 2000) {
  return new Promise((resolve) => {
    let progress = 0;
    const interval = 20; // ms
    const step = 100 / (duration / interval);
    const timer = setInterval(() => {
      progress += step;
      if (progress >= 99) progress = 99; // 실제 완료 전까지 99%
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

  await simulateProgress(1500); // 1.5초 동안 프로그레스 바 시뮬레이션

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

// 추출 및 화면 표시
extractBtn.addEventListener("click", () => {
  if (!workbookData) return alert("먼저 엑셀 파일을 업로드하세요.");

  const firstSheetName = workbookData.SheetNames[0];
  const worksheet = workbookData.Sheets[firstSheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet);

  // 회원코드별 개수 세기
  const codeCount = {};
  jsonData.forEach((row) => {
    const code = row["회원코드"];
    if (code) {
      codeCount[code] = (codeCount[code] || 0) + 1;
    }
  });

  // 2개 이상인 회원코드만 필터링
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

// 다운로드
downloadBtn.addEventListener("click", () => {
  if (!filteredData || filteredData.length === 0) return;

  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.json_to_sheet(filteredData);
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Filtered");
  XLSX.writeFile(newWorkbook, "filtered.xlsx");
});

// 화면에 테이블 렌더링
function renderTable(data) {
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
