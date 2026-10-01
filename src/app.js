import * as XLSX from 'xlsx';

/**
 * ==========================================================================
 * 숭실대학교 기숙사 생활 도우미 (SSU DORM) - 메인 자바스크립트
 * ==========================================================================
 * 주요 기능:
 * 1. 상단 현재 날짜 및 '오늘의 기숙사 한마디' 표시
 * 2. 세탁기 / 건조기 실시간 사용 시작 및 타이머 카운트다운
 * 3. 택배 목록 조회 및 수령 완료 처리
 * 4. 시설 고장 신고 폼 유효성 검사 및 접수 완료 메시지
 * 5. 입력값 실시간 엑셀(Excel) 대장 자동 연동 및 .xlsx/.csv 다운로드
 * 6. 공용 물품 대여 / 반납 confirm 처리
 * 7. 상단 요약 정보 (세탁기, 택배, 공지) 자동 실시간 연동
 * 8. 🌙 다크 모드 토글 (밝은 모드 <-> 다크 모드)
 * 9. 모바일 하단 네비게이션 부드러운 스크롤 이동
 */

// DOM이 완전히 로드된 후 스크립트 실행
document.addEventListener('DOMContentLoaded', () => {
  // ------------------------------------------------------------------------
  // 1. 날짜 및 오늘의 한마디 초기화
  // ------------------------------------------------------------------------
  initDateDisplay();
  initDormQuotes();

  // ------------------------------------------------------------------------
  // 2. 다크 모드 초기화
  // ------------------------------------------------------------------------
  initThemeToggle();

  // ------------------------------------------------------------------------
  // 3. 각 기능별 모듈 초기화
  // ------------------------------------------------------------------------
  initLaundrySection();
  initDeliverySection();
  initRepairSection();
  initExcelSection();
  initRentalSection();
  initNavigation();

  // ------------------------------------------------------------------------
  // 4. 상단 요약 정보 최초 1회 갱신
  // ------------------------------------------------------------------------
  updateSummaryStats();
});

/* ==========================================================================
   1. 현재 날짜 표시 함수
   ========================================================================== */
function initDateDisplay() {
  const dateElement = document.getElementById('current-date');
  if (!dateElement) return;

  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const date = now.getDate();
  const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
  const dayName = dayNames[now.getDay()];

  // 예: "2026. 10. 1 (목)" 형태로 표시
  dateElement.textContent = `${year}년 ${month}월 ${date}일 (${dayName})`;
}

/* ==========================================================================
   2. 오늘의 기숙사 한마디 팁 목록 & 변경 기능
   ========================================================================== */
const dormQuotes = [
  '세탁이 끝난 빨래는 빠르게 수거해주세요 🧺',
  '밤 11시 이후 정숙 시간 준수로 서로를 배려해요 🤫',
  '외출 시 창문과 냉난방기 전원을 꼭 확인해주세요 💡',
  '택배함에 보관된 택배는 분실 방지를 위해 바로 수령해요 📦',
  '기숙사 내 허가되지 않은 전열기구 사용은 금지됩니다 🚫',
  '공용 물품은 사용 후 제자리에 깨끗이 반납해주세요 🧹'
];

let currentQuoteIndex = 0;

function initDormQuotes() {
  const quoteTextElement = document.getElementById('dorm-quote-text');
  const quoteRefreshBtn = document.getElementById('btn-quote-refresh');

  if (quoteTextElement) {
    quoteTextElement.textContent = dormQuotes[currentQuoteIndex];
  }

  if (quoteRefreshBtn && quoteTextElement) {
    quoteRefreshBtn.addEventListener('click', () => {
      currentQuoteIndex = (currentQuoteIndex + 1) % dormQuotes.length;
      quoteTextElement.textContent = dormQuotes[currentQuoteIndex];
    });
  }
}

/* ==========================================================================
   3. 다크모드 전환 기능
   ========================================================================== */
function initThemeToggle() {
  const themeToggleBtn = document.getElementById('btn-theme-toggle');
  if (!themeToggleBtn) return;

  // 로컬스토리지에 저장된 테마 불러오기 (기본은 라이트 모드)
  const savedTheme = localStorage.getItem('ssu-dorm-theme');
  if (savedTheme === 'dark') {
    document.body.classList.add('dark-mode');
    themeToggleBtn.textContent = '☀️';
    themeToggleBtn.setAttribute('title', '라이트 모드로 전환');
  }

  themeToggleBtn.addEventListener('click', () => {
    const isDark = document.body.classList.toggle('dark-mode');
    if (isDark) {
      themeToggleBtn.textContent = '☀️';
      themeToggleBtn.setAttribute('title', '라이트 모드로 전환');
      localStorage.setItem('ssu-dorm-theme', 'dark');
      showToast('다크 모드가 적용되었습니다 🌙');
    } else {
      themeToggleBtn.textContent = '🌙';
      themeToggleBtn.setAttribute('title', '다크 모드로 전환');
      localStorage.setItem('ssu-dorm-theme', 'light');
      showToast('라이트 모드가 적용되었습니다 ☀️');
    }
  });
}

/* ==========================================================================
   4. 세탁실 현황 (세탁기 & 건조기) 데이터 및 기능
   ========================================================================== */
// 세탁기 및 건조기 초기 상태 데이터
const machineData = [
  { id: 'washer-1', type: 'washer', name: '세탁기 1', status: 'available', remainingMinutes: 0 },
  { id: 'washer-2', type: 'washer', name: '세탁기 2', status: 'inuse', remainingMinutes: 23 },
  { id: 'washer-3', type: 'washer', name: '세탁기 3', status: 'broken', remainingMinutes: 0 },
  { id: 'dryer-1', type: 'dryer', name: '건조기 1', status: 'available', remainingMinutes: 0 },
  { id: 'dryer-2', type: 'dryer', name: '건조기 2', status: 'inuse', remainingMinutes: 15 },
  { id: 'dryer-3', type: 'dryer', name: '건조기 3', status: 'available', remainingMinutes: 0 }
];

// 현재 선택된 탭 필터 ('all' | 'washer' | 'dryer')
let currentMachineTab = 'all';

function initLaundrySection() {
  const container = document.getElementById('machine-list-container');
  const tabs = document.querySelectorAll('.machine-tab-btn');

  // 탭 클릭 이벤트 연결
  tabs.forEach(tab => {
    tab.addEventListener('click', (e) => {
      tabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      currentMachineTab = tab.getAttribute('data-tab') || 'all';
      renderMachines();
    });
  });

  // 초기 렌더링
  renderMachines();

  // 1분(60초)마다 사용 중인 세탁기/건조기 시간 1분씩 감소하는 타이머
  setInterval(() => {
    let stateChanged = false;
    machineData.forEach(machine => {
      if (machine.status === 'inuse' && machine.remainingMinutes > 0) {
        machine.remainingMinutes -= 1;
        stateChanged = true;
        // 시간이 다 되면 자동으로 사용 가능 상태로 변경
        if (machine.remainingMinutes <= 0) {
          machine.status = 'available';
          showToast(`🔔 ${machine.name} 사용이 완료되었습니다!`);
        }
      }
    });

    if (stateChanged) {
      renderMachines();
      updateSummaryStats();
    }
  }, 60000);
}

// 세탁기/건조기 목록 화면 렌더링
function renderMachines() {
  const container = document.getElementById('machine-list-container');
  if (!container) return;

  const filteredMachines = machineData.filter(machine => {
    if (currentMachineTab === 'all') return true;
    return machine.type === currentMachineTab;
  });

  container.innerHTML = '';

  filteredMachines.forEach(machine => {
    const card = document.createElement('div');
    card.className = 'machine-card';

    // 기기 아이콘 (세탁기: 🧺 / 건조기: ♨️)
    const icon = machine.type === 'washer' ? '🧺' : '♨️';

    // 상태에 따른 배지 및 텍스트 구성
    let badgeHtml = '';
    let actionBtnHtml = '';

    if (machine.status === 'available') {
      badgeHtml = `<span class="badge badge-available">● 사용 가능</span>`;
      actionBtnHtml = `<button class="btn btn-primary btn-sm btn-start-machine" data-id="${machine.id}">사용 시작</button>`;
    } else if (machine.status === 'inuse') {
      badgeHtml = `<span class="badge badge-inuse">● 사용 중</span>`;
      actionBtnHtml = `
        <div style="display:flex; flex-direction:column; align-items:flex-end; gap:4px;">
          <span class="time-remaining" style="font-size:0.8rem;">남은 시간: ${machine.remainingMinutes}분</span>
          <button class="btn btn-secondary btn-sm btn-finish-machine" data-id="${machine.id}" style="font-size:0.75rem; padding:4px 8px;">완료하기</button>
        </div>`;
    } else {
      badgeHtml = `<span class="badge badge-broken">● 고장</span>`;
      actionBtnHtml = `<span style="font-size:0.78rem; color:var(--status-broken-text); font-weight:600;">점검 중</span>`;
    }

    card.innerHTML = `
      <div class="machine-info">
        <div class="machine-icon-wrapper">${icon}</div>
        <div>
          <div class="machine-name">${machine.name}</div>
          <div class="machine-detail">
            ${badgeHtml}
          </div>
        </div>
      </div>
      <div>
        ${actionBtnHtml}
      </div>
    `;

    container.appendChild(card);
  });

  // '사용 시작' 버튼 이벤트 바인딩
  container.querySelectorAll('.btn-start-machine').forEach(btn => {
    btn.addEventListener('click', () => {
      const machineId = btn.getAttribute('data-id');
      startMachineUsage(machineId);
    });
  });

  // '완료하기' 버튼 이벤트 바인딩
  container.querySelectorAll('.btn-finish-machine').forEach(btn => {
    btn.addEventListener('click', () => {
      const machineId = btn.getAttribute('data-id');
      finishMachineUsage(machineId);
    });
  });
}

// 세탁기/건조기 사용 시작 처리
function startMachineUsage(machineId) {
  const target = machineData.find(m => m.id === machineId);
  if (!target) return;

  // 상태를 "사용 중"으로 변경하고 기본 시간 40분 설정
  target.status = 'inuse';
  target.remainingMinutes = 40;

  renderMachines();
  updateSummaryStats();
  showToast(`${target.name} 사용을 시작했습니다. (40분 설정)`);
}

// 세탁기/건조기 사용 완료 처리
function finishMachineUsage(machineId) {
  const target = machineData.find(m => m.id === machineId);
  if (!target) return;

  target.status = 'available';
  target.remainingMinutes = 0;

  renderMachines();
  updateSummaryStats();
  showToast(`${target.name} 사용이 완료 처리되었습니다.`);
}

/* ==========================================================================
   5. 택배 확인 데이터 및 기능
   ========================================================================== */
const deliveryData = [
  {
    id: 'pkg-1',
    carrier: 'CJ대한통운',
    item: '의류',
    time: '오후 2:30',
    status: 'pending' // pending: 수령 대기, received: 수령 완료
  },
  {
    id: 'pkg-2',
    carrier: '한진택배',
    item: '생활용품',
    time: '오전 11:20',
    status: 'received'
  },
  {
    id: 'pkg-3',
    carrier: '우체국택배',
    item: '전공서적 (컴퓨터학부)',
    time: '오후 4:10',
    status: 'pending'
  },
  {
    id: 'pkg-4',
    carrier: '로젠택배',
    item: '간식 세트',
    time: '어제 오후 5:40',
    status: 'received'
  }
];

function initDeliverySection() {
  renderDeliveries();
}

function renderDeliveries() {
  const container = document.getElementById('delivery-list-container');
  if (!container) return;

  container.innerHTML = '';

  deliveryData.forEach(delivery => {
    const card = document.createElement('div');
    card.className = 'delivery-card';

    const isPending = delivery.status === 'pending';
    const statusBadge = isPending
      ? `<span class="badge badge-inuse">수령 대기</span>`
      : `<span class="badge badge-available">수령 완료</span>`;

    const actionButton = isPending
      ? `<button class="btn btn-primary btn-sm btn-receive-pkg" data-id="${delivery.id}">수령 완료</button>`
      : `<span style="font-size:0.78rem; color:var(--color-text-muted);">수령 완료됨 ✓</span>`;

    card.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:2px;">
        <div class="delivery-company">
          <span>📦</span>
          <span>${delivery.carrier}</span>
        </div>
        <div class="delivery-meta">
          <span><strong>상품:</strong> ${delivery.item}</span>
          <span><strong>도착 시간:</strong> ${delivery.time}</span>
        </div>
      </div>
      <div class="delivery-action">
        ${statusBadge}
        ${actionButton}
      </div>
    `;

    container.appendChild(card);
  });

  // "수령 완료" 버튼 클릭 이벤트 바인딩
  container.querySelectorAll('.btn-receive-pkg').forEach(btn => {
    btn.addEventListener('click', () => {
      const pkgId = btn.getAttribute('data-id');
      markPackageReceived(pkgId);
    });
  });
}

function markPackageReceived(pkgId) {
  const target = deliveryData.find(d => d.id === pkgId);
  if (!target) return;

  // 상태를 '수령 완료'로 변경
  target.status = 'received';
  renderDeliveries();
  updateSummaryStats();
  showToast(`[${target.carrier}] 택배 수령이 완료되었습니다 🎉`);
}

/* ==========================================================================
   6. 시설 고장 신고 폼 유효성 검사 및 접수 기능
   ========================================================================== */
function initRepairSection() {
  const repairForm = document.getElementById('repair-form');
  const resultBox = document.getElementById('repair-success-msg');

  if (!repairForm) return;

  repairForm.addEventListener('submit', (e) => {
    e.preventDefault();

    // 1. 이름 또는 닉네임 입력 확인
    const nameInput = document.getElementById('reporter-name');
    const nameVal = nameInput ? nameInput.value.trim() : '';

    if (!nameVal) {
      alert('이름 또는 닉네임을 입력해주세요');
      if (nameInput) nameInput.focus();
      return;
    }

    // 2. 위치 입력 확인
    const locationInput = document.getElementById('repair-location');
    const locationVal = locationInput ? locationInput.value.trim() : '';

    if (!locationVal) {
      alert('문제가 발생한 위치를 입력해주세요');
      if (locationInput) locationInput.focus();
      return;
    }

    // 3. 고장 종류 선택 확인
    const categorySelect = document.getElementById('repair-category');
    const categoryVal = categorySelect ? categorySelect.value : '';

    if (!categoryVal) {
      alert('고장 종류를 선택해주세요');
      if (categorySelect) categorySelect.focus();
      return;
    }

    // 4. 긴급도 라디오 버튼 선택 값 가져오기
    const urgencyRadio = document.querySelector('input[name="urgency"]:checked');
    const urgencyVal = urgencyRadio ? urgencyRadio.value : '보통';

    // 5. 상세 내용 가져오기
    const detailInput = document.getElementById('repair-details');
    const detailVal = detailInput ? detailInput.value.trim() : '';

    // 6. 실시간 엑셀 대장 데이터 객체 생성 및 저장
    const now = new Date();
    const formattedDate = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    const newReportItem = {
      id: `SSU-${Date.now().toString().slice(-6)}`,
      timestamp: formattedDate,
      name: nameVal,
      location: locationVal,
      category: categoryVal,
      urgency: urgencyVal,
      details: detailVal || '(상세 내용 없음)',
      status: '접수 완료'
    };

    // 엑셀 대장 데이터에 신규 행 추가
    addReportToExcel(newReportItem);

    // 접수 완료 메시지 표시
    // 예: "김민수님, 5층 세탁실의 세탁기 고장 신고가 접수되었습니다."
    if (resultBox) {
      resultBox.innerHTML = `
        <div style="font-weight:700; margin-bottom:4px; font-size:0.92rem;">✅ 신고 접수 및 엑셀 대장 등록 완료</div>
        <div><strong>${escapeHtml(nameVal)}</strong>님, <strong>${escapeHtml(locationVal)}</strong>의 <strong>${escapeHtml(categoryVal)}</strong> 고장 신고(긴급도: ${escapeHtml(urgencyVal)})가 정상 접수되었습니다.</div>
        <div style="font-size:0.78rem; opacity:0.85; margin-top:6px;">📊 하단 <strong>실시간 엑셀 접수 대장</strong>에 즉시 기록되었으며, <strong>[엑셀 다운로드]</strong> 버튼으로 언제든 내려받을 수 있습니다.</div>
      `;
      resultBox.classList.add('show');
    }

    showToast('📊 엑셀 접수 대장에 1건이 추가되었습니다!');

    // 폼 초기화 (단, 긴급도는 '보통' 기본 유지)
    repairForm.reset();
    const defaultUrgency = document.getElementById('urgency-normal');
    if (defaultUrgency) defaultUrgency.checked = true;

    // 접수 결과 박스로 부드럽게 스크롤
    if (resultBox) {
      resultBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  });

  // 다시 작성(리셋) 버튼 처리
  const resetBtn = document.getElementById('btn-reset-repair');
  if (resetBtn) {
    resetBtn.addEventListener('click', () => {
      if (resultBox) {
        resultBox.classList.remove('show');
        resultBox.innerHTML = '';
      }
    });
  }
}

/* ==========================================================================
   6-1. 엑셀(Excel) 실시간 데이터 연동 및 내보내기 모듈
   ========================================================================== */
// 로컬 스토리지 키
const EXCEL_STORAGE_KEY = 'ssu-dorm-excel-reports';

// 초기 샘플 데이터 (첫 화면에서도 엑셀 대장 동작을 직관적으로 확인 가능)
const defaultReports = [
  {
    id: 'SSU-260901',
    timestamp: '2026-09-28 14:20',
    name: '김민수',
    location: '5층 세탁실',
    category: '세탁기',
    urgency: '보통',
    details: '2번 세탁기 탈수 시 소음 발생',
    status: '처리 완료'
  },
  {
    id: 'SSU-260902',
    timestamp: '2026-09-30 09:15',
    name: '이수진',
    location: '304호',
    category: '전등',
    urgency: '낮음',
    details: '책상 위 형광등 깜빡임 증상',
    status: '접수 완료'
  }
];

let repairReportsList = [];

// 로컬스토리지에서 엑셀 신고 내역 로드
function loadReportsFromStorage() {
  try {
    const saved = localStorage.getItem(EXCEL_STORAGE_KEY);
    if (saved) {
      repairReportsList = JSON.parse(saved);
    } else {
      repairReportsList = [...defaultReports];
      saveReportsToStorage();
    }
  } catch (e) {
    repairReportsList = [...defaultReports];
  }
}

// 로컬스토리지에 엑셀 신고 내역 저장
function saveReportsToStorage() {
  try {
    localStorage.setItem(EXCEL_STORAGE_KEY, JSON.stringify(repairReportsList));
  } catch (e) {
    console.error('Failed to save to localStorage', e);
  }
}

// 엑셀 관리 섹션 초기화
function initExcelSection() {
  loadReportsFromStorage();
  renderExcelTable();

  // 1) 엑셀(.xlsx) 다운로드 버튼 이벤트
  const excelBtn = document.getElementById('btn-download-excel');
  if (excelBtn) {
    excelBtn.addEventListener('click', exportToExcel);
  }

  // 2) CSV 다운로드 버튼 이벤트
  const csvBtn = document.getElementById('btn-download-csv');
  if (csvBtn) {
    csvBtn.addEventListener('click', exportToCSV);
  }

  // 3) 테스트 샘플 추가 버튼
  const addSampleBtn = document.getElementById('btn-add-sample-report');
  if (addSampleBtn) {
    addSampleBtn.addEventListener('click', () => {
      const sampleNames = ['박준형', '최유진', '정현우', '강민지', '윤태호'];
      const sampleLocations = ['201호', '3층 복도', '405호', '지하 세탁실', '2층 화장실'];
      const sampleCats = ['에어컨', '수도', '난방', '인터넷/Wi-Fi', '화장실'];
      const sampleUrgencies = ['보통', '긴급', '낮음'];

      const randomIdx = Math.floor(Math.random() * sampleNames.length);
      const now = new Date();
      const timeStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

      const sampleItem = {
        id: `SSU-${Date.now().toString().slice(-6)}`,
        timestamp: timeStr,
        name: sampleNames[randomIdx],
        location: sampleLocations[randomIdx],
        category: sampleCats[randomIdx],
        urgency: sampleUrgencies[randomIdx % 3],
        details: `${sampleCats[randomIdx]} 점검 요청 드립니다.`,
        status: '접수 완료'
      };

      addReportToExcel(sampleItem);
      showToast(`[${sampleItem.name}] 님의 샘플 데이터가 엑셀 대장에 추가되었습니다.`);
    });
  }

  // 4) 대장 비우기 버튼
  const clearBtn = document.getElementById('btn-clear-reports');
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      if (confirm('엑셀 접수 대장의 모든 내역을 비우시겠습니까?')) {
        repairReportsList = [];
        saveReportsToStorage();
        renderExcelTable();
        showToast('엑셀 접수 대장이 초기화되었습니다.');
      }
    });
  }
}

// 엑셀 대장에 새 항목 추가
function addReportToExcel(reportItem) {
  repairReportsList.unshift(reportItem);
  saveReportsToStorage();
  renderExcelTable(reportItem.id);
}

// 엑셀 대장 화면 테이블 렌더링
function renderExcelTable(highlightId = null) {
  const tbody = document.getElementById('excel-table-body');
  const countBadge = document.getElementById('excel-count-badge');
  if (!tbody) return;

  if (countBadge) {
    countBadge.textContent = `${repairReportsList.length}건`;
  }

  if (repairReportsList.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" style="text-align:center; padding: 22px; color: var(--color-text-muted);">
          등록된 신고 내역이 없습니다. 위 폼에서 고장 신고를 작성하면 엑셀 대장에 실시간 기록됩니다.
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = '';

  repairReportsList.forEach((item, index) => {
    const tr = document.createElement('tr');
    if (highlightId && item.id === highlightId) {
      tr.className = 'new-row';
    }

    // 긴급도 뱃지
    let urgencyBadge = `<span style="color:var(--color-text-muted);">${escapeHtml(item.urgency)}</span>`;
    if (item.urgency === '긴급') {
      urgencyBadge = `<span style="color:var(--status-broken-text); font-weight:700;">🚨 ${escapeHtml(item.urgency)}</span>`;
    } else if (item.urgency === '보통') {
      urgencyBadge = `<span style="color:var(--status-inuse-text); font-weight:600;">${escapeHtml(item.urgency)}</span>`;
    }

    // 상태 뱃지
    const isDone = item.status === '처리 완료';
    const statusHtml = isDone
      ? `<span class="badge badge-available" style="font-size:0.72rem;">처리 완료</span>`
      : `<span class="badge badge-inuse" style="font-size:0.72rem;">접수 완료</span>`;

    tr.innerHTML = `
      <td style="text-align:center; font-weight:600; color:var(--color-text-muted);">${index + 1}</td>
      <td style="font-family: monospace; font-size:0.75rem;">${escapeHtml(item.timestamp)}</td>
      <td><strong>${escapeHtml(item.name)}</strong></td>
      <td>${escapeHtml(item.location)}</td>
      <td><span class="badge" style="background-color:var(--color-point-light); color:var(--color-point);">${escapeHtml(item.category)}</span></td>
      <td>${urgencyBadge}</td>
      <td style="max-width: 150px; overflow:hidden; text-overflow:ellipsis;" title="${escapeHtml(item.details)}">${escapeHtml(item.details)}</td>
      <td>${statusHtml}</td>
    `;

    tbody.appendChild(tr);
  });
}

// 엑셀(.xlsx) 파일 내보내기 함수 (SheetJS 활용)
function exportToExcel() {
  if (repairReportsList.length === 0) {
    alert('엑셀로 내보낼 데이터가 없습니다. 먼저 신고 내역을 입력해주세요.');
    return;
  }

  // 1. 시트 1: 시설 고장 신고 대장 데이터 포맷팅
  const reportRows = repairReportsList.map((item, index) => ({
    '연번': index + 1,
    '접수번호': item.id,
    '접수일시': item.timestamp,
    '신고자': item.name,
    '발생위치': item.location,
    '고장종류': item.category,
    '긴급도': item.urgency,
    '상세내용': item.details,
    '처리상태': item.status
  }));

  const wb = XLSX.utils.book_new();
  const wsReports = XLSX.utils.json_to_sheet(reportRows);

  // 열 너비(Column width) 보기 편하게 최적화 설정
  wsReports['!cols'] = [
    { wch: 6 },  // 연번
    { wch: 14 }, // 접수번호
    { wch: 18 }, // 접수일시
    { wch: 12 }, // 신고자
    { wch: 18 }, // 발생위치
    { wch: 14 }, // 고장종류
    { wch: 10 }, // 긴급도
    { wch: 32 }, // 상세내용
    { wch: 12 }  // 처리상태
  ];

  XLSX.utils.book_append_sheet(wb, wsReports, '시설고장신고대장');

  // 2. 시트 2: 공용 물품 대여 현황도 함께 워크북에 추가
  if (typeof rentalItemsData !== 'undefined' && rentalItemsData.length > 0) {
    const rentalRows = rentalItemsData.map((item, index) => ({
      '연번': index + 1,
      '물품명': item.name,
      '현재상태': item.status === 'available' ? '대여 가능' : '대여 중',
      '보관위치': '생활관 1층 행정실 / 로비',
      '비고': '사생증 확인 후 대여'
    }));
    const wsRentals = XLSX.utils.json_to_sheet(rentalRows);
    wsRentals['!cols'] = [{ wch: 6 }, { wch: 16 }, { wch: 14 }, { wch: 28 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, wsRentals, '공용물품현황');
  }

  // 3. 실제 엑셀(.xlsx) 파일 생성 및 브라우저 다운로드
  const todayStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const fileName = `숭실대_기숙사_접수대장_${todayStr}.xlsx`;
  XLSX.writeFile(wb, fileName);

  showToast('📥 엑셀(.xlsx) 파일이 다운로드되었습니다!');
}

// CSV 파일 내보내기 함수 (엑셀 한글 깨짐 방지를 위한 UTF-8 with BOM 적용)
function exportToCSV() {
  if (repairReportsList.length === 0) {
    alert('내보낼 데이터가 없습니다.');
    return;
  }

  const headers = ['연번', '접수번호', '접수일시', '신고자', '발생위치', '고장종류', '긴급도', '상세내용', '처리상태'];
  const rows = repairReportsList.map((item, index) => [
    index + 1,
    `"${item.id}"`,
    `"${item.timestamp}"`,
    `"${item.name.replace(/"/g, '""')}"`,
    `"${item.location.replace(/"/g, '""')}"`,
    `"${item.category}"`,
    `"${item.urgency}"`,
    `"${item.details.replace(/"/g, '""')}"`,
    `"${item.status}"`
  ]);

  const csvContent = '\uFEFF' + [headers.join(','), ...rows.map(r => r.join(','))].join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  const todayStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  link.setAttribute('download', `숭실대_기숙사_신고대장_${todayStr}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);

  showToast('📄 CSV 파일이 다운로드되었습니다!');
}

/* ==========================================================================
   7. 공용 물품 대여 기능
   ========================================================================== */
const rentalItemsData = [
  { id: 'item-1', name: '공구 세트', icon: '🔨', status: 'available' },
  { id: 'item-2', name: '청소기', icon: '🧹', status: 'available' },
  { id: 'item-3', name: '보드게임', icon: '♟️', status: 'inuse' },
  { id: 'item-4', name: '우산', icon: '☂️', status: 'available' },
  { id: 'item-5', name: '멀티탭', icon: '🔌', status: 'available' }
];

function initRentalSection() {
  renderRentalItems();
}

function renderRentalItems() {
  const container = document.getElementById('rental-list-container');
  if (!container) return;

  container.innerHTML = '';

  rentalItemsData.forEach(item => {
    const card = document.createElement('div');
    card.className = 'rental-card';

    const isAvailable = item.status === 'available';

    // 상태 배지
    const badgeHtml = isAvailable
      ? `<span class="badge badge-available">● 대여 가능</span>`
      : `<span class="badge badge-inuse">● 대여 중</span>`;

    // 액션 버튼
    const actionBtnHtml = isAvailable
      ? `<button class="btn btn-primary btn-sm btn-rent" data-id="${item.id}" data-name="${item.name}">대여하기</button>`
      : `<button class="btn btn-secondary btn-sm btn-return" data-id="${item.id}" data-name="${item.name}">반납하기</button>`;

    card.innerHTML = `
      <div class="rental-info">
        <div class="rental-icon">${item.icon}</div>
        <div>
          <div class="rental-name">${item.name}</div>
          <div>${badgeHtml}</div>
        </div>
      </div>
      <div>
        ${actionBtnHtml}
      </div>
    `;

    container.appendChild(card);
  });

  // 대여하기 버튼 이벤트 바인딩
  container.querySelectorAll('.btn-rent').forEach(btn => {
    btn.addEventListener('click', () => {
      const itemId = btn.getAttribute('data-id');
      const itemName = btn.getAttribute('data-name');

      // confirm 창으로 대여 확인
      // 예: "청소기를 대여하시겠습니까?"
      const isConfirmed = confirm(`${itemName}을(를) 대여하시겠습니까?`);
      if (isConfirmed) {
        const target = rentalItemsData.find(i => i.id === itemId);
        if (target) {
          target.status = 'inuse';
          renderRentalItems();
          showToast(`${itemName} 대여가 완료되었습니다!`);
        }
      }
    });
  });

  // 반납하기 버튼 이벤트 바인딩
  container.querySelectorAll('.btn-return').forEach(btn => {
    btn.addEventListener('click', () => {
      const itemId = btn.getAttribute('data-id');
      const itemName = btn.getAttribute('data-name');

      const isConfirmed = confirm(`${itemName}을(를) 반납하시겠습니까?`);
      if (isConfirmed) {
        const target = rentalItemsData.find(i => i.id === itemId);
        if (target) {
          target.status = 'available';
          renderRentalItems();
          showToast(`${itemName}이(가) 반납되었습니다.`);
        }
      }
    });
  });
}

/* ==========================================================================
   8. 상단 요약 정보 자동 계산 및 동적 업데이트
   ========================================================================== */
function updateSummaryStats() {
  // 1) 사용 가능한 세탁기 대수 계산 (washer 타입 중 status === 'available')
  const availableWashersCount = machineData.filter(
    m => m.type === 'washer' && m.status === 'available'
  ).length;

  const laundrySummaryEl = document.getElementById('summary-laundry');
  if (laundrySummaryEl) {
    laundrySummaryEl.textContent = `사용 가능한 세탁기 ${availableWashersCount}대`;
  }

  // 2) 도착한 택배 개수 계산 (status === 'pending' 수령 대기 건수)
  const pendingDeliveriesCount = deliveryData.filter(
    d => d.status === 'pending'
  ).length;

  const deliverySummaryEl = document.getElementById('summary-delivery');
  if (deliverySummaryEl) {
    deliverySummaryEl.textContent = `도착한 택배 ${pendingDeliveriesCount}개`;
  }

  // 3) 새로운 공지 개수 (중요 및 최근 공지)
  const noticeSummaryEl = document.getElementById('summary-notice');
  if (noticeSummaryEl) {
    noticeSummaryEl.textContent = `새로운 공지 1개`;
  }
}

/* ==========================================================================
   9. 하단 네비게이션 부드러운 스크롤 이동
   ========================================================================== */
function initNavigation() {
  const navItems = document.querySelectorAll('.nav-item');

  navItems.forEach(item => {
    item.addEventListener('click', (e) => {
      e.preventDefault();

      const targetId = item.getAttribute('href');
      if (!targetId) return;

      const targetElement = document.querySelector(targetId);
      if (targetElement) {
        // 활성화 스타일 적용
        navItems.forEach(n => n.classList.remove('active'));
        item.classList.add('active');

        // 스크롤 이동 (헤더 높이 고려)
        targetElement.scrollIntoView({
          behavior: 'smooth',
          block: 'start'
        });
      }
    });
  });

  // 스크롤 시 현재 위치에 맞는 네비게이션 항목 자동 하이라이트
  window.addEventListener('scroll', () => {
    const sections = ['#home', '#laundry', '#delivery', '#report', '#notice'];
    const scrollPosition = window.scrollY + 120;

    for (let i = sections.length - 1; i >= 0; i--) {
      const section = document.querySelector(sections[i]);
      if (section && section.offsetTop <= scrollPosition) {
        navItems.forEach(item => {
          if (item.getAttribute('href') === sections[i]) {
            item.classList.add('active');
          } else {
            item.classList.remove('active');
          }
        });
        break;
      }
    }
  });
}

/* ==========================================================================
   10. 유틸리티 함수 (토스트 알림 및 XSS 방지)
   ========================================================================== */
let toastTimeout;
function showToast(message) {
  let toast = document.getElementById('app-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'app-toast';
    toast.className = 'toast';
    document.body.appendChild(toast);
  }

  toast.textContent = message;
  toast.classList.add('show');

  clearTimeout(toastTimeout);
  toastTimeout = setTimeout(() => {
    toast.classList.remove('show');
  }, 2500);
}

function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return String(text).replace(/[&<>"']/g, m => map[m]);
}
