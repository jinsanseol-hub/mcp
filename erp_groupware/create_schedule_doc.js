const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak
} = require('docx');
const fs = require('fs');

// ────────────────────────────────────────────────
// 공통 테두리 / 스타일 헬퍼
// ────────────────────────────────────────────────
const border = { style: BorderStyle.SINGLE, size: 1, color: 'B0C4DE' };
const borders = { top: border, bottom: border, left: border, right: border };
const headerBorder = { style: BorderStyle.SINGLE, size: 1, color: '2E5FA3' };
const headerBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };

const cellPad = { top: 100, bottom: 100, left: 140, right: 140 };

function hCell(text, width) {
  return new TableCell({
    borders: headerBorders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: '1F4E79', type: ShadingType.CLEAR },
    margins: cellPad,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, color: 'FFFFFF', size: 20, font: '맑은 고딕' })]
    })]
  });
}

function dCell(text, width, shade, align) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: shade || 'FFFFFF', type: ShadingType.CLEAR },
    margins: cellPad,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: align || AlignmentType.LEFT,
      children: [new TextRun({ text: text || '', size: 19, font: '맑은 고딕' })]
    })]
  });
}

function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 360, after: 120 },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 26,
        font: '맑은 고딕',
        color: '1F4E79'
      })
    ],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '2E5FA3', space: 4 } }
  });
}

function makeTable(cols, rows) {
  const colWidths = cols.map(c => c.width);
  const total = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: rows
  });
}

// ────────────────────────────────────────────────
// 데이터: 엔지니어별 이번 주(3.8~3.14) 일정
// ────────────────────────────────────────────────
const weeklyData = [
  {
    name: '배종인', team: '기술2팀',
    schedules: [
      { date: '3.9(월)', time: '09:00~13:00', customer: 'LG 디스플레이', project: 'M2603004', type: 'OPS', note: '스토리지 유지보수 정기점검' },
      { date: '3.9(월)', time: '14:00~18:00', customer: '중소벤처기업진흥공단', project: 'MPHS24002', type: 'OPS', note: '' },
      { date: '3.10(화)', time: '종일', customer: '대법원', project: 'M2512009', type: 'INC', note: '' },
      { date: '3.11(수)', time: '종일', customer: '한국사회보장정보원', project: 'M2602027', type: 'OPS', note: '' },
      { date: '3.12(목)', time: '종일', customer: '대법원', project: 'M2512009', type: 'INC', note: '' },
      { date: '3.13(금)', time: '종일', customer: '대법원', project: 'M2512009', type: 'INC', note: '' },
      { date: '3.14(토)', time: '22:00~23:59', customer: '대법원', project: 'MRHJ25005', type: 'CHG', note: '야간 변경작업' },
    ]
  },
  {
    name: '나두균', team: '기술2팀',
    schedules: [
      { date: '3.8(일)', time: '종일', customer: '금융결제원', project: 'MKHS24021', type: 'CHG', note: '' },
      { date: '3.9(월)', time: '종일', customer: '한국가스공사', project: 'MRHJ25008', type: 'CHG', note: '' },
      { date: '3.10(화)', time: '종일', customer: '한국가스공사', project: 'MRHJ25008', type: 'CHG', note: '' },
      { date: '3.12(목)', time: '종일', customer: '금융결제원', project: 'MKHS24021', type: 'OPS', note: '' },
      { date: '3.13(금)', time: '종일', customer: '아시아나항공', project: 'MPHS25001', type: 'OPS', note: '' },
    ]
  },
  {
    name: '권순현', team: '기술2팀',
    schedules: [
      { date: '3.9(월)', time: '10:00~18:00', customer: 'MUFG', project: 'M2510001', type: 'OPS', note: '' },
      { date: '3.11(수)', time: '10:00~18:00', customer: 'MUFG', project: 'M2510001', type: 'OPS', note: '' },
      { date: '3.12(목)', time: '10:00~18:00', customer: 'MUFG', project: 'M2510001', type: 'OPS', note: '' },
    ]
  },
  {
    name: '박영민', team: '기술2팀',
    schedules: [
      { date: '3.9(월)', time: '12:00~22:00', customer: 'SK대덕', project: 'VBON25749', type: 'INC', note: '' },
      { date: '3.11(수)', time: '종일', customer: '삼성디스플레이', project: 'VBON25549', type: 'INC', note: '' },
      { date: '3.12(목)', time: '14:00~18:00', customer: 'MUFG', project: 'M2510001', type: 'OPS', note: '' },
    ]
  },
  {
    name: '이승민', team: '기술2팀',
    schedules: [
      { date: '3.8(일)', time: '종일', customer: '금융결제원', project: 'MKHS24021', type: 'CHG', note: '' },
      { date: '3.10(화)', time: '종일', customer: '아시아나항공', project: 'C2601006', type: 'OPS', note: '' },
      { date: '3.12(목)', time: '종일', customer: '금융결제원', project: 'MKHS24021', type: 'OPS', note: '' },
    ]
  },
  {
    name: '김성은', team: '기술2팀',
    schedules: [
      { date: '3.10(화)', time: '종일', customer: '아시아나항공', project: 'C2601006', type: 'OPS', note: '' },
      { date: '3.12(목)', time: '종일', customer: '아시아나항공', project: 'C2601006', type: 'OPS', note: '' },
      { date: '3.13(금)', time: '13:00~18:00', customer: 'MUFG', project: 'M2510001', type: 'OPS', note: '' },
    ]
  },
  {
    name: '정윤지', team: '기술2팀',
    schedules: [
      { date: '3.10(화)', time: '종일', customer: '아시아나항공', project: 'C2601006', type: 'OPS', note: '' },
      { date: '3.12(목)', time: '종일', customer: '아시아나항공', project: 'C2601006', type: 'OPS', note: '' },
    ]
  },
  {
    name: '김우현', team: '기술2팀',
    schedules: [
      { date: '3.10(화)', time: '종일', customer: '한국가스공사', project: 'MRHJ24012', type: 'INC', note: '' },
      { date: '3.11(수)', time: '종일', customer: '근로복지공단', project: 'PRHJ24004', type: 'OPS', note: '' },
      { date: '3.13(금)', time: '종일', customer: '국민연금공단', project: 'MKHS25006', type: 'INC', note: '' },
    ]
  },
  {
    name: '최은준', team: '기술2팀',
    schedules: [
      { date: '3.9(월)', time: '종일', customer: 'OT 휴무', project: '-', type: '-', note: '' },
      { date: '3.11(수)', time: '종일', customer: 'SK대전(MVS)', project: 'mvs', type: 'OPS', note: '' },
      { date: '3.12(목)', time: '14:00~18:00', customer: '기상청', project: 'm2601014', type: 'OPS', note: '' },
      { date: '3.13(금)', time: '18:00~23:59', customer: 'DS투자증권', project: 'M2510003', type: 'CHG', note: '야간 변경작업' },
      { date: '3.14(토)', time: '00:00~02:00', customer: 'DS투자증권', project: 'M2510003', type: 'CHG', note: '' },
    ]
  },
  {
    name: '오태석', team: '기술2팀',
    schedules: [
      { date: '3.9(월)', time: '14:00~18:00', customer: 'SK증권', project: 'MKHS25001', type: 'OPS', note: '' },
    ]
  },
  {
    name: '이홍빈', team: '기술2팀',
    schedules: [
      { date: '3.12(목)', time: '종일', customer: '예비군훈련', project: '-', type: '공가', note: '' },
    ]
  },
  {
    name: '최효정', team: '기술1팀',
    schedules: [
      { date: '3.10(화)~3.14(토)', time: '종일', customer: 'ESAZ (삼성SDS)', project: 'PKMH25011', type: 'DEP', note: '참여: 김의성, 최범락, 최효정' },
    ]
  },
  {
    name: '강찬혁', team: '기술3팀',
    schedules: [
      { date: '3.10(화)~3.14(토)', time: '종일', customer: 'ESAZ (삼성SDS)', project: 'PKMH25013', type: 'PRJ', note: '참여: 한장희, 강찬혁' },
    ]
  },
  {
    name: '지근영', team: '기술3팀',
    schedules: [
      { date: '3.10(화)', time: '종일', customer: 'IBK기업은행', project: 'none/INC', type: 'INC', note: '참여: 지근영, 김상민' },
    ]
  },
  {
    name: '이정인', team: '기술1팀',
    schedules: [
      { date: '3.9(월)', time: '09:00~14:00', customer: 'ESALL (SBON)', project: 'SBON_TO25_ALL', type: 'OPS', note: '' },
    ]
  },
  {
    name: '이영석', team: '기술1팀',
    schedules: [
      { date: '3.11(수)', time: '08:00~18:00', customer: 'MBC', project: 'P2601039', type: 'BTS', note: '' },
    ]
  },
];

// 중복지원 현황
const multiEngineerData = [
  { customer: '아시아나항공 (C2601006)', engineers: '이승민, 정윤지, 김성은', count: '3명', days: '3.10(화), 3.12(목)' },
  { customer: 'MUFG (M2510001)', engineers: '권순현, 박영민, 김성은', count: '3명', days: '3.9, 3.11, 3.12, 3.13' },
  { customer: '금융결제원 (MKHS24021)', engineers: '나두균, 이승민', count: '2명', days: '3.8(일), 3.12(목)' },
  { customer: '한국가스공사', engineers: '나두균, 김우현', count: '2명', days: '3.9~3.10, 3.10' },
  { customer: '대법원 (M2512009/MRHJ25005)', engineers: '배종인 집중 지원', count: '1명', days: '3.10~3.14' },
];

// ────────────────────────────────────────────────
// 문서 생성
// ────────────────────────────────────────────────
const children = [];

// ── 제목 ──
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 120 },
    children: [new TextRun({
      text: '엔지니어별 고객사 기술지원 현황',
      bold: true, size: 40, font: '맑은 고딕', color: '1F4E79'
    })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({
      text: '(주)본정보 기술지원본부 기술2팀',
      size: 24, font: '맑은 고딕', color: '595959'
    })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 400 },
    children: [new TextRun({
      text: '조회기간: 2026년 3월 8일(일) ~ 3월 14일(토)',
      size: 22, font: '맑은 고딕', color: '595959', italics: true
    })]
  })
);

// ── 섹션 1: 엔지니어별 상세 일정 ──
children.push(sectionTitle('1. 엔지니어별 주간 지원 고객사 현황'));
children.push(new Paragraph({ spacing: { before: 0, after: 160 }, children: [new TextRun('')] }));

// 컬럼 너비 (총 9360 DXA = A4 여백 2.5cm)
const COL = [900, 1200, 1380, 1960, 1380, 900, 1640];
// 날짜 | 시간 | 고객사 | 프로젝트코드 | 업무유형 | 비고

for (const eng of weeklyData) {
  // 엔지니어 이름 헤더 행
  children.push(
    new Paragraph({
      spacing: { before: 280, after: 60 },
      children: [
        new TextRun({ text: `▶  ${eng.name}`, bold: true, size: 22, font: '맑은 고딕', color: '1F4E79' }),
        new TextRun({ text: `  (${eng.team})`, size: 20, font: '맑은 고딕', color: '595959' }),
      ]
    })
  );

  const tableRows = [
    // 헤더
    new TableRow({
      tableHeader: true,
      children: [
        hCell('날짜', COL[0]),
        hCell('시간', COL[1]),
        hCell('고객사', COL[2]),
        hCell('프로젝트코드', COL[3]),
        hCell('업무유형', COL[4]),
        hCell('비고', COL[5] + COL[6]),
      ]
    }),
    ...eng.schedules.map((s, i) => {
      const bg = i % 2 === 0 ? 'F0F4FA' : 'FFFFFF';
      // 업무유형 색
      const typeColor =
        s.type === 'INC' ? 'C00000' :
        s.type === 'CHG' ? 'E36C09' :
        s.type === 'OPS' ? '375623' :
        s.type === 'PRJ' ? '17375E' :
        s.type === 'DEP' ? '17375E' : '404040';

      return new TableRow({
        children: [
          dCell(s.date, COL[0], bg, AlignmentType.CENTER),
          dCell(s.time, COL[1], bg, AlignmentType.CENTER),
          new TableCell({
            borders,
            width: { size: COL[2], type: WidthType.DXA },
            shading: { fill: bg, type: ShadingType.CLEAR },
            margins: cellPad,
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              children: [new TextRun({ text: s.customer, size: 19, font: '맑은 고딕', bold: true })]
            })]
          }),
          dCell(s.project, COL[3], bg, AlignmentType.CENTER),
          new TableCell({
            borders,
            width: { size: COL[4], type: WidthType.DXA },
            shading: { fill: bg, type: ShadingType.CLEAR },
            margins: cellPad,
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: s.type, size: 19, font: '맑은 고딕', bold: true, color: typeColor })]
            })]
          }),
          dCell(s.note, COL[5] + COL[6], bg),
        ]
      });
    })
  ];

  children.push(makeTable(
    COL.slice(0, 5).concat([{ width: COL[5] + COL[6] }]).map((c, i) => ({ width: typeof c === 'object' ? c.width : c })),
    tableRows
  ));
}

// 페이지 나누기
children.push(new Paragraph({ children: [new PageBreak()] }));

// ── 섹션 2: 고객사별 담당 엔지니어 현황 ──
children.push(sectionTitle('2. 고객사별 담당 엔지니어 현황'));
children.push(new Paragraph({ spacing: { before: 0, after: 160 }, children: [new TextRun('')] }));

// 고객사 → 엔지니어 맵 생성
const customerMap = {};
for (const eng of weeklyData) {
  for (const s of eng.schedules) {
    if (s.type === '-' || s.type === '공가') continue;
    const key = s.customer;
    if (!customerMap[key]) customerMap[key] = { engineers: new Set(), types: new Set(), dates: [] };
    customerMap[key].engineers.add(`${eng.name}(${eng.team})`);
    customerMap[key].types.add(s.type);
    customerMap[key].dates.push(s.date);
  }
}

const custCols = [2400, 3200, 1560, 2200];
const custTotal = custCols.reduce((a, b) => a + b, 0);

const custRows = [
  new TableRow({
    tableHeader: true,
    children: [
      hCell('고객사', custCols[0]),
      hCell('담당 엔지니어', custCols[1]),
      hCell('업무유형', custCols[2]),
      hCell('지원일', custCols[3]),
    ]
  }),
  ...Object.entries(customerMap).map(([cust, val], i) => {
    const bg = i % 2 === 0 ? 'F0F4FA' : 'FFFFFF';
    return new TableRow({
      children: [
        new TableCell({
          borders,
          width: { size: custCols[0], type: WidthType.DXA },
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: cellPad,
          children: [new Paragraph({ children: [new TextRun({ text: cust, bold: true, size: 19, font: '맑은 고딕' })] })]
        }),
        dCell([...val.engineers].join(', '), custCols[1], bg),
        dCell([...val.types].join(', '), custCols[2], bg, AlignmentType.CENTER),
        dCell([...new Set(val.dates)].join(', '), custCols[3], bg),
      ]
    });
  })
];

children.push(new Table({
  width: { size: custTotal, type: WidthType.DXA },
  columnWidths: custCols,
  rows: custRows
}));

children.push(new Paragraph({ spacing: { before: 400, after: 200 }, children: [new TextRun('')] }));

// ── 섹션 3: 다중 엔지니어 지원 현황 ──
children.push(sectionTitle('3. 복수 엔지니어 동시 지원 고객사'));
children.push(new Paragraph({ spacing: { before: 0, after: 160 }, children: [new TextRun('')] }));

const multiCols = [2800, 2800, 800, 2160];
const multiTotal = multiCols.reduce((a, b) => a + b, 0);

const multiRows = [
  new TableRow({
    tableHeader: true,
    children: [
      hCell('고객사', multiCols[0]),
      hCell('담당 엔지니어', multiCols[1]),
      hCell('인원', multiCols[2]),
      hCell('지원일정', multiCols[3]),
    ]
  }),
  ...multiEngineerData.map((m, i) => {
    const bg = i % 2 === 0 ? 'FFF2CC' : 'FFFFFF';
    return new TableRow({
      children: [
        new TableCell({
          borders, width: { size: multiCols[0], type: WidthType.DXA },
          shading: { fill: bg, type: ShadingType.CLEAR }, margins: cellPad,
          children: [new Paragraph({ children: [new TextRun({ text: m.customer, bold: true, size: 19, font: '맑은 고딕' })] })]
        }),
        dCell(m.engineers, multiCols[1], bg),
        new TableCell({
          borders, width: { size: multiCols[2], type: WidthType.DXA },
          shading: { fill: bg, type: ShadingType.CLEAR }, margins: cellPad,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: m.count, bold: true, size: 19, font: '맑은 고딕', color: 'C00000' })] })]
        }),
        dCell(m.days, multiCols[3], bg),
      ]
    });
  })
];

children.push(new Table({
  width: { size: multiTotal, type: WidthType.DXA },
  columnWidths: multiCols,
  rows: multiRows
}));

// ── 업무유형 범례 ──
children.push(new Paragraph({ spacing: { before: 480, after: 120 }, children: [new TextRun('')] }));
children.push(sectionTitle('4. 업무유형 범례'));
children.push(new Paragraph({ spacing: { before: 0, after: 160 }, children: [new TextRun('')] }));

const legendData = [
  { type: 'OPS', desc: '운영지원 (정기점검, 모니터링)', color: '375623' },
  { type: 'INC', desc: '장애 대응 (Incident)', color: 'C00000' },
  { type: 'CHG', desc: '변경 작업 (Change)', color: 'E36C09' },
  { type: 'PRJ', desc: '프로젝트 수행', color: '17375E' },
  { type: 'DEP', desc: '배포 작업 (Deployment)', color: '17375E' },
  { type: 'BTS', desc: 'BTS 기술지원', color: '595959' },
];

const legCols = [1200, 7760];
const legTotal = legCols.reduce((a, b) => a + b, 0);

const legRows = [
  new TableRow({
    tableHeader: true,
    children: [hCell('유형', legCols[0]), hCell('설명', legCols[1])]
  }),
  ...legendData.map((l, i) => {
    const bg = i % 2 === 0 ? 'F5F5F5' : 'FFFFFF';
    return new TableRow({
      children: [
        new TableCell({
          borders, width: { size: legCols[0], type: WidthType.DXA },
          shading: { fill: bg, type: ShadingType.CLEAR }, margins: cellPad,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: l.type, bold: true, size: 20, font: '맑은 고딕', color: l.color })] })]
        }),
        dCell(l.desc, legCols[1], bg)
      ]
    });
  })
];

children.push(new Table({
  width: { size: legTotal, type: WidthType.DXA },
  columnWidths: legCols,
  rows: legRows
}));

// ────────────────────────────────────────────────
// 문서 조립
// ────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: '맑은 고딕', size: 22 } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } // 2cm
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '2E5FA3', space: 4 } },
          children: [new TextRun({ text: '(주)본정보 기술지원본부  |  엔지니어별 고객사 기술지원 현황  |  2026.03', size: 18, font: '맑은 고딕', color: '595959' })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'B0C4DE', space: 4 } },
          children: [
            new TextRun({ text: '(주)본정보  |  대외비  |  ', size: 18, font: '맑은 고딕', color: '595959' }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: '맑은 고딕', color: '595959' }),
            new TextRun({ text: ' / ', size: 18, font: '맑은 고딕', color: '595959' }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: '맑은 고딕', color: '595959' }),
          ]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  const outPath = 'C:\\mcp\\erp_groupware\\엔지니어별_고객사_기술지원현황_202603.docx';
  fs.writeFileSync(outPath, buf);
  console.log('생성 완료:', outPath);
});
