/**
 * SCREAM ROOC 2026 — ปุ่มจัดตี้อัตโนมัติ
 *
 * วิธีติดตั้ง: ส่วนขยาย > Apps Script > ลบโค้ดเดิม > วางไฟล์นี้ทั้งหมด > บันทึก
 * จากนั้นรีเฟรชชีต จะมีเมนู "⚔️ จัดตี้" ขึ้นบนแถบเมนู
 * (ถ้าอยากได้ปุ่มบนชีต: แทรก > ภาพวาด > วาดปุ่ม > คลิกขวาที่ปุ่ม > ⋮ > กำหนดสคริปต์ > openPartyImporter)
 */

const PARTY_SHEET_NAME = 'รายชื่อตี้';   // ชื่อแท็บที่มีตารางปาตี้
const ROWS_PER_PARTY = 5;
const PARTIES_PER_TEAM = 8;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚔️ จัดตี้')
    .addItem('วางรายชื่อ → จัดตี้อัตโนมัติ', 'openPartyImporter')
    .addToUi();
}

function openPartyImporter() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:sans-serif;margin:0;padding:12px}
      textarea{width:100%;height:380px;box-sizing:border-box;font-family:monospace;font-size:12px}
      button{margin-top:8px;padding:8px 16px;background:#5b3c99;color:#fff;border:0;border-radius:6px;cursor:pointer}
      #msg{margin-top:8px;white-space:pre-wrap;font-size:12px}
    </style>
    <div>วางรายชื่อตี้ (ฟอร์แมต TEAM 1 / TEAM 2 → Party → 1. อาชีพ - ชื่อ)</div>
    <textarea id="txt"></textarea>
    <button id="btn" onclick="go()">จัดตี้เลย</button>
    <div id="msg"></div>
    <script>
      function go(){
        const btn=document.getElementById('btn'), msg=document.getElementById('msg');
        btn.disabled=true; msg.textContent='กำลังจัดตี้...';
        google.script.run
          .withSuccessHandler(r=>{msg.textContent=r;btn.disabled=false;})
          .withFailureHandler(e=>{msg.textContent='❌ '+e.message;btn.disabled=false;})
          .fillPartiesFromText(document.getElementById('txt').value);
      }
    </script>`).setWidth(520).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'จัดตี้อัตโนมัติ');
}

/** แปลงข้อความรายชื่อ → [{slot, run, remark, members:[{job,name}]}] */
function parsePartyText(text) {
  const parties = [];
  let team = 0, teamRun = '', teamRemark = '', current = null;

  text.split(/\r?\n/).forEach(raw => {
    const line = raw.trim();
    if (!line) return;

    const teamMatch = line.match(/TEAM\s*(\d)/i);
    if (teamMatch) {
      team = Number(teamMatch[1]);
      const paren = (line.match(/\(([^)]*)\)/) || [])[1] || '';
      teamRun = (paren.match(/RUN\s*\d+/i) || [`RUN ${team}`])[0].toUpperCase().replace(/RUN\s*/, 'RUN ');
      teamRemark = paren.split('/').map(s => s.trim()).find(s => s && !/^RUN/i.test(s)) || '';
      current = null;
      return;
    }

    const partyMatch = line.match(/^Party\s*(\d+)\s*(?:\(([^)]*)\))?/i);
    if (partyMatch) {
      if (!team) throw new Error(`เจอ "${line}" ก่อนหัวข้อ TEAM 1 / TEAM 2`);
      let n = Number(partyMatch[1]);
      if (team === 2 && n <= PARTIES_PER_TEAM) n += PARTIES_PER_TEAM; // TEAM 2 Party 1 = ตี้ล่างตารางที่ 9
      current = { slot: n, run: teamRun, remark: (partyMatch[2] || teamRemark).trim(), members: [] };
      parties.push(current);
      return;
    }

    const memberMatch = line.match(/^(\d+)[.)]\s*(.+?)\s+-\s+(.+)$/);
    if (memberMatch && current) {
      current.members.push({ job: memberMatch[2].trim(), name: memberMatch[3].trim() });
    }
  });
  return parties;
}

/** หาตารางปาตี้ทั้งหมดจากหัวคอลัมน์ "MEMBER" เรียงจากบนลงล่าง ซ้ายไปขวา */
function findPartyTables(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const tables = [];
  values.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (cell.trim().toUpperCase() !== 'MEMBER') return;
      const findCol = (label, from, step) => {
        for (let i = from; i >= 0 && i < row.length && Math.abs(i - c) <= 3; i += step) {
          if (row[i].toUpperCase().includes(label)) return i;
        }
        throw new Error(`หาคอลัมน์ ${label} ของตารางที่แถว ${r + 1} ไม่เจอ`);
      };
      tables.push({
        dataRow: r + 2,                       // แถวแรกของข้อมูล (1-based)
        jobCol: findCol('CLASS', c - 1, -1) + 1,
        memberCol: c + 1,
        posCol: findCol('POSITION', c + 1, 1) + 1,
        remarkCol: findCol('REMARK', c + 1, 1) + 1,
      });
    });
  });
  return tables.sort((a, b) => a.dataRow - b.dataRow || a.memberCol - b.memberCol);
}

/** ปรับตัวพิมพ์ให้ตรงกับตัวเลือกใน dropdown (เช่น Gunslinger → gunslinger) */
function matchOption(range, value, fallbackOptions) {
  const rule = range.getDataValidation();
  const options = (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST)
    ? rule.getCriteriaValues()[0] : fallbackOptions;
  const hit = (options || []).find(o => String(o).trim().toLowerCase() === value.toLowerCase());
  return hit !== undefined ? String(hit).trim() : value;
}

function fillPartiesFromText(text) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(PARTY_SHEET_NAME) || ss.getActiveSheet(); // ชื่อแท็บไม่ตรง → ใช้แท็บที่เปิดอยู่

  const parties = parsePartyText(text);
  if (!parties.length) throw new Error('ไม่เจอรายชื่อตี้ในข้อความ');

  const tables = findPartyTables(sheet);
  if (tables.length < PARTIES_PER_TEAM * 2) {
    throw new Error(`เจอตารางแค่ ${tables.length} ตาราง (ต้องมี ${PARTIES_PER_TEAM * 2})`);
  }

  // เก็บค่าที่ใช้อยู่ในชีตไว้เทียบตัวพิมพ์ เผื่อคอลัมน์ไม่มี dropdown แบบ list
  const known = { job: new Set(), pos: new Set() };
  tables.forEach(t => {
    sheet.getRange(t.dataRow, t.jobCol, ROWS_PER_PARTY).getDisplayValues().flat().forEach(v => v && known.job.add(v));
    sheet.getRange(t.dataRow, t.posCol, ROWS_PER_PARTY).getDisplayValues().flat().forEach(v => v && known.pos.add(v));
  });

  const warnings = [];
  parties.forEach(p => {
    const t = tables[p.slot - 1];
    if (!t) { warnings.push(`ข้าม Party ${p.slot}: ไม่มีตารางนี้`); return; }
    if (p.members.length > ROWS_PER_PARTY) warnings.push(`ตารางที่ ${p.slot}: มี ${p.members.length} คน ใส่ได้แค่ ${ROWS_PER_PARTY}`);

    for (let i = 0; i < ROWS_PER_PARTY; i++) {
      const m = p.members[i];
      const row = t.dataRow + i;
      const jobCell = sheet.getRange(row, t.jobCol);
      const posCell = sheet.getRange(row, t.posCol);
      jobCell.setValue(m ? matchOption(jobCell, m.job, [...known.job]) : '');
      sheet.getRange(row, t.memberCol).setValue(m ? m.name : '');
      posCell.setValue(m ? matchOption(posCell, p.run, [...known.pos]) : '');
      sheet.getRange(row, t.remarkCol).setValue(m && p.remark ? `- ${p.remark}` : '');
    }
  });

  const members = parties.reduce((n, p) => n + p.members.length, 0);
  return `✅ จัดเสร็จ ${parties.length} ตี้ / ${members} คน` + (warnings.length ? '\n⚠️ ' + warnings.join('\n⚠️ ') : '');
}
