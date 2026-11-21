#!/usr/bin/env node
/* =========================================================================
   excel-to-data.js
   -------------------------------------------------------------------------
   Convert an Excel student list (.xlsx / .xls / .csv) into students-data.js —
   the file that EntryFormBUTEX.html loads directly from source code (no backend
   server, no manual file upload needed).

   Usage:
       node tools/excel-to-data.js path/to/students.xlsx
       node tools/excel-to-data.js path/to/students.xlsx output.js

   The first sheet is read. Column headers are matched using the SAME mapping
   as the web app (Bengali headers, with a few short aliases accepted).

   Subjects: the sheet stores 7 subject codes followed by 7 subject names as a
   BLOCK (even if the header labels them interleaved Code1/Name1/...). These are
   collected into window.STUDENT_SUBJECTS and shown for every student.

   Requires:  npm install   (installs the `xlsx` package)  — run once in tools/
   ========================================================================= */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const inputPath = process.argv[2];
const outputPath = process.argv[3] || path.resolve(__dirname, '..', 'students-data.js');

if (!inputPath) {
  console.error('Usage: node tools/excel-to-data.js path/to/students.xlsx [output.js]');
  process.exit(1);
}
if (!fs.existsSync(inputPath)) {
  console.error('File not found: ' + inputPath);
  process.exit(1);
}

// Must mirror COLUMN_MAP / COLUMN_MAP_ALT in EntryFormBUTEX.html
const COLUMN_MAP = {
  'আইডি (ইংরেজি সংখ্যায়)': 'id',
  'নাম (ইংরেজি বড় অক্ষরে)': 'name_english',
  'নাম (বাংলায়)': 'name_bangla',
  'পিতার নাম (এসএসসি সার্টিফিকেট অনুযায়ী বাংলায়)': 'father_name',
  'মাতার নাম (এসএসসি সার্টিফিকেট অনুযায়ী বাংলায়)': 'mother_name',
  'জন্ম তারিখ (এসএসসির সার্টিফিকেট অনুযায়ী)': 'dob',
  'বর্তমান ঠিকানা (বাংলায়)': 'current_address',
  'গ্রাম (বাংলায়)': 'village',
  'ডাকঘর (বাংলায়)': 'post_office',
  'থানা/উপজেলা (বাংলায়)': 'thana',
  'জেলা (বাংলায়)': 'district',
  'ধর্ম (বাংলায়)': 'religion',
  'এসএসসি পাশের সাল (বাংলায়)': 'ssc_year',
  'এসএসসি রোল (বাংলায়)': 'ssc_roll',
  'এসএসসি বোর্ড (বাংলায়)': 'ssc_board',
  'এসএসসি স্কুল (বাংলায়)': 'ssc_school',
  'এইচএসসি সাল (বাংলায়)': 'hsc_year',
  'এইচএসসি রোল (বাংলায়)': 'hsc_roll',
  'এইচএসসি বোর্ড (বাংলায়)': 'hsc_board',
  'এইচএসসি কলেজ (বাংলায়)': 'hsc_college',
  'মোবাইল নম্বর': 'mobile',
  'বিষয় কোড': 'subject_code',
  'বিষয়ের নাম': 'subject_name'
};

const COLUMN_MAP_ALT = {
  'আইডি': 'id',
  'পিতার নাম': 'father_name',
  'মাতার নাম': 'mother_name',
  'জন্ম তারিখ': 'dob',
  'বর্তমান ঠিকানা': 'current_address',
  'গ্রাম': 'village',
  'ডাকঘর': 'post_office',
  'থানা/উপজেলা': 'thana',
  'জেলা': 'district',
  'ধর্ম': 'religion',
  'এসএসসি পাশের সাল': 'ssc_year',
  'এসএসসি রোল': 'ssc_roll',
  'এসএসসি বোর্ড': 'ssc_board',
  'এসএসসি স্কুল': 'ssc_school',
  'এইচএসসি সাল': 'hsc_year',
  'এইচএসসি রোল': 'hsc_roll',
  'এইচএসসি বোর্ড': 'hsc_board',
  'এইচএসসি কলেজ': 'hsc_college'
};

function headerKey(v) {
  return String(v == null ? '' : v).trim().replace(/\s+/g, ' ').replace(/য়/g, 'য়').replace(/ড়/g, 'ড়');
}

function normalizeStudent(raw) {
  const n = {};
  const cleanRaw = {};
  Object.keys(raw).forEach(function (k) {
    cleanRaw[String(k).trim()] = String(raw[k] == null ? '' : raw[k]).trim();
  });
  Object.keys(cleanRaw).forEach(function (rawKey) {
    const nk = headerKey(rawKey);
    Object.keys(COLUMN_MAP).forEach(function (mapKey) {
      if (nk === headerKey(mapKey)) n[COLUMN_MAP[mapKey]] = cleanRaw[rawKey];
    });
    Object.keys(COLUMN_MAP_ALT).forEach(function (mapKey) {
      if (n[COLUMN_MAP_ALT[mapKey]] === undefined && nk === headerKey(mapKey)) {
        n[COLUMN_MAP_ALT[mapKey]] = cleanRaw[rawKey];
      }
    });
  });
  return n;
}

const wb = XLSX.readFile(inputPath);
const ws = wb.Sheets[wb.SheetNames[0]];
// Array-of-arrays so we can read the subject block by column position.
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
if (!aoa.length) {
  console.error('No data rows found in the first sheet of ' + inputPath);
  process.exit(1);
}
const headerRow = aoa[0];

// Locate the 14 "Subject Code/Name N" columns, in spreadsheet order.
const subjectPos = [];
headerRow.forEach((h, i) => {
  if (/^Subject (Code|Name)\s*\d$/.test(String(h == null ? '' : h).trim())) subjectPos.push(i);
});
const codePos = subjectPos.slice(0, 7);
const namePos = subjectPos.slice(7, 14);

const students = [];
const seen = {};
const subjects = [];
const idHeader = 'আইডি (ইংরেজি সংখ্যায়)';

for (let r = 1; r < aoa.length; r++) {
  const row = aoa[r];
  if (!row || !row.length) continue;
  const raw = {};
  headerRow.forEach((h, i) => { raw[h] = row[i]; });
  if (!String(raw[idHeader] == null ? '' : raw[idHeader]).trim()) continue; // skip blank rows

  students.push(normalizeStudent(raw));

  // Collect the 7-subject curriculum (block: codes then names), deduplicated.
  for (let n = 0; n < 7; n++) {
    const code = String(row[codePos[n]] == null ? '' : row[codePos[n]]).trim();
    const name = String(row[namePos[n]] == null ? '' : row[namePos[n]]).trim();
    if (!code && !name) continue;
    const key = code + '|' + name;
    if (!seen[key]) { seen[key] = true; subjects.push({ code: code, name: name }); }
  }
}

if (!students.length) {
  console.error('No student rows found (missing আইডি column?).');
  process.exit(1);
}

const header =
  '/* AUTO-GENERATED by tools/excel-to-data.js — do not edit by hand.\n' +
  '   Regenerate with: node tools/excel-to-data.js path/to/students.xlsx */\n';
const body =
  'window.STUDENT_DATA = ' + JSON.stringify(students, null, 2) + ';\n\n' +
  'window.STUDENT_SUBJECTS = ' + JSON.stringify(subjects, null, 2) + ';\n';

fs.writeFileSync(outputPath, header + body, 'utf8');
console.log('Wrote ' + students.length + ' student(s) and ' + subjects.length + ' subject(s) to ' + outputPath);
