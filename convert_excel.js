const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

function excelDateToYMD(excelDate) {
  if (!excelDate) return '';
  if (typeof excelDate === 'string') return excelDate.trim();
  if (typeof excelDate === 'number') {
    const date = XLSX.SSF.parse_date_code(excelDate);
    if (date) {
      const y = date.y;
      const m = String(date.m).padStart(2, '0');
      const d = String(date.d).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
  }
  return String(excelDate);
}

const excelPath = path.join(__dirname, 'data', 'warga_dukapil.xlsx');
if (!fs.existsSync(excelPath)) {
  console.error("File not found:", excelPath);
  process.exit(1);
}

const wb = XLSX.readFile(excelPath);
let allWarga = [];

wb.SheetNames.forEach(sheetName => {
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
  // Extract RT number from sheet name (e.g., 'RT 01' -> '01')
  const rtMatch = sheetName.match(/\d+/);
  const rtStr = rtMatch ? String(rtMatch[0]).padStart(2, '0') : '01';

  let currentKk = '';

  rows.forEach((row, rowIdx) => {
    if (rowIdx < 2) return; // skip header rows
    if (!row || row.length === 0) return;

    // Check NIK and NAMA
    const nik = row[3] ? String(row[3]).trim() : '';
    const nama = row[4] ? String(row[4]).trim() : '';
    if (!nama || !isNaN(nama) || nama.toUpperCase() === 'NAMA' || nama.toUpperCase() === 'NAMA LENGKAP' || nama.toUpperCase() === 'NO') return;

    const rawKk = row[2] ? String(row[2]).trim() : '';
    if (rawKk && rawKk.length >= 10 && !isNaN(rawKk)) {
      currentKk = rawKk;
    }

    const jkRaw = row[6] ? String(row[6]).trim().toUpperCase() : 'L';
    const jk = (jkRaw.startsWith('P') || jkRaw === 'PEREMPUAN') ? 'Perempuan' : 'Laki-Laki';

    const tempatLahir = row[7] ? String(row[7]).trim() : '-';
    const tglLahir = excelDateToYMD(row[8]);
    const agama = row[9] ? String(row[9]).trim() : '-';
    const pendidikan = row[10] ? String(row[10]).trim() : '-';
    const pekerjaan = row[12] ? String(row[12]).trim() : '-';
    const statusKawinRaw = row[13] ? String(row[13]).trim() : '-';
    const sdhk = row[15] ? String(row[15]).trim() : '-';

    let statusKawin = statusKawinRaw;
    if (statusKawinRaw.includes('KAWIN') && !statusKawinRaw.includes('BELUM') && !statusKawinRaw.includes('B.')) {
      statusKawin = 'Kawin';
    } else if (statusKawinRaw.includes('B.') || statusKawinRaw.includes('BELUM')) {
      statusKawin = 'Belum Kawin';
    }

    allWarga.push({
      nama: nama,
      nik: nik || '-',
      kk: currentKk || '-',
      jenis_kelamin: jk,
      tempat_lahir: tempatLahir,
      tanggal_lahir: tglLahir,
      agama: agama,
      pendidikan: pendidikan,
      pekerjaan: pekerjaan,
      status_pernikahan: statusKawin,
      sdhk: sdhk,
      rt: rtStr,
      rw: '02',
      alamat: `RT ${rtStr}/RW 02 Kelurahan Cipaganti`
    });
  });
});

console.log('Total Warga Extracted from warga_dukapil.xlsx:', allWarga.length);
if (allWarga.length > 0) {
  console.log('Sample #1:', allWarga[0]);
  console.log('Sample #50:', allWarga[Math.min(49, allWarga.length - 1)]);
  
  // Write to data/dataset_warga.json
  const outputPath = path.join(__dirname, 'data', 'dataset_warga.json');
  fs.writeFileSync(outputPath, JSON.stringify(allWarga, null, 2), 'utf8');
  console.log(`Successfully written ${allWarga.length} records to ${outputPath}`);
}
