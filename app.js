
function cipabotCalculateAge(tglLahirStr) {
  if (!tglLahirStr) return -1;
  const str = String(tglLahirStr).trim();
  const parts = str.split(/[-/]/);
  let d = 0, m = 0, y = 0;
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      y = parseInt(parts[0]);
      m = parseInt(parts[1]) - 1;
      d = parseInt(parts[2]);
    } else if (parts[2].length === 4) {
      d = parseInt(parts[0]);
      m = parseInt(parts[1]) - 1;
      y = parseInt(parts[2]);
    }
  }
  if (!y || isNaN(y)) return -1;
  const today = new Date();
  let age = today.getFullYear() - y;
  const monthDiff = today.getMonth() - m;
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < d)) {
    age--;
  }
  return age;
}

// ==========================================
// CIPABOT DATA EXPORT UTILITIES (EXCEL & PDF)
// ==========================================
window.cipabotExportExcel = function(dataList, filename = "cipabot_export_warga.xlsx") {
  try {
    if (!dataList || dataList.length === 0) {
      alert("Tidak ada data untuk diekspor!");
      return;
    }
    
    const exportData = dataList.map((w, idx) => {
      const ageNum = w.umur || cipabotCalculateAge(w.tanggal_lahir);
      return {
        "No": idx + 1,
        "Nama Lengkap": w.nama || "-",
        "NIK": w.nik || "-",
        "No. KK": w.kk || "-",
        "Jenis Kelamin": w.jenis_kelamin || "-",
        "Tempat Lahir": w.tempat_lahir || "-",
        "Tanggal Lahir": w.tanggal_lahir || "-",
        "Umur": ageNum > 0 ? ageNum + " Tahun" : "-",
        "Agama": w.agama || "-",
        "Pendidikan Terakhir": w.pendidikan || "-",
        "Pekerjaan / Profesi": w.pekerjaan || "-",
        "Status Pernikahan": w.status_pernikahan || "-",
        "Hubungan Keluarga (SDHK)": w.sdhk || "-",
        "RT": w.rt || "01",
        "RW": w.rw || "02",
        "Alamat Lengkap": w.alamat || "-"
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    worksheet['!cols'] = [
      { wch: 5 }, { wch: 25 }, { wch: 20 }, { wch: 20 }, { wch: 14 },
      { wch: 15 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 22 },
      { wch: 25 }, { wch: 18 }, { wch: 20 }, { wch: 6 }, { wch: 6 }, { wch: 35 }
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Data Warga");
    XLSX.writeFile(workbook, filename);
  } catch(err) {
    console.error("Export Excel Error:", err);
    alert("Gagal mengunduh file Excel: " + err.message);
  }
};

window.cipabotExportPdf = function(reportTitle, dataList) {
  try {
    if (!dataList || dataList.length === 0) {
      alert("Tidak ada data untuk dicetak!");
      return;
    }

    const printWindow = window.open('', '_blank');
    const now = new Date();
    const dateFormatted = `${now.getDate()} ${['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'][now.getMonth()]} ${now.getFullYear()}`;

    const rowsHtml = dataList.map((w, idx) => {
      const ageNum = w.umur || cipabotCalculateAge(w.tanggal_lahir);
      const ageVal = ageNum > 0 ? ageNum + ' Thn' : '-';
      return `
        <tr>
          <td style="text-align: center;">${idx + 1}</td>
          <td><strong>${w.nama || '-'}</strong></td>
          <td>${w.nik || '-'}</td>
          <td>${w.jenis_kelamin || '-'}</td>
          <td style="text-align: center;">${ageVal}</td>
          <td style="text-align: center;">RT ${w.rt || '01'}/RW ${w.rw || '02'}</td>
          <td>${w.pekerjaan || '-'}</td>
          <td>${w.pendidikan || '-'}</td>
          <td>${w.alamat || '-'}</td>
        </tr>
      `;
    }).join("");

    const htmlContent = `
      <!DOCTYPE html>
      <html lang="id">
      <head>
        <meta charset="UTF-8">
        <title>${reportTitle}</title>
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; padding: 25px; color: #1e293b; }
          .header { text-align: center; border-bottom: 3px double #059669; padding-bottom: 12px; margin-bottom: 16px; }
          .header h2 { margin: 0; font-size: 18px; color: #047857; text-transform: uppercase; letter-spacing: 0.5px; }
          .header h3 { margin: 4px 0 0 0; font-size: 13px; color: #475569; font-weight: 500; }
          .report-tag { margin: 8px 0 0 0; font-size: 14px; font-weight: 700; color: #059669; text-transform: uppercase; }
          .meta { display: flex; justify-content: space-between; margin-bottom: 14px; font-size: 11px; color: #64748b; }
          table { width: 100%; border-collapse: collapse; font-size: 11px; }
          th, td { border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; }
          th { background-color: #f1f5f9; color: #0f172a; font-weight: 700; text-transform: uppercase; font-size: 10px; }
          tr:nth-child(even) { background-color: #f8fafc; }
          .footer { margin-top: 30px; display: flex; justify-content: space-between; font-size: 11px; color: #475569; }
          @media print { body { padding: 0; } }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>PEMERINTAH KOTA BANDUNG - KECAMATAN COBLONG</h2>
          <h3>KELURAHAN CIPAGANTI - SISTEM CIPABOT</h3>
          <div class="report-tag">LAPORAN DATA: ${reportTitle.toUpperCase()}</div>
        </div>
        <div class="meta">
          <span>Tanggal Cetak: ${dateFormatted}</span>
          <span>Total Rekod: ${dataList.length} Warga</span>
        </div>
        <table>
          <thead>
            <tr>
              <th style="width: 30px;">No</th>
              <th>Nama Warga</th>
              <th>NIK</th>
              <th>L/P</th>
              <th style="width: 50px;">Usia</th>
              <th style="width: 70px;">Wilayah</th>
              <th>Pekerjaan</th>
              <th>Pendidikan</th>
              <th>Alamat</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
        <div class="footer">
          <div>CIPABOT - Kelurahan Cipaganti, Kota Bandung</div>
          <div style="text-align: right;">
            <p style="margin:0;">Petugas Kelurahan Cipaganti</p>
            <br><br>
            <p style="margin:0; font-weight: 700;">( ____________________________ )</p>
          </div>
        </div>
        <script>
          window.onload = function() { window.print(); }
        </script>
      </body>
      </html>
    `;

    printWindow.document.write(htmlContent);
    printWindow.document.close();
  } catch(err) {
    console.error("Export PDF Error:", err);
    alert("Gagal membuka PDF cetak: " + err.message);
  }
};

// Dataset Warga Kelurahan Cipaganti (RW 02)
let DEFAULT_WARGA = [];
const DATASET_WARGA = [
  {
    nama: "ANDRI RUSTANDI",
    nik: "3273010101900001",
    kk: "3273012010123456",
    jenis_kelamin: "Laki-laki",
    tempat_lahir: "Bandung",
    tanggal_lahir: "15-05-1990",
    agama: "Islam",
    pendidikan: "S1 Teknik Informatika",
    pekerjaan: "Karyawan Swasta",
    status_pernikahan: "Kawin",
    alamat: "Jl. Cipaganti No. 12, RT 01/RW 02",
    rt: "01",
    rw: "02",
    status: "Aktif"
  },
  {
    nama: "SITI AMINAH",
    nik: "3273014203920003",
    kk: "3273012010126789",
    jenis_kelamin: "Perempuan",
    tempat_lahir: "Bandung",
    tanggal_lahir: "02-03-1992",
    agama: "Islam",
    pendidikan: "D3 Kebidanan",
    pekerjaan: "Tenaga Kesehatan",
    status_pernikahan: "Kawin",
    alamat: "Jl. Sampurna No. 4, RT 02/RW 02",
    rt: "02",
    rw: "02",
    status: "Aktif"
  },
  {
    nama: "GILANG RAMADHAN",
    nik: "3273011210010002",
    kk: "3273012010154321",
    jenis_kelamin: "Laki-laki",
    tempat_lahir: "Bandung",
    tanggal_lahir: "12-10-2001",
    agama: "Islam",
    pendidikan: "S1 Sistem Informasi",
    pekerjaan: "Mahasiswa",
    status_pernikahan: "Belum Kawin",
    alamat: "Jl. Cipaganti Barat No. 18, RT 03/RW 02",
    rt: "03",
    rw: "02",
    status: "Aktif"
  },
  {
    nama: "DEDI KUSNADI",
    nik: "3273012508750001",
    kk: "3273012010129876",
    jenis_kelamin: "Laki-laki",
    tempat_lahir: "Sumedang",
    tanggal_lahir: "25-08-1975",
    agama: "Islam",
    pendidikan: "SMA",
    pekerjaan: "Wiraswasta",
    status_pernikahan: "Kawin",
    alamat: "Jl. Cipaganti Tengah No. 5, RT 01/RW 02",
    rt: "01",
    rw: "02",
    status: "Aktif"
  },
  {
    nama: "NENG RATNASARI",
    nik: "3273015509800004",
    kk: "3273012010129876",
    jenis_kelamin: "Perempuan",
    tempat_lahir: "Garut",
    tanggal_lahir: "15-09-1980",
    agama: "Islam",
    pendidikan: "SMA",
    pekerjaan: "Ibu Rumah Tangga",
    status_pernikahan: "Kawin",
    alamat: "Jl. Cipaganti Tengah No. 5, RT 01/RW 02",
    rt: "01",
    rw: "02",
    status: "Aktif"
  },
  {
    nama: "BUDI SANTOSO",
    nik: "3273011406850005",
    kk: "3273012010124455",
    jenis_kelamin: "Laki-laki",
    tempat_lahir: "Semarang",
    tanggal_lahir: "14-06-1985",
    agama: "Kristen",
    pendidikan: "S1 Ekonomi",
    pekerjaan: "Pegawai Negeri Sipil",
    status_pernikahan: "Kawin",
    alamat: "Jl. Cipaganti No. 27, RT 02/RW 02",
    rt: "02",
    rw: "02",
    status: "Aktif"
  },
  {
    nama: "CECEP SURYANA",
    nik: "3273010512680001",
    kk: "3273012010123388",
    jenis_kelamin: "Laki-laki",
    tempat_lahir: "Bandung",
    tanggal_lahir: "05-12-1968",
    agama: "Islam",
    pendidikan: "SMP",
    pekerjaan: "Buruh Harian Lepas",
    status_pernikahan: "Kawin",
    alamat: "Jl. Jurang No. 15, RT 03/RW 02",
    rt: "03",
    rw: "02",
    status: "Aktif"
  },
  {
    nama: "EKA SETIAWAN",
    nik: "3273012211950002",
    kk: "3273012010125599",
    jenis_kelamin: "Laki-laki",
    tempat_lahir: "Bandung",
    tanggal_lahir: "22-11-1995",
    agama: "Islam",
    pendidikan: "D3 Desain Grafis",
    pekerjaan: "Karyawan Swasta",
    status_pernikahan: "Belum Kawin",
    alamat: "Jl. Sampurna Gg. III No. 8, RT 02/RW 02",
    rt: "02",
    rw: "02",
    status: "Aktif"
  },
  {
    nama: "IIS DAHLIAWATI",
    nik: "3273016010720003",
    kk: "3273012010121122",
    jenis_kelamin: "Perempuan",
    tempat_lahir: "Tasikmalaya",
    tanggal_lahir: "20-10-1972",
    agama: "Islam",
    pendidikan: "SMA",
    pekerjaan: "Wiraswasta",
    status_pernikahan: "Cerai Hidup",
    alamat: "Jl. Sampurna No. 10, RT 02/RW 02",
    rt: "02",
    rw: "02",
    status: "Aktif"
  },
  {
    nama: "YAYAN RUHIAN",
    nik: "3273011910820001",
    kk: "3273012010129900",
    jenis_kelamin: "Laki-laki",
    tempat_lahir: "Tasikmalaya",
    tanggal_lahir: "19-10-1982",
    agama: "Islam",
    pendidikan: "SMA",
    pekerjaan: "Instruktur Beladiri",
    status_pernikahan: "Kawin",
    alamat: "Jl. Cipaganti No. 34, RT 01/RW 02",
    rt: "01",
    rw: "02",
    status: "Aktif"
  },
  {
    nama: "ANDREAS WIJAYA",
    nik: "3273010904940003",
    kk: "3273012010127733",
    jenis_kelamin: "Laki-laki",
    tempat_lahir: "Bandung",
    tanggal_lahir: "09-04-1994",
    agama: "Katolik",
    pendidikan: "S1 Teknik Elektro",
    pekerjaan: "Karyawan Swasta",
    status_pernikahan: "Belum Kawin",
    alamat: "Jl. Sampurna Barat No. 11, RT 02/RW 02",
    rt: "02",
    rw: "02",
    status: "Aktif"
  },
  {
    nama: "RIZKY AMALIA",
    nik: "3273014502980004",
    kk: "3273012010128844",
    jenis_kelamin: "Perempuan",
    tempat_lahir: "Jakarta",
    tanggal_lahir: "05-02-1998",
    agama: "Islam",
    pendidikan: "S1 Psikologi",
    pekerjaan: "Karyawan Swasta",
    status_pernikahan: "Belum Kawin",
    alamat: "Jl. Cipaganti Tengah No. 2, RT 01/RW 02",
    rt: "01",
    rw: "02",
    status: "Aktif"
  }
];

// Knuth-Morris-Pratt (KMP) Algorithm Implementation
function kmpSearch(text, pattern) {
  let comparisons = 0;
  text = text.toLowerCase();
  pattern = pattern.toLowerCase();

  const n = text.length;
  const m = pattern.length;
  if (m === 0) return { found: true, index: 0, comparisons: 0 };

  // Compute LPS table
  const lps = new Array(m).fill(0);
  let len = 0;
  let i = 1;
  while (i < m) {
    comparisons++;
    if (pattern[i] === pattern[len]) {
      len++;
      lps[i] = len;
      i++;
    } else {
      if (len !== 0) {
        len = lps[len - 1];
      } else {
        lps[i] = 0;
        i++;
      }
    }
  }

  // Search pattern in text
  i = 0; // index for text
  let j = 0; // index for pattern
  while (i < n) {
    comparisons++;
    if (text[i] === pattern[j]) {
      i++;
      j++;
    }
    if (j === m) {
      return { found: true, index: i - j, comparisons };
    } else if (i < n && text[i] !== pattern[j]) {
      if (j !== 0) {
        j = lps[j - 1];
      } else {
        i++;
      }
    }
  }

  return { found: false, index: -1, comparisons };
}

// Boyer-Moore-Horspool (BM) Algorithm Implementation
function bmhSearch(text, pattern) {
  let comparisons = 0;
  text = text.toLowerCase();
  pattern = pattern.toLowerCase();

  const n = text.length;
  const m = pattern.length;
  if (m === 0) return { found: true, index: 0, comparisons: 0 };

  // Precompute Bad Character table
  const badChar = {};
  for (let k = 0; k < m - 1; k++) {
    badChar[pattern[k]] = m - 1 - k;
  }

  let i = 0; // Alignment shift
  while (i <= n - m) {
    let j = m - 1;

    // Check match from right to left
    while (j >= 0) {
      comparisons++;
      if (pattern[j] === text[i + j]) {
        j--;
      } else {
        break;
      }
    }

    if (j < 0) {
      return { found: true, index: i, comparisons };
    }

    // Shift based on character aligned with pattern's last char
    const shiftChar = text[i + m - 1];
    const shift = badChar[shiftChar] !== undefined ? badChar[shiftChar] : m;
    i += shift;
  }

  return { found: false, index: -1, comparisons };
}

// --- Natural Language Token Extractor ---
function extractSearchTokens(queryStr) {
  let q = String(queryStr).trim().toLowerCase();
  
  const stopWords = new Set([
    "carikan", "cari", "carikanlah", "tolong", "mohon", "saya", "aku", "kami", "kita",
    "siapa", "siapakah", "apakah", "ada", "adakah", "tampilkan", "tunjukkan", "data", "warga", "penduduk",
    "nama", "bernama", "dengan", "informasi", "tentang", "dong", "sih", "kah", "ya", "yang", "di", "ke", "dari",
    "itu", "ini", "tersebut"
  ]);

  const words = q.split(/\s+/).filter(w => w.length > 0);
  const coreWords = words.filter(w => !stopWords.has(w));

  if (coreWords.length > 0) {
    return coreWords.join(" ");
  }
  return q;
}

// --- Levenshtein Distance for Typo Detection ---
function levenshteinDistance(a, b) {
  a = String(a).toLowerCase();
  b = String(b).toLowerCase();
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;

  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + cost
      );
    }
  }
  return dp[m][n];
}

// --- Find Closest Matching Citizen Name for Typo Suggestions ("Mungkin Maksud Anda...") ---
function findFuzzySuggestion(rawQuery) {
  const queryStr = String(rawQuery).trim().toLowerCase();
  if (queryStr.length < 3) return null;

  const cleanStr = extractSearchTokens(queryStr);
  const queryWords = cleanStr.split(/\s+/).filter(w => w.length >= 3);

  let bestMatch = null;
  let minDistance = Infinity;

  for (const warga of DATASET_WARGA) {
    if (!warga.nama) continue;
    const nameLower = warga.nama.toLowerCase();
    const nameWords = nameLower.split(/\s+/);

    // 1. Full string distance
    const distFull = levenshteinDistance(cleanStr, nameLower);
    if (distFull < minDistance && distFull <= Math.max(3, Math.floor(nameLower.length * 0.4))) {
      minDistance = distFull;
      bestMatch = warga;
    }

    // 2. Word-by-word distance (e.g. "Gilagn" vs "Gilang")
    for (const qWord of queryWords) {
      for (const nWord of nameWords) {
        const distWord = levenshteinDistance(qWord, nWord);
        const maxAllowedDist = nWord.length <= 4 ? 1 : 2;
        if (distWord <= maxAllowedDist && distWord < minDistance) {
          minDistance = distWord;
          bestMatch = warga;
        }
      }
    }
  }

  return bestMatch;
}



// Core search runner using KMP & Boyer-Moore with Multi-Criteria Filtering (RT, RW, Name, NIK, etc.)
function searchWarga(queryStr, activeAlgo = "kmp") {
  const rawQuery = String(queryStr).trim();
  
  // --- Extract Multi-Criteria Filters ---
  let tempQuery = rawQuery;
  let targetRt = null;
  let targetRw = null;

  // 1. Extract RT (e.g., "di rt 10", "rt 10", "rt. 10", "rt10", "rt 010")
  const rtMatch = tempQuery.match(/\b(di\s+)?rt\.?\s*[:\=]?\s*0*(\d+)\b/i);
  if (rtMatch) {
    targetRt = parseInt(rtMatch[2], 10);
    tempQuery = tempQuery.replace(rtMatch[0], ' ');
  }

  // 2. Extract RW (e.g., "rw 02", "rw. 2", "rw2", "rw 002")
  const rwMatch = tempQuery.match(/\b(di\s+)?rw\.?\s*[:\=]?\s*0*(\d+)\b/i);
  if (rwMatch) {
    targetRw = parseInt(rwMatch[2], 10);
    tempQuery = tempQuery.replace(rwMatch[0], ' ');
  }

  // 3. Extract Clean Text Keyword
  let cleanKeyword = extractSearchTokens(tempQuery);

  // If cleanKeyword consists only of generic filler words (e.g. "warga", "data", "penduduk"), clear it
  const genericWords = new Set(["warga", "penduduk", "data", "orang", "masyarakat", "informasi"]);
  if (genericWords.has(cleanKeyword.toLowerCase())) {
    cleanKeyword = "";
  }

  // Helper to test if a citizen satisfies RT/RW criteria
  function matchesCriteria(warga) {
    if (targetRt !== null) {
      const wRt = parseInt(warga.rt || '0', 10);
      if (wRt !== targetRt) return false;
    }
    if (targetRw !== null) {
      const wRw = parseInt(warga.rw || '0', 10);
      if (wRw !== targetRw) return false;
    }
    return true;
  }

  function runKmpSearch(pattern) {
    let matches = [];
    let comps = 0;
    
    for (const warga of DATASET_WARGA) {
      if (!matchesCriteria(warga)) continue;

      if (!pattern || pattern.length === 0) {
        matches.push(warga);
        continue;
      }

      const searchNama = kmpSearch(warga.nama || '', pattern);
      comps += searchNama.comparisons;
      let found = searchNama.found;

      if (!found) {
        const searchNik = kmpSearch(warga.nik || '', pattern);
        comps += searchNik.comparisons;
        if (searchNik.found) found = true;
      }
      if (!found) {
        const searchPek = kmpSearch(warga.pekerjaan || '', pattern);
        comps += searchPek.comparisons;
        if (searchPek.found) found = true;
      }
      if (!found) {
        const searchAgm = kmpSearch(warga.agama || '', pattern);
        comps += searchAgm.comparisons;
        if (searchAgm.found) found = true;
      }
      if (!found) {
        const searchPnd = kmpSearch(warga.pendidikan || '', pattern);
        comps += searchPnd.comparisons;
        if (searchPnd.found) found = true;
      }

      if (found) matches.push(warga);
    }
    return { matches, comps };
  }

  // 1. Try search with cleanKeyword
  const kmpStart = performance.now();
  let kmpRes = runKmpSearch(cleanKeyword);
  let usedPattern = cleanKeyword;

  // 2. If no match and keyword has multiple words, try searching each word
  if (kmpRes.matches.length === 0 && cleanKeyword.includes(" ")) {
    const words = cleanKeyword.split(/\s+/).filter(w => w.length >= 2);
    let matchedSet = null;
    let totalComps = kmpRes.comps;

    for (const word of words) {
      const wordRes = runKmpSearch(word);
      totalComps += wordRes.comps;
      if (matchedSet === null) {
        matchedSet = wordRes.matches;
      } else {
        matchedSet = matchedSet.filter(m => wordRes.matches.some(wm => wm.nik === m.nik && wm.nama === m.nama));
      }
    }
    if (matchedSet && matchedSet.length > 0) {
      kmpRes.matches = matchedSet;
      kmpRes.comps = totalComps;
      usedPattern = words.join(" ");
    }
  }

  const kmpEnd = performance.now();
  const kmpDuration = kmpEnd - kmpStart;

  // 3. Boyer-Moore Search Simulation
  const bmStart = performance.now();
  let bmComparisons = 0;
  let bmMatches = [];

  for (const warga of DATASET_WARGA) {
    if (!matchesCriteria(warga)) continue;

    if (!usedPattern || usedPattern.length === 0) {
      bmMatches.push(warga);
      continue;
    }

    const searchNama = bmhSearch(warga.nama || '', usedPattern);
    bmComparisons += searchNama.comparisons;
    let found = searchNama.found;

    if (!found) {
      const searchNik = bmhSearch(warga.nik || '', usedPattern);
      bmComparisons += searchNik.comparisons;
      if (searchNik.found) found = true;
    }
    if (!found) {
      const searchPek = bmhSearch(warga.pekerjaan || '', usedPattern);
      bmComparisons += searchPek.comparisons;
      if (searchPek.found) found = true;
    }
    if (!found) {
      const searchAgm = bmhSearch(warga.agama || '', usedPattern);
      bmComparisons += searchAgm.comparisons;
      if (searchAgm.found) found = true;
    }
    if (!found) {
      const searchPnd = bmhSearch(warga.pendidikan || '', usedPattern);
      bmComparisons += searchPnd.comparisons;
      if (searchPnd.found) found = true;
    }

    if (found) bmMatches.push(warga);
  }
  const bmEnd = performance.now();
  const bmDuration = bmEnd - bmStart;

  const finalMatches = kmpRes.matches.length > 0 ? kmpRes.matches : bmMatches;

  // 4. Fuzzy typo suggestion if no exact match found
  let suggestion = null;
  if (finalMatches.length === 0) {
    suggestion = findFuzzySuggestion(rawQuery);
  }

  return {
    query: rawQuery,
    usedPattern: usedPattern,
    targetRt: targetRt,
    targetRw: targetRw,
    warga: finalMatches.length > 0 ? finalMatches[0] : null,
    matches: finalMatches,
    suggestion: suggestion,
    activeAlgo: activeAlgo,
    kmp: {
      comparisons: kmpRes.comps,
      duration: kmpDuration.toFixed(4),
      status: finalMatches.length > 0 ? "Cocok" : "Tidak Ditemukan"
    },
    bm: {
      comparisons: bmComparisons,
      duration: bmDuration.toFixed(4),
      status: finalMatches.length > 0 ? "Cocok" : "Tidak Ditemukan"
    }
  };
}


document.addEventListener("DOMContentLoaded", () => {
  const chatBody = document.getElementById("chat-body");
  const chatInput = document.getElementById("chat-input-field");
  const sendBtn = document.getElementById("chat-send-btn");
  const algoSelect = document.getElementById("algo-select");
  const welcomeCard = document.getElementById("welcome-card");
  const newSearchBtn = document.getElementById("btn-new-search");
  const historyList = document.getElementById("history-list");
  const excelFileInput = document.getElementById("excel-file-input");
  const excelUploadLabel = document.getElementById("excel-upload-label");
  
  let history = JSON.parse(localStorage.getItem("cipabot_history")) || [];

  function updateHistoryUI() {
    if (!historyList) return;
    historyList.innerHTML = "";
    
    if (history.length === 0) {
      historyList.innerHTML = `<div style="font-size: 12px; color: var(--text-muted); text-align: center; padding: 12px 0;">Belum ada riwayat</div>`;
      return;
    }

    history.forEach((item, index) => {
      const div = document.createElement("div");
      div.className = "history-item";
      div.innerHTML = `
        <div class="history-query">${item.query}</div>
        <div class="history-meta">${item.algo.toUpperCase()} • ${item.time}</div>
      `;
      div.addEventListener("click", () => {
        if (chatInput) chatInput.value = item.query;
        if (algoSelect) algoSelect.value = item.algo;
        handleSearch(item.query, item.algo);
      });
      historyList.appendChild(div);
    });
  }

  function saveToHistory(query, algo) {
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    // Deduplicate
    history = history.filter(item => item.query.toLowerCase() !== query.toLowerCase());
    history.unshift({ query, algo, time: timeStr });
    if (history.length > 8) history.pop();
    updateHistoryUI();
  }

  function appendUserMessage(text) {
    const wrapper = document.createElement("div");
    wrapper.className = "chat-bubble-wrapper user";
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    
    wrapper.innerHTML = `
      <div class="chat-bubble user">${text}</div>
      <div class="chat-time">${timeStr}</div>
    `;
    chatBody.appendChild(wrapper);
    scrollToBottom();
  }

  function appendBotTyping() {
    const wrapper = document.createElement("div");
    wrapper.className = "chat-bubble-wrapper bot typing-indicator-wrapper";
    wrapper.innerHTML = `
      <div class="typing-indicator">
        <div class="typing-dot"></div>
        <div class="typing-dot"></div>
        <div class="typing-dot"></div>
      </div>
    `;
    chatBody.appendChild(wrapper);
    scrollToBottom();
    return wrapper;
  }

  function appendExcelLoadSuccess(fileName, count, sheetCount) {
    const wrapper = document.createElement("div");
    wrapper.className = "chat-bubble-wrapper bot";
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    
    wrapper.innerHTML = `
      <div class="chat-bubble bot" style="background-color: var(--success-light); border: 1px solid rgba(16, 185, 129, 0.15); color: #065F46;">
        <div style="display: flex; align-items: center; gap: 8px;">
          <i data-lucide="file-spreadsheet" style="width: 20px; height: 20px; color: var(--success);"></i>
          <span>
            Berhasil memuat dataset Excel <strong>${fileName}</strong> berisi <strong>${sheetCount} sheet</strong> dengan total <strong>${count} data warga</strong>. Semua pencarian sekarang akan dijalankan pada seluruh data dari semua sheet tersebut!
          </span>
        </div>
      </div>
      <div class="chat-time">${timeStr}</div>
    `;
    chatBody.appendChild(wrapper);
    if (typeof lucide !== "undefined") {
      lucide.createIcons();
    }
    scrollToBottom();
  }

  function appendBotResult(result, algo) {
    const wrapper = document.createElement("div");
    wrapper.className = "chat-bubble-wrapper bot";
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    const cardId = 'search-result-' + Math.random().toString(36).substring(2, 9);

    if (!result.matches || result.matches.length === 0) {
      const suggestHtml = result.suggestion ? `
        <div style="background-color: #FEF3C7; border: 1px solid #FDE68A; border-radius: 12px; padding: 14px 16px; margin-top: 10px; margin-bottom: 6px;">
          <div style="font-size: 13px; font-weight: 700; color: #B45309; display: flex; align-items: center; gap: 6px; margin-bottom: 8px;">
            <i data-lucide="sparkles" style="width: 16px; height: 16px; color: #D97706;"></i>
            Mungkin maksud Anda:
          </div>
          <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
            <button class="btn-fuzzy-suggest" data-name="${result.suggestion.nama}" style="background-color: #F59E0B; color: white; border: none; padding: 8px 16px; border-radius: 100px; font-size: 12px; font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 6px; transition: all 0.2s ease; box-shadow: 0 2px 6px rgba(245,158,11,0.2);">
              <i data-lucide="search" style="width: 13px; height: 13px;"></i>
              Cari "${result.suggestion.nama}"
            </button>
          </div>
        </div>
      ` : '';

      wrapper.innerHTML = `
        <div class="chat-bubble bot">
          Maaf, data warga dengan kueri <strong>"${result.query || ''}"</strong> tidak ditemukan dalam dataset Excel Kelurahan Cipaganti.
        </div>
        ${suggestHtml}
        <div class="result-card" style="border-color: #FCA5A5;">
          <div class="result-card-header" style="background-color: #FEF2F2;">
            <div class="result-card-title" style="color: #EF4444;">
              <i data-lucide="alert-circle" style="width: 16px; height: 16px;"></i>
              Pencarian Gagal
            </div>
            <span class="result-badge" style="background-color: #FEF2F2; color: #EF4444; border: 1px solid #FCA5A5;">Selesai</span>
          </div>
          
<!-- Algoritma Comparison Bar -->
          <div style="background-color: var(--bg-light); border-top: 1px solid var(--border); padding: 16px 20px;">
            <div style="font-size: 11px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 12px; display: flex; align-items: center; gap: 6px;">
              <i data-lucide="bar-chart-2" style="width: 14px; height: 14px; color: var(--primary);"></i>
              Perbandingan Algoritma (Kueri Tidak Cocok)
            </div>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
              <div style="background-color: var(--bg-white); border: 1px solid var(--border); padding: 12px; border-radius: var(--radius-md); ${algo === 'kmp' ? 'border-color: var(--primary); box-shadow: 0 0 0 2px var(--primary-light);' : ''}">
                <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">KMP Algorithm</div>
                <div style="font-size: 11px; color: var(--text-muted);">Waktu: <strong>${result.kmp.duration} ms</strong></div>
                <div style="font-size: 11px; color: var(--text-muted);">Perbandingan: <strong>${result.kmp.comparisons} Karakter</strong></div>
              </div>
              <div style="background-color: var(--bg-white); border: 1px solid var(--border); padding: 12px; border-radius: var(--radius-md); ${algo === 'bm' ? 'border-color: var(--primary); box-shadow: 0 0 0 2px var(--primary-light);' : ''}">
                <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">Boyer-Moore Algorithm</div>
                <div style="font-size: 11px; color: var(--text-muted);">Waktu: <strong>${result.bm.duration} ms</strong></div>
                <div style="font-size: 11px; color: var(--text-muted);">Perbandingan: <strong>${result.bm.comparisons} Karakter</strong></div>
              </div>
            </div>
          </div>
        </div>
        <div class="chat-time">${timeStr}</div>
      `;
    } else {
      // Determine efficiency winner
      let winnerMsg = "";
      if (result.bm.comparisons === result.kmp.comparisons) {
        winnerMsg = `Kedua algoritma melakukan jumlah perbandingan karakter yang sama (${result.bm.comparisons}).`;
      } else if (result.bm.comparisons < result.kmp.comparisons) {
        const diff = result.kmp.comparisons - result.bm.comparisons;
        winnerMsg = `<strong>Boyer-Moore lebih efisien</strong> dengan <strong>${diff} lebih sedikit</strong> perbandingan karakter dibanding KMP.`;
      } else {
        const diff = result.bm.comparisons - result.kmp.comparisons;
        winnerMsg = `<strong>KMP lebih efisien</strong> dengan <strong>${diff} lebih sedikit</strong> perbandingan karakter dibanding Boyer-Moore.`;
      }

      const isMultiMatch = result.matches.length > 1;

      if (isMultiMatch) {
        const matchesRowsHtml = result.matches.map((item, i) => `
          <div class="search-item-row" data-idx="${i}" style="background-color: var(--bg-white); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px; margin-bottom: 8px; display: flex; align-items: center; justify-content: space-between; cursor: pointer; transition: all 0.2s ease;">
            <div style="display: flex; align-items: center; gap: 10px;">
              <div style="width: 28px; height: 28px; border-radius: 50%; background-color: var(--primary-light); color: var(--primary); display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700;">
                ${i + 1}
              </div>
              <div>
                <div style="font-size: 13px; font-weight: 700; color: var(--text-main); display: flex; align-items: center; gap: 6px;">
                  ${item.nama}
                  <span style="font-size: 10px; font-weight: 600; padding: 2px 6px; border-radius: 4px; background-color: var(--bg-light); color: var(--text-muted);">${item.jenis_kelamin || '-'}</span>
                  <i data-lucide="chevron-right" style="width: 14px; height: 14px; color: var(--primary);"></i>
                </div>
                <div style="font-size: 11px; color: var(--text-muted); font-family: monospace;">NIK: ${item.nik || '-'} • KK: ${item.kk || '-'} • ${item.tempat_lahir || ''} ${item.tanggal_lahir || ''}</div>
              </div>
            </div>
            <div style="text-align: right;">
              <span style="font-size: 11px; color: var(--primary); font-weight: 700;">Lihat Biodata →</span>
            </div>
          </div>
        `).join("");

        wrapper.innerHTML = `
          <div class="chat-bubble bot">
            Ditemukan <strong>${result.matches.length} data warga</strong> yang cocok dengan kueri <strong>"${result.query}"</strong>:
          </div>
          <div class="result-card" id="${cardId}">
            
            <!-- VIEW 1: DAFTAR WARGA DITEMUKAN -->
            <div class="multi-list-view">
              <div class="result-card-header">
                <div class="result-card-title">
                  <i data-lucide="users" style="width: 16px; height: 16px; color: var(--primary);"></i>
                  ${result.matches.length} Warga Ditemukan ("${result.query}")
                </div>
                <span class="result-badge">${result.matches.length} Hasil</span>
              </div>

              <div style="padding: 16px 20px;">
                <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
                  <i data-lucide="list" style="width: 14px; height: 14px; color: var(--primary);"></i>
                  Klik nama warga untuk melihat detail biodata lengkap:
                </div>
                <div style="max-height: 280px; overflow-y: auto; padding-right: 4px;">
                  ${matchesRowsHtml}
                </div>
              </div>

              <!-- Algoritma Comparison Bar -->
              <div style="background-color: var(--bg-light); border-top: 1px solid var(--border); padding: 16px 20px;">
                <div style="font-size: 11px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 12px; display: flex; align-items: center; gap: 6px;">
                  <i data-lucide="bar-chart-2" style="width: 14px; height: 14px; color: var(--primary);"></i>
                  Perbandingan Algoritma Pencocokan String
                </div>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 12px;">
                  <div style="background-color: var(--bg-white); border: 1px solid var(--border); padding: 12px; border-radius: var(--radius-md); ${algo === 'kmp' ? 'border-color: var(--primary); box-shadow: 0 0 0 2px var(--primary-light);' : ''}">
                    <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">KMP (Knuth-Morris-Pratt)</div>
                    <div style="font-size: 11px; color: var(--text-muted);">Waktu: <strong style="color: var(--text-main);">${result.kmp.duration} ms</strong></div>
                    <div style="font-size: 11px; color: var(--text-muted);">Perbandingan: <strong style="color: var(--text-main);">${result.kmp.comparisons} Karakter</strong></div>
                  </div>
                  <div style="background-color: var(--bg-white); border: 1px solid var(--border); padding: 12px; border-radius: var(--radius-md); ${algo === 'bm' ? 'border-color: var(--primary); box-shadow: 0 0 0 2px var(--primary-light);' : ''}">
                    <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">BM (Boyer-Moore)</div>
                    <div style="font-size: 11px; color: var(--text-muted);">Waktu: <strong style="color: var(--text-main);">${result.bm.duration} ms</strong></div>
                    <div style="font-size: 11px; color: var(--text-muted);">Perbandingan: <strong style="color: var(--text-main);">${result.bm.comparisons} Karakter</strong></div>
                  </div>
                </div>
                <div style="background-color: var(--primary-light); border: 1px solid rgba(37,99,235,0.1); padding: 10px 14px; border-radius: 8px; font-size: 12px; color: var(--primary); font-weight: 500; display: flex; align-items: center; gap: 8px;">
                  <i data-lucide="info" style="width: 16px; height: 16px; flex-shrink: 0;"></i>
                  <span>${winnerMsg}</span>
                </div>
              </div>

            </div>

            <!-- VIEW 2: DETAIL WARGA (Hidden initially) -->
            <div class="multi-detail-view" style="display: none;">
              <div class="result-card-header" style="background-color: var(--primary-light);">
                <button class="btn-back-multi" style="background: white; border: 1px solid var(--border); padding: 6px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; color: var(--primary); cursor: pointer; display: flex; align-items: center; gap: 6px;">
                  <i data-lucide="arrow-left" style="width: 14px; height: 14px;"></i>
                  ← Kembali ke Hasil Pencarian ("${result.query}")
                </button>
                <span class="result-badge" style="background-color: var(--primary); color: white;">Terverifikasi</span>
              </div>

              <div class="multi-detail-content" style="padding: 16px 20px;">
                <!-- Dynamically filled on item click -->
              </div>

              <div class="performance-metrics-bar" style="background-color: var(--bg-light); font-size: 11px; padding: 10px 20px; color: var(--text-muted); border-top: 1px solid var(--border);">
                <button class="btn-back-multi-bottom" style="background: transparent; border: none; font-size: 12px; font-weight: 600; color: var(--primary); cursor: pointer; display: flex; align-items: center; gap: 6px;">
                  ← Kembali ke Daftar Hasil Pencarian (${result.matches.length} warga)
                </button>
              </div>
            </div>

          </div>
          <div class="chat-time">${timeStr}</div>
        `;
      } else {
        // Single match UI
        const w = result.warga;
        wrapper.innerHTML = `
          <div class="chat-bubble bot">
            Data warga berhasil ditemukan. Berikut adalah rincian informasi penduduk:
          </div>
          <div class="result-card" id="${cardId}">
            <div class="result-card-header">
              <div class="result-card-title">
                <i data-lucide="user" style="width: 16px; height: 16px; color: var(--primary);"></i>
                Data Warga Ditemukan
              </div>
              <div style="display: flex; gap: 6px; align-items: center;">
    <span class="result-badge">Terverifikasi</span>
    <button class="btn-export-single-excel" style="background-color: #ECFDF5; border: 1px solid #A7F3D0; color: #047857; border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 4px;">📥 Excel</button>
    <button class="btn-export-single-pdf" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; color: #1D4ED8; border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 4px;">📄 PDF</button>
  </div>
            </div>
            <div class="result-card-grid">
              <div class="result-item">
                <div class="result-label">Nama Lengkap</div>
                <div class="result-value">${w.nama}</div>
              </div>
              <div class="result-item">
                <div class="result-label">NIK</div>
                <div class="result-value">${w.nik || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label">Nomor KK</div>
                <div class="result-value">${w.kk || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label">Jenis Kelamin</div>
                <div class="result-value">${w.jenis_kelamin || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label">Tempat Lahir</div>
                <div class="result-value">${w.tempat_lahir || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label">Tanggal Lahir</div>
                <div class="result-value">${w.tanggal_lahir || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label">Agama</div>
                <div class="result-value">${w.agama || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label">Pendidikan Terakhir</div>
                <div class="result-value">${w.pendidikan || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label">Pekerjaan</div>
                <div class="result-value">${w.pekerjaan || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label">Status Pernikahan</div>
                <div class="result-value">${w.status_pernikahan || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label">Hub. Keluarga (SDHK)</div>
                <div class="result-value">${w.sdhk || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label">RT / RW</div>
                <div class="result-value">RT ${w.rt || '01'} / RW ${w.rw || '02'}</div>
              </div>
              ${w.alamat ? `<div class="result-item">
                <div class="result-label">Alamat Lengkap</div>
                <div class="result-value">${w.alamat}</div>
              </div>` : ''}
            </div>

            <!-- Algoritma Comparison Bar -->
            <div style="background-color: var(--bg-light); border-top: 1px solid var(--border); padding: 16px 20px;">
              <div style="font-size: 11px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 12px; display: flex; align-items: center; gap: 6px;">
                <i data-lucide="bar-chart-2" style="width: 14px; height: 14px; color: var(--primary);"></i>
                Perbandingan Algoritma Pencocokan String
              </div>
              <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 12px;">
                <div style="background-color: var(--bg-white); border: 1px solid var(--border); padding: 12px; border-radius: var(--radius-md); ${algo === 'kmp' ? 'border-color: var(--primary); box-shadow: 0 0 0 2px var(--primary-light);' : ''}">
                  <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">KMP (Knuth-Morris-Pratt)</div>
                  <div style="font-size: 11px; color: var(--text-muted);">Waktu: <strong style="color: var(--text-main);">${result.kmp.duration} ms</strong></div>
                  <div style="font-size: 11px; color: var(--text-muted);">Perbandingan: <strong style="color: var(--text-main);">${result.kmp.comparisons} Karakter</strong></div>
                </div>
                <div style="background-color: var(--bg-white); border: 1px solid var(--border); padding: 12px; border-radius: var(--radius-md); ${algo === 'bm' ? 'border-color: var(--primary); box-shadow: 0 0 0 2px var(--primary-light);' : ''}">
                  <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">BM (Boyer-Moore)</div>
                  <div style="font-size: 11px; color: var(--text-muted);">Waktu: <strong style="color: var(--text-main);">${result.bm.duration} ms</strong></div>
                  <div style="font-size: 11px; color: var(--text-muted);">Perbandingan: <strong style="color: var(--text-main);">${result.bm.comparisons} Karakter</strong></div>
                </div>
              </div>
              <div style="background-color: var(--primary-light); border: 1px solid rgba(37,99,235,0.1); padding: 10px 14px; border-radius: 8px; font-size: 12px; color: var(--primary); font-weight: 500; display: flex; align-items: center; gap: 8px;">
                <i data-lucide="info" style="width: 16px; height: 16px; flex-shrink: 0;"></i>
                <span>${winnerMsg}</span>
              </div>
            </div>
          </div>
          <div class="chat-time">${timeStr}</div>
        `;
      }
    }
    
    chatBody.appendChild(wrapper);
    if (typeof lucide !== "undefined") {
      lucide.createIcons();
    }

    // Attach click events if multi-match
    if (result.matches && result.matches.length > 1) {
      const cardEl = document.getElementById(cardId);
      if (cardEl) {
        const listView = cardEl.querySelector('.multi-list-view');
        const detailView = cardEl.querySelector('.multi-detail-view');
        const detailContentArea = cardEl.querySelector('.multi-detail-content');
        const itemRows = cardEl.querySelectorAll('.search-item-row');
        const backBtns = cardEl.querySelectorAll('.btn-back-multi, .btn-back-multi-bottom');

        itemRows.forEach(rowEl => {
          rowEl.addEventListener('click', () => {
            const idx = parseInt(rowEl.getAttribute('data-idx'));
            const w = result.matches[idx];
            if (!w) return;

            detailContentArea.innerHTML = `
              <div style="font-size: 14px; font-weight: 700; color: var(--text-main); margin-bottom: 12px; display: flex; align-items: center; gap: 8px;">
                <i data-lucide="user-check" style="width: 18px; height: 18px; color: var(--primary);"></i>
                Rincian Biodata Warga: ${w.nama}
              </div>

              <div class="result-card-grid" style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Nama Lengkap</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 700; color: var(--text-main);">${w.nama}</div>
                </div>
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">NIK</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 700; color: var(--text-main);">${w.nik || '-'}</div>
                </div>
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Nomor KK</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 700; color: var(--text-main);">${w.kk || '-'}</div>
                </div>
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Jenis Kelamin</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.jenis_kelamin || '-'}</div>
                </div>
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tempat Lahir</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tempat_lahir || '-'}</div>
                </div>
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tanggal Lahir</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tanggal_lahir || '-'}</div>
                </div>
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Agama</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.agama || '-'}</div>
                </div>
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Pendidikan Terakhir</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.pendidikan || '-'}</div>
                </div>
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Pekerjaan</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.pekerjaan || '-'}</div>
                </div>
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Status Pernikahan</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.status_pernikahan || '-'}</div>
                </div>
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Hub. Keluarga (SDHK)</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.sdhk || '-'}</div>
                </div>
                <div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">RT / RW</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">RT ${w.rt || '01'} / RW ${w.rw || '02'}</div>
                </div>
                ${w.alamat ? `<div class="result-item">
                  <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Alamat Lengkap</div>
                  <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.alamat}</div>
                </div>` : ''}
              </div>
            `;

            listView.style.display = 'none';
            detailView.style.display = 'block';
            scrollToBottom();
          });
        });

        backBtns.forEach(btn => {
          btn.addEventListener('click', () => {
            detailView.style.display = 'none';
            listView.style.display = 'block';
            scrollToBottom();
          });
        });
      }
    }

    // Attach click listener for fuzzy suggestion button if present
    const suggestBtn = wrapper.querySelector('.btn-fuzzy-suggest');
    if (suggestBtn) {
      suggestBtn.addEventListener('click', () => {
        const suggestName = suggestBtn.getAttribute('data-name');
        if (suggestName) {
          const chatInput = document.getElementById("chat-input-field");
          const algoSelect = document.getElementById("algo-select");
          if (chatInput) {
            chatInput.value = suggestName;
            handleSearch(suggestName, algoSelect ? algoSelect.value : "kmp");
          }
        }
      });
    }

    
    // Attach Export Listeners for Search Result
    const searchCardEl = document.getElementById(cardId);
    if (searchCardEl) {
      // Single warga result (btn-export-single-excel / btn-export-single-pdf)
      if (result.warga) {
        searchCardEl.querySelectorAll('.btn-export-single-excel').forEach(btn => {
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            if (typeof XLSX === 'undefined') { alert('Library XLSX belum dimuat. Refresh halaman dan coba lagi.'); return; }
            window.cipabotExportExcel([result.warga], `warga_${result.warga.nama}.xlsx`);
          });
        });
        searchCardEl.querySelectorAll('.btn-export-single-pdf').forEach(btn => {
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            window.cipabotExportPdf(`Biodata Warga - ${result.warga.nama}`, [result.warga]);
          });
        });
      }
      // Multi matches result (btn-export-multi-excel / btn-export-multi-pdf)
      if (result.matches && result.matches.length > 0) {
        searchCardEl.querySelectorAll('.btn-export-multi-excel').forEach(btn => {
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            if (typeof XLSX === 'undefined') { alert('Library XLSX belum dimuat. Refresh halaman dan coba lagi.'); return; }
            window.cipabotExportExcel(result.matches, `hasil_pencarian_${result.query}.xlsx`);
          });
        });
        searchCardEl.querySelectorAll('.btn-export-multi-pdf').forEach(btn => {
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            window.cipabotExportPdf(`Hasil Pencarian Warga - ${result.query}`, result.matches);
          });
        });
      }
    }

    scrollToBottom();
  }

  function appendBotMessage(text) {
    const wrapper = document.createElement("div");
    wrapper.className = "chat-bubble-wrapper bot";
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    
    wrapper.innerHTML = `
      <div class="chat-bubble bot">${text}</div>
      <div class="chat-time">${timeStr}</div>
    `;
    chatBody.appendChild(wrapper);
    scrollToBottom();
  }

  function appendBotStatsCard(total, laki, pctLaki, perempuan, pctPerempuan) {
    const wrapper = document.createElement("div");
    wrapper.className = "chat-bubble-wrapper bot";
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    
    wrapper.innerHTML = `
      <div class="result-card" style="border-color: var(--primary-light);">
        <div class="result-card-header" style="background-color: var(--primary-light);">
          <div class="result-card-title">
            <i data-lucide="bar-chart-3" style="width: 16px; height: 16px; color: var(--primary);"></i>
            Statistik Demografi Warga
          </div>
          <span class="result-badge" style="background-color: var(--primary); color: white;">Aktif</span>
        </div>
        <div class="result-card-grid" style="grid-template-columns: 1fr; padding: 20px 24px; gap: 16px;">
          
          <!-- Total Warga -->
          <div style="display: flex; align-items: center; justify-content: space-between; padding-bottom: 12px; border-bottom: 1px solid var(--border);">
            <span style="font-size: 13px; font-weight: 600; color: var(--text-muted);">TOTAL DATASET WARGA</span>
            <span style="font-size: 18px; font-weight: 700; color: var(--text-main);">${total} Orang</span>
          </div>

          <!-- Laki-laki -->
          <div style="display: flex; flex-direction: column; gap: 6px;">
            <div style="display: flex; justify-content: space-between; align-items: center; font-size: 13px; font-weight: 600;">
              <span style="display: flex; align-items: center; gap: 6px; color: var(--text-main);">
                <i data-lucide="user" style="width: 14px; height: 14px; color: var(--primary);"></i>
                Laki-laki (Pria)
              </span>
              <span style="color: var(--text-main); font-weight: 700;">${laki} Orang (${pctLaki}%)</span>
            </div>
            <!-- Progress Bar -->
            <div style="width: 100%; height: 8px; background-color: var(--bg-light); border-radius: 100px; overflow: hidden;">
              <div style="width: ${pctLaki}%; height: 100%; background-color: var(--primary); border-radius: 100px;"></div>
            </div>
          </div>

          <!-- Perempuan -->
          <div style="display: flex; flex-direction: column; gap: 6px; margin-top: 4px;">
            <div style="display: flex; justify-content: space-between; align-items: center; font-size: 13px; font-weight: 600;">
              <span style="display: flex; align-items: center; gap: 6px; color: var(--text-main);">
                <i data-lucide="user" style="width: 14px; height: 14px; color: #EC4899;"></i>
                Perempuan (Wanita)
              </span>
              <span style="color: var(--text-main); font-weight: 700;">${perempuan} Orang (${pctPerempuan}%)</span>
            </div>
            <!-- Progress Bar -->
            <div style="width: 100%; height: 8px; background-color: var(--bg-light); border-radius: 100px; overflow: hidden;">
              <div style="width: ${pctPerempuan}%; height: 100%; background-color: #EC4899; border-radius: 100px;"></div>
            </div>
          </div>

        </div>
        <div class="performance-metrics-bar" style="background-color: var(--bg-light); font-size: 11px; padding: 10px 24px; color: var(--text-muted);">
          <span>Dihitung otomatis berdasarkan file Excel / dataset aktif secara real-time.</span>
        </div>
      </div>
      <div class="chat-time">${timeStr}</div>
    `;
    
    chatBody.appendChild(wrapper);
    if (typeof lucide !== "undefined") {
      lucide.createIcons();
    }
    scrollToBottom();
  }

  function calculateAge(tglLahirStr) {
    if (!tglLahirStr) return -1;
    const str = String(tglLahirStr).trim();
    const parts = str.split(/[-/]/);
    let d = 0, m = 0, y = 0;
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        y = parseInt(parts[0]);
        m = parseInt(parts[1]) - 1;
        d = parseInt(parts[2]);
      } else if (parts[2].length === 4) {
        d = parseInt(parts[0]);
        m = parseInt(parts[1]) - 1;
        y = parseInt(parts[2]);
      }
    }
    if (!y || isNaN(y)) return -1;
    const today = new Date();
    let age = today.getFullYear() - y;
    const monthDiff = today.getMonth() - m;
    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < d)) {
      age--;
    }
    return age;
  }

  function appendBotLansiaCard(totalDataset, lansiaList, genderFilter = "all", totalAllLansia = 0) {
    const wrapper = document.createElement("div");
    wrapper.className = "chat-bubble-wrapper bot";
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    
    const countLansia = lansiaList.length;
    const pctLansia = totalDataset > 0 ? ((countLansia / totalDataset) * 100).toFixed(1) : 0;
    const cardId = 'lansia-card-' + Math.random().toString(36).substring(2, 9);

    let cardTitle = "Data Warga Lansia (Usia ≥ 60 Tahun)";
    let badgeLabel = `Total: ${countLansia} Orang`;
    let badgeColor = "#F59E0B";
    let headerBg = "#FEF3C7";
    let titleColor = "#B45309";
    let iconName = "heart-handshake";

    if (genderFilter === "male") {
      cardTitle = "Data Lansia Laki-laki (Usia ≥ 60 Tahun)";
      badgeLabel = `${countLansia} Laki-laki`;
      badgeColor = "#2563EB";
      headerBg = "#EFF6FF";
      titleColor = "#1D4ED8";
      iconName = "user";
    } else if (genderFilter === "female") {
      cardTitle = "Data Lansia Perempuan (Usia ≥ 60 Tahun)";
      badgeLabel = `${countLansia} Perempuan`;
      badgeColor = "#EC4899";
      headerBg = "#FDF2F8";
      titleColor = "#BE185D";
      iconName = "user";
    }

    let lansiaRowsHtml = "";
    if (countLansia === 0) {
      lansiaRowsHtml = `
        <div style="text-align: center; padding: 20px; color: var(--text-muted); font-size: 13px;">
          Tidak ditemukan warga lansia (${genderFilter === 'male' ? 'Laki-laki' : genderFilter === 'female' ? 'Perempuan' : 'kategori lansia'}) dalam dataset ini.
        </div>
      `;
    } else {
      lansiaRowsHtml = lansiaList.map((item, i) => `
        <div class="lansia-item-row" data-idx="${i}" style="background-color: var(--bg-white); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px; margin-bottom: 8px; display: flex; align-items: center; justify-content: space-between; cursor: pointer; transition: all 0.2s ease;">
          <div style="display: flex; align-items: center; gap: 10px;">
            <div style="width: 28px; height: 28px; border-radius: 50%; background-color: ${genderFilter === 'male' ? '#DBEAFE' : genderFilter === 'female' ? '#FCE7F3' : '#FEF3C7'}; color: ${genderFilter === 'male' ? '#1E40AF' : genderFilter === 'female' ? '#9D174D' : '#D97706'}; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700;">
              ${i + 1}
            </div>
            <div>
              <div style="font-size: 13px; font-weight: 700; color: var(--text-main); display: flex; align-items: center; gap: 6px;">
                ${item.nama}
                <span style="font-size: 10px; font-weight: 600; padding: 2px 6px; border-radius: 4px; background-color: ${item.jenis_kelamin.includes('Perempuan') || item.jenis_kelamin === 'P' ? '#FCE7F3' : '#DBEAFE'}; color: ${item.jenis_kelamin.includes('Perempuan') || item.jenis_kelamin === 'P' ? '#BE185D' : '#1E40AF'};">${item.jenis_kelamin || '-'}</span>
                <i data-lucide="chevron-right" style="width: 14px; height: 14px; color: var(--primary);"></i>
              </div>
              <div style="font-size: 11px; color: var(--text-muted); font-family: monospace;">NIK: ${item.nik || '-'} • Lahir: ${item.tanggal_lahir || '-'} (${item.tempat_lahir || '-'})</div>
            </div>
          </div>
          <div style="text-align: right; display: flex; align-items: center; gap: 8px;">
            <span style="font-size: 12px; font-weight: 700; background-color: ${genderFilter === 'male' ? '#DBEAFE' : genderFilter === 'female' ? '#FCE7F3' : '#FEF3C7'}; color: ${genderFilter === 'male' ? '#1E40AF' : genderFilter === 'female' ? '#9D174D' : '#B45309'}; padding: 4px 10px; border-radius: 100px;">
              ${item.umur > 0 ? item.umur + ' Thn' : '≥ 60 Thn'}
            </span>
            <span style="font-size: 11px; color: var(--primary); font-weight: 600;">Lihat Detail →</span>
          </div>
        </div>
      `).join("");
    }
    
    wrapper.innerHTML = `
      <div class="result-card" id="${cardId}" style="border-color: ${badgeColor};">
        <!-- VIEW 1: DAFTAR LANSIA -->
        <div class="lansia-list-view">
          <div class="result-card-header" style="background-color: ${headerBg};">
            <div class="result-card-title" style="color: ${titleColor};">
              <i data-lucide="${iconName}" style="width: 16px; height: 16px; color: ${badgeColor};"></i>
              ${cardTitle}
            </div>
            <div style="display: flex; gap: 6px; align-items: center;">
              <span class="result-badge" style="background-color: ${badgeColor}; color: white;">${badgeLabel}</span>
              <button class="btn-export-lansia-excel" style="background-color: #ECFDF5; border: 1px solid #A7F3D0; color: #047857; border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 4px;">📥 Excel</button>
              <button class="btn-export-lansia-pdf" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; color: #1D4ED8; border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 4px;">📄 PDF</button>
            </div>
          </div>
          
          <div style="padding: 16px 20px;">
            <div style="display: flex; align-items: center; justify-content: space-between; padding-bottom: 12px; margin-bottom: 14px; border-bottom: 1px solid var(--border);">
              <div>
                <div style="font-size: 11px; font-weight: 700; color: var(--text-muted); text-transform: uppercase;">PROPORSI POPULASI LANSIA</div>
                <div style="font-size: 18px; font-weight: 700; color: ${titleColor};">${countLansia} dari ${totalDataset} Warga (${pctLansia}%)</div>
              </div>
              <div style="font-size: 11px; background-color: var(--bg-light); border: 1px solid var(--border); padding: 6px 10px; border-radius: 6px; color: var(--text-main);">
                Kategori: <strong>Usia ≥ 60 Tahun</strong>
              </div>
            </div>

            <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
              <i data-lucide="users" style="width: 14px; height: 14px; color: #D97706;"></i>
              Daftar Lengkap Warga Lansia (Klik untuk Detail):
            </div>
            
            <div style="max-height: 280px; overflow-y: auto; padding-right: 4px;">
              ${lansiaRowsHtml}
            </div>
          </div>

          <div class="performance-metrics-bar" style="background-color: #FFFBEB; font-size: 11px; padding: 10px 20px; color: #92400E; border-top: 1px solid #FDE68A;">
            <span>Perhitungan tanggal lahir & usia secara otomatis berdasarkan dataset aktif Kelurahan Cipaganti.</span>
          </div>
        </div>

        <!-- VIEW 2: DETAIL WARGA (Hidden initially) -->
        <div class="lansia-detail-view" style="display: none;">
          <div class="result-card-header" style="background-color: var(--primary-light);">
            <button class="btn-back-lansia" style="background: white; border: 1px solid var(--border); padding: 6px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; color: var(--primary); cursor: pointer; display: flex; align-items: center; gap: 6px;">
              <i data-lucide="arrow-left" style="width: 14px; height: 14px;"></i>
              ← Kembali ke Daftar Lansia
            </button>
            <span class="result-badge" style="background-color: var(--primary); color: white;">Detail Lansia</span>
          </div>

          <div class="detail-content-area" style="padding: 16px 20px;">
            <!-- Rendered dynamically on click -->
          </div>

          <div class="performance-metrics-bar" style="background-color: var(--bg-light); font-size: 11px; padding: 10px 20px; color: var(--text-muted); border-top: 1px solid var(--border);">
            <button class="btn-back-lansia-bottom" style="background: transparent; border: none; font-size: 12px; font-weight: 600; color: var(--primary); cursor: pointer; display: flex; align-items: center; gap: 6px;">
              ← Kembali ke Daftar Lengkap Lansia
            </button>
          </div>
        </div>

      </div>
      <div class="chat-time">${timeStr}</div>
    `;
    
    chatBody.appendChild(wrapper);
    if (typeof lucide !== "undefined") {
      lucide.createIcons();
    }

    // Attach Click Handlers
    const cardEl = document.getElementById(cardId);
    if (cardEl) {
      const listView = cardEl.querySelector('.lansia-list-view');
      const detailView = cardEl.querySelector('.lansia-detail-view');
      const detailContentArea = cardEl.querySelector('.detail-content-area');
      const itemRows = cardEl.querySelectorAll('.lansia-item-row');
      const backBtns = cardEl.querySelectorAll('.btn-back-lansia, .btn-back-lansia-bottom');

      itemRows.forEach(rowEl => {
        rowEl.addEventListener('click', () => {
          const idx = parseInt(rowEl.getAttribute('data-idx'));
          const w = lansiaList[idx];
          if (!w) return;

          // Render detail content inside View 2
          detailContentArea.innerHTML = `
            <div style="font-size: 14px; font-weight: 700; color: var(--text-main); margin-bottom: 12px; display: flex; align-items: center; gap: 8px;">
              <i data-lucide="user-check" style="width: 18px; height: 18px; color: var(--primary);"></i>
              Rincian Biodata Lansia: ${w.nama}
            </div>

            <div class="result-card-grid" style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Nama Lengkap</div>
                <div class="result-value" style="font-size: 13px; font-weight: 700; color: var(--text-main);">${w.nama}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">NIK</div>
                <div class="result-value" style="font-size: 13px; font-weight: 700; color: var(--text-main);">${w.nik || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Nomor KK</div>
                <div class="result-value" style="font-size: 13px; font-weight: 700; color: var(--text-main);">${w.kk || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Usia Lansia</div>
                <div class="result-value" style="font-size: 13px; font-weight: 700; color: #B45309;">${w.umur > 0 ? w.umur + ' Tahun' : '≥ 60 Tahun'}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Jenis Kelamin</div>
                <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.jenis_kelamin || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tempat Lahir</div>
                <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tempat_lahir || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tanggal Lahir</div>
                <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tanggal_lahir || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Agama</div>
                <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.agama || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Pendidikan Terakhir</div>
                <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.pendidikan || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Pekerjaan</div>
                <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.pekerjaan || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Status Pernikahan</div>
                <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.status_pernikahan || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Hub. Keluarga (SDHK)</div>
                <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.sdhk || '-'}</div>
              </div>
              <div class="result-item">
                <div class="result-label" style="font-size: 11px; color: var(--text-muted);">RT / RW</div>
                <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">RT ${w.rt || '01'} / RW ${w.rw || '02'}</div>
              </div>
            </div>
          `;

          if (typeof lucide !== "undefined") {
            lucide.createIcons();
          }

          // Switch views
          listView.style.display = 'none';
          detailView.style.display = 'block';
          
    const lansiaCardEl = document.getElementById(cardId);
    if (lansiaCardEl) {
      const btnLansiaExcel = lansiaCardEl.querySelector('.btn-export-lansia-excel');
      const btnLansiaPdf = lansiaCardEl.querySelector('.btn-export-lansia-pdf');
      if (btnLansiaExcel) {
        btnLansiaExcel.addEventListener('click', (e) => { e.stopPropagation(); window.cipabotExportExcel(lansiaList, 'data_warga_lansia.xlsx'); });
      }
      if (btnLansiaPdf) {
        btnLansiaPdf.addEventListener('click', (e) => { e.stopPropagation(); window.cipabotExportPdf('Daftar Warga Lansia (Usia >= 60 Tahun)', lansiaList); });
      }
    }

    scrollToBottom();
        });
      });

      backBtns.forEach(btn => {
        btn.addEventListener('click', () => {
          detailView.style.display = 'none';
          listView.style.display = 'block';
          scrollToBottom();
        });
      });
    }

    scrollToBottom();
  }

  function appendBotLansiaPerRtCard(totalDataset, allLansia) {
    const wrapper = document.createElement("div");
    wrapper.className = "chat-bubble-wrapper bot";
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    const cardId = 'lansia-rt-card-' + Math.random().toString(36).substring(2, 9);
    
    // Group lansia by RT
    const rtMap = {};
    allLansia.forEach(w => {
      let rtKey = "RT 01";
      if (w.rt) {
        const num = String(w.rt).replace(/\D/g, '');
        rtKey = num ? `RT ${num.padStart(2, '0')}` : `RT ${w.rt}`;
      }
      if (!rtMap[rtKey]) {
        rtMap[rtKey] = { total: 0, laki: 0, perempuan: 0, list: [] };
      }
      rtMap[rtKey].total += 1;
      const jk = (w.jenis_kelamin || '').toLowerCase();
      if (jk.includes("laki") || jk === "pria" || jk === "l") {
        rtMap[rtKey].laki += 1;
      } else {
        rtMap[rtKey].perempuan += 1;
      }
      rtMap[rtKey].list.push(w);
    });

    const sortedRtKeys = Object.keys(rtMap).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

    let maxRtKey = "";
    let maxCount = -1;
    sortedRtKeys.forEach(k => {
      if (rtMap[k].total > maxCount) {
        maxCount = rtMap[k].total;
        maxRtKey = k;
      }
    });

    const totalLansia = allLansia.length;

    const rtBarsHtml = sortedRtKeys.map(rtKey => {
      const data = rtMap[rtKey];
      const pct = totalLansia > 0 ? ((data.total / totalLansia) * 100).toFixed(1) : 0;
      return `
        <div class="rt-item-row" data-rt="${rtKey}" style="background-color: var(--bg-white); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px; margin-bottom: 8px; cursor: pointer; transition: all 0.2s ease;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
            <span style="font-size: 13px; font-weight: 700; color: var(--text-main); display: flex; align-items: center; gap: 6px;">
              <i data-lucide="map-pin" style="width: 14px; height: 14px; color: #D97706;"></i>
              ${rtKey} / RW 02
              <i data-lucide="chevron-right" style="width: 14px; height: 14px; color: var(--primary);"></i>
            </span>
            <span style="font-size: 12px; font-weight: 700; color: #B45309; background-color: #FEF3C7; padding: 2px 8px; border-radius: 100px;">
              ${data.total} Lansia (${pct}%) →
            </span>
          </div>
          <div style="width: 100%; height: 6px; background-color: var(--bg-light); border-radius: 100px; overflow: hidden; margin-bottom: 6px;">
            <div style="width: ${pct}%; height: 100%; background-color: #D97706; border-radius: 100px;"></div>
          </div>
          <div style="font-size: 11px; color: var(--text-muted); display: flex; gap: 12px;">
            <span>👨 <strong>${data.laki}</strong> Laki-laki</span>
            <span>👩 <strong>${data.perempuan}</strong> Perempuan</span>
          </div>
        </div>
      `;
    }).join("");

    wrapper.innerHTML = `
      <div class="chat-bubble bot">
        Berikut adalah rekapitulasi sebaran warga lansia (usia ≥ 60 tahun) <strong>per RT di wilayah RW 02 Kelurahan Cipaganti</strong> (Klik RT untuk melihat daftar nama warga):
      </div>
      <div class="result-card" id="${cardId}" style="border-color: #F59E0B;">
        
        <!-- LEVEL 1: SUMMARY PER RT -->
        <div class="rt-summary-view">
          <div class="result-card-header" style="background-color: #FEF3C7;">
            <div class="result-card-title" style="color: #B45309;">
              <i data-lucide="map-pin" style="width: 16px; height: 16px; color: #D97706;"></i>
              Sebaran Warga Lansia per RT (RW 02)
            </div>
            <span class="result-badge" style="background-color: #D97706; color: white;">${sortedRtKeys.length} RT Terdata</span>
          </div>

          <div style="padding: 16px 20px;">
            <div style="display: flex; align-items: center; justify-content: space-between; padding-bottom: 12px; margin-bottom: 14px; border-bottom: 1px solid var(--border);">
              <div>
                <div style="font-size: 11px; font-weight: 700; color: var(--text-muted); text-transform: uppercase;">TOTAL LANSIA RW 02</div>
                <div style="font-size: 18px; font-weight: 700; color: #B45309;">${totalLansia} Warga Lansia</div>
              </div>
              ${maxRtKey ? `
                <div style="font-size: 11px; background-color: #FEF3C7; border: 1px solid #FDE68A; padding: 6px 10px; border-radius: 6px; color: #B45309; text-align: right;">
                  Lansia Terbanyak: <strong>${maxRtKey} (${maxCount} lansia)</strong>
                </div>
              ` : ''}
            </div>

            <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
              <i data-lucide="mouse-pointer-click" style="width: 14px; height: 14px; color: var(--primary);"></i>
              Klik salah satu RT di bawah untuk melihat daftar nama warga lansianya:
            </div>

            <div style="max-height: 320px; overflow-y: auto; padding-right: 4px;">
              ${rtBarsHtml}
            </div>
          </div>

          <div class="performance-metrics-bar" style="background-color: var(--bg-light); font-size: 11px; padding: 10px 20px; color: var(--text-muted); border-top: 1px solid var(--border);">
            <span>Diklik untuk membuka daftar warga lansia per RT.</span>
          </div>
        </div>

        <!-- LEVEL 2: DAFTAR NAMA WARGA LANSIA DI RT DIPIILIH -->
        <div class="rt-citizens-view" style="display: none;">
          <div class="result-card-header" style="background-color: #FEF3C7;">
            <button class="btn-back-to-rt-summary" style="background: white; border: 1px solid #FDE68A; padding: 6px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; color: #B45309; cursor: pointer; display: flex; align-items: center; gap: 6px;">
              <i data-lucide="arrow-left" style="width: 14px; height: 14px;"></i>
              ← Kembali ke Sebaran RT
            </button>
            <span class="rt-badge-header result-badge" style="background-color: #D97706; color: white;">RT 01</span>
          </div>

          <div style="padding: 16px 20px;">
            <div class="rt-title-header" style="font-size: 14px; font-weight: 700; color: #B45309; margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
              <!-- Dynamic Title -->
            </div>
            <div class="rt-citizen-list-container" style="max-height: 300px; overflow-y: auto; padding-right: 4px;">
              <!-- Dynamic citizen list -->
            </div>
          </div>
        </div>

        <!-- LEVEL 3: DETAIL BIODATA WARGA LANSIA -->
        <div class="rt-citizen-detail-view" style="display: none;">
          <div class="result-card-header" style="background-color: var(--primary-light);">
            <button class="btn-back-to-rt-citizens" style="background: white; border: 1px solid var(--border); padding: 6px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; color: var(--primary); cursor: pointer; display: flex; align-items: center; gap: 6px;">
              <i data-lucide="arrow-left" style="width: 14px; height: 14px;"></i>
              ← Kembali ke Daftar Lansia RT
            </button>
            <span class="result-badge" style="background-color: var(--primary); color: white;">Terverifikasi</span>
          </div>

          <div class="rt-detail-content" style="padding: 16px 20px;">
            <!-- Dynamic full biodata -->
          </div>

          <div class="performance-metrics-bar" style="background-color: var(--bg-light); font-size: 11px; padding: 10px 20px; color: var(--text-muted); border-top: 1px solid var(--border);">
            <button class="btn-back-to-rt-citizens-bottom" style="background: transparent; border: none; font-size: 12px; font-weight: 600; color: var(--primary); cursor: pointer; display: flex; align-items: center; gap: 6px;">
              ← Kembali ke Daftar Lansia RT
            </button>
          </div>
        </div>

      </div>
      <div class="chat-time">${timeStr}</div>
    `;

    chatBody.appendChild(wrapper);
    if (typeof lucide !== "undefined") {
      lucide.createIcons();
    }

    // Attach level 1 -> level 2 -> level 3 event listeners
    const cardEl = document.getElementById(cardId);
    if (cardEl) {
      const summaryView = cardEl.querySelector('.rt-summary-view');
      const citizensView = cardEl.querySelector('.rt-citizens-view');
      const detailView = cardEl.querySelector('.rt-citizen-detail-view');
      
      const rtBadgeHeader = cardEl.querySelector('.rt-badge-header');
      const rtTitleHeader = cardEl.querySelector('.rt-title-header');
      const citizenListContainer = cardEl.querySelector('.rt-citizen-list-container');
      const detailContentArea = cardEl.querySelector('.rt-detail-content');

      const btnBackSummary = cardEl.querySelector('.btn-back-to-rt-summary');
      const btnBackCitizens = cardEl.querySelectorAll('.btn-back-to-rt-citizens, .btn-back-to-rt-citizens-bottom');

      let currentActiveRt = "";
      let currentRtList = [];

      // Level 1: Click RT row
      cardEl.querySelectorAll('.rt-item-row').forEach(row => {
        row.addEventListener('click', () => {
          const rtKey = row.getAttribute('data-rt');
          const data = rtMap[rtKey];
          if (!data) return;

          currentActiveRt = rtKey;
          currentRtList = data.list;

          rtBadgeHeader.textContent = `${rtKey} / RW 02`;
          rtTitleHeader.innerHTML = `
            <div style="display: flex; align-items: center; justify-content: space-between; width: 100%;">
              <div style="display: flex; align-items: center; gap: 6px;">
                <i data-lucide="users" style="width: 16px; height: 16px; color: #D97706;"></i>
                <span>Daftar Lansia di ${rtKey} (${data.total} Warga)</span>
              </div>
              <div style="display: flex; gap: 4px;">
                <button class="btn-export-rt-excel" style="background-color: #ECFDF5; border: 1px solid #A7F3D0; color: #047857; border-radius: 4px; padding: 2px 6px; font-size: 10px; font-weight: 700; cursor: pointer;">📥 Excel</button>
                <button class="btn-export-rt-pdf" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; color: #1D4ED8; border-radius: 4px; padding: 2px 6px; font-size: 10px; font-weight: 700; cursor: pointer;">📄 PDF</button>
              </div>
            </div>
          `;

          const btnRtExcel = rtTitleHeader.querySelector('.btn-export-rt-excel');
          const btnRtPdf = rtTitleHeader.querySelector('.btn-export-rt-pdf');
          if (btnRtExcel) btnRtExcel.addEventListener('click', (e) => { e.stopPropagation(); if (typeof XLSX === 'undefined') { alert('XLSX belum dimuat.'); return; } window.cipabotExportExcel(data.list, 'data_lansia_' + rtKey.replace(' ', '_') + '.xlsx'); });
          if (btnRtPdf) btnRtPdf.addEventListener('click', (e) => { e.stopPropagation(); window.cipabotExportPdf('Daftar Lansia ' + rtKey, data.list); });

          citizenListContainer.innerHTML = data.list.map((item, i) => `
            <div class="rt-citizen-item-row" data-idx="${i}" style="background-color: var(--bg-white); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px; margin-bottom: 8px; display: flex; align-items: center; justify-content: space-between; cursor: pointer; transition: all 0.2s ease;">
              <div style="display: flex; align-items: center; gap: 10px;">
                <div style="width: 28px; height: 28px; border-radius: 50%; background-color: #FEF3C7; color: #B45309; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700;">
                  ${i + 1}
                </div>
                <div>
                  <div style="font-size: 13px; font-weight: 700; color: var(--text-main); display: flex; align-items: center; gap: 6px;">
                    ${item.nama}
                    <span style="font-size: 10px; font-weight: 600; padding: 2px 6px; border-radius: 4px; background-color: ${item.jenis_kelamin.includes('Perempuan') || item.jenis_kelamin === 'P' ? '#FCE7F3' : '#DBEAFE'}; color: ${item.jenis_kelamin.includes('Perempuan') || item.jenis_kelamin === 'P' ? '#BE185D' : '#1E40AF'};">${item.jenis_kelamin || '-'}</span>
                    <i data-lucide="chevron-right" style="width: 14px; height: 14px; color: var(--primary);"></i>
                  </div>
                  <div style="font-size: 11px; color: var(--text-muted); font-family: monospace;">NIK: ${item.nik || '-'} • Umur: ${item.umur > 0 ? item.umur + ' Thn' : '≥ 60 Thn'} • Lahir: ${item.tanggal_lahir || '-'}</div>
                </div>
              </div>
              <div style="text-align: right;">
                <span style="font-size: 11px; color: var(--primary); font-weight: 700;">Lihat Biodata →</span>
              </div>
            </div>
          `).join("");

          if (typeof lucide !== "undefined") lucide.createIcons();

          // Level 2: Click citizen row to see detail
          citizenListContainer.querySelectorAll('.rt-citizen-item-row').forEach(cRow => {
            cRow.addEventListener('click', () => {
              const idx = parseInt(cRow.getAttribute('data-idx'));
              const w = currentRtList[idx];
              if (!w) return;

              detailContentArea.innerHTML = `
                <div style="font-size: 14px; font-weight: 700; color: var(--text-main); margin-bottom: 12px; display: flex; align-items: center; gap: 8px;">
                  <i data-lucide="user-check" style="width: 18px; height: 18px; color: var(--primary);"></i>
                  Rincian Biodata Warga Lansia: ${w.nama}
                </div>

                <div class="result-card-grid" style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Nama Lengkap</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 700; color: var(--text-main);">${w.nama}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">NIK</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 700; color: var(--text-main);">${w.nik || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Nomor KK</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 700; color: var(--text-main);">${w.kk || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Usia Lansia</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 700; color: #B45309;">${w.umur > 0 ? w.umur + ' Tahun' : '≥ 60 Tahun'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Jenis Kelamin</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.jenis_kelamin || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tempat Lahir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tempat_lahir || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tanggal Lahir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tanggal_lahir || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Agama</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.agama || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Pendidikan Terakhir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.pendidikan || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Pekerjaan</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.pekerjaan || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Status Pernikahan</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.status_pernikahan || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Hub. Keluarga (SDHK)</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.sdhk || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">RT / RW</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">RT ${w.rt || '01'} / RW ${w.rw || '02'}</div>
                  </div>
                  ${w.alamat ? `<div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Alamat Lengkap</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.alamat}</div>
                  </div>` : ''}
                </div>
              `;

              if (typeof lucide !== "undefined") lucide.createIcons();

              citizensView.style.display = 'none';
              detailView.style.display = 'block';
              scrollToBottom();
            });
          });

          summaryView.style.display = 'none';
          citizensView.style.display = 'block';
          scrollToBottom();
        });
      });

      // Back buttons
      btnBackSummary.addEventListener('click', () => {
        citizensView.style.display = 'none';
        summaryView.style.display = 'block';
        scrollToBottom();
      });

      btnBackCitizens.forEach(b => {
        b.addEventListener('click', () => {
          detailView.style.display = 'none';
          citizensView.style.display = 'block';
          scrollToBottom();
        });
      });
    }

    scrollToBottom();
  }


  function appendBotPekerjaanCard(totalDataset, allWarga, genderFilter = "all") {
    const wrapper = document.createElement("div");
    wrapper.className = "chat-bubble-wrapper bot";
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    const cardId = 'pek-card-' + Math.random().toString(36).substring(2, 9);
    
    // Filter by gender
    let filteredWarga = allWarga;
    if (genderFilter === "male") {
      filteredWarga = allWarga.filter(w => {
        const jk = (w.jenis_kelamin || '').toLowerCase();
        return jk.includes("laki") || jk === "pria" || jk === "l";
      });
    } else if (genderFilter === "female") {
      filteredWarga = allWarga.filter(w => {
        const jk = (w.jenis_kelamin || '').toLowerCase();
        return jk.includes("perempuan") || jk === "wanita" || jk === "p";
      });
    }

    // Group by Pekerjaan (Categorized)
    const pekMap = {};
    filteredWarga.forEach(w => {
      let raw = w.pekerjaan ? String(w.pekerjaan).trim() : "Belum/Tidak Terdata";
      if (!raw || raw === "-" || /^\d+$/.test(raw) || raw.toLowerCase() === "pekerjaan") {
        raw = "Belum/Tidak Terdata";
      }
      
      const lw = raw.toLowerCase();
      let pekKey = "";
      if (lw.includes("pelajar") || lw.includes("mahasiswa") || lw.includes("siswa")) pekKey = "Pelajar / Mahasiswa";
      else if (lw.includes("pns") || lw.includes("pegawai negeri") || lw.includes("tni") || lw.includes("polri")) pekKey = "PNS / TNI / POLRI";
      else if (lw.includes("swasta") || lw.includes("karyawan") || lw.includes("buruh") || lw.includes("pegawai") || lw.includes("staf") || lw.includes("marketing") || lw.includes("sales")) pekKey = "Karyawan / Buruh Swasta";
      else if (lw.includes("wiraswasta") || lw.includes("dagang") || lw.includes("usaha") || lw.includes("wirausaha") || lw.includes("wira swasta")) pekKey = "Wiraswasta / Pekerja Mandiri";
      else if (lw.includes("rumah tangga") || lw.includes("irt") || lw.includes("urus") || lw.includes("ibu rumah")) pekKey = "Mengurus Rumah Tangga";
      else if (lw.includes("belum") || lw.includes("tidak bekerja") || lw.includes("pengangguran") || lw.includes("tidak terdata")) pekKey = "Belum / Tidak Bekerja";
      else if (lw.includes("pensiun") || lw.includes("purnawirawan")) pekKey = "Pensiunan";
      else if (lw.includes("guru") || lw.includes("dosen") || lw.includes("pengajar")) pekKey = "Pendidik (Guru/Dosen)";
      else if (lw.includes("medis") || lw.includes("dokter") || lw.includes("perawat") || lw.includes("bidan") || lw.includes("kesehatan")) pekKey = "Tenaga Kesehatan";
      else {
        // Fallback to Title Case
        pekKey = raw.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
      }

      if (!pekMap[pekKey]) {
        pekMap[pekKey] = { total: 0, laki: 0, perempuan: 0, list: [] };
      }
      pekMap[pekKey].total += 1;
      const jk = (w.jenis_kelamin || '').toLowerCase();
      if (jk.includes("laki") || jk === "pria" || jk === "l") {
        pekMap[pekKey].laki += 1;
      } else {
        pekMap[pekKey].perempuan += 1;
      }
      pekMap[pekKey].list.push(w);
    });

    const sortedPekKeys = Object.keys(pekMap).sort((a, b) => pekMap[b].total - pekMap[a].total);

    let maxPekKey = sortedPekKeys.length > 0 ? sortedPekKeys[0] : "";
    let maxCount = maxPekKey ? pekMap[maxPekKey].total : 0;
    const totalWarga = filteredWarga.length;

    const pekBarsHtml = sortedPekKeys.map(pekKey => {
      const data = pekMap[pekKey];
      const pct = totalWarga > 0 ? ((data.total / totalWarga) * 100).toFixed(1) : 0;
      return `
        <div class="pek-item-row" data-pek="${pekKey}" style="background-color: var(--bg-white); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px; margin-bottom: 8px; cursor: pointer; transition: all 0.2s ease;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
            <span style="font-size: 13px; font-weight: 700; color: var(--text-main); display: flex; align-items: center; gap: 6px;">
              <i data-lucide="briefcase" style="width: 14px; height: 14px; color: #059669;"></i>
              ${pekKey}
              <i data-lucide="chevron-right" style="width: 14px; height: 14px; color: var(--primary);"></i>
            </span>
            <span style="font-size: 12px; font-weight: 700; color: #047857; background-color: #D1FAE5; padding: 2px 8px; border-radius: 100px;">
              ${data.total} Warga (${pct}%) →
            </span>
          </div>
          <div style="width: 100%; height: 6px; background-color: var(--bg-light); border-radius: 100px; overflow: hidden; margin-bottom: 6px;">
            <div style="width: ${pct}%; height: 100%; background-color: #059669; border-radius: 100px;"></div>
          </div>
          <div style="font-size: 11px; color: var(--text-muted); display: flex; gap: 12px;">
            <span>👨 <strong>${data.laki}</strong> Laki-laki</span>
            <span>👩 <strong>${data.perempuan}</strong> Perempuan</span>
          </div>
        </div>
      `;
    }).join("");

    let titleSuffix = "";
    if (genderFilter === "male") titleSuffix = " (Laki-laki)";
    if (genderFilter === "female") titleSuffix = " (Perempuan)";

    wrapper.innerHTML = `
      <div class="chat-bubble bot">
        Berikut adalah rekapitulasi sebaran <strong>Status Pekerjaan / Kategori Profesi${titleSuffix}</strong> warga (Klik profesi untuk melihat daftar nama warga):
      </div>
      <div class="result-card" id="${cardId}" style="border-color: #10B981;">
        
        <!-- LEVEL 1: SUMMARY PER PEKERJAAN -->
        <div class="pek-summary-view">
          <div class="result-card-header" style="background-color: #D1FAE5;">
            <div class="result-card-title" style="color: #047857;">
              <i data-lucide="briefcase" style="width: 16px; height: 16px; color: #059669;"></i>
              Statistik Kategori Pekerjaan
            </div>
            <div style="display: flex; gap: 6px; align-items: center;">
              <span class="result-badge" style="background-color: #059669; color: white;">${sortedPekKeys.length} Kategori</span>
              <button class="btn-export-pek-excel" style="background-color: #ECFDF5; border: 1px solid #A7F3D0; color: #047857; border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 4px;">📥 Excel</button>
              <button class="btn-export-pek-pdf" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; color: #1D4ED8; border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 4px;">📄 PDF</button>
            </div>
          </div>

          <div style="padding: 16px 20px;">
            <div style="display: flex; align-items: center; justify-content: space-between; padding-bottom: 12px; margin-bottom: 14px; border-bottom: 1px solid var(--border);">
              <div>
                <div style="font-size: 11px; font-weight: 700; color: var(--text-muted); text-transform: uppercase;">TOTAL DATA WARGA</div>
                <div style="font-size: 18px; font-weight: 700; color: #047857;">${totalWarga} Warga</div>
              </div>
              ${maxPekKey ? `
                <div style="font-size: 11px; background-color: #D1FAE5; border: 1px solid #A7F3D0; padding: 6px 10px; border-radius: 6px; color: #047857; text-align: right; max-width: 140px; line-height: 1.3;">
                  Terbanyak: <strong>${maxPekKey} (${maxCount})</strong>
                </div>
              ` : ''}
            </div>

            <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
              <i data-lucide="mouse-pointer-click" style="width: 14px; height: 14px; color: var(--primary);"></i>
              Klik salah satu kategori di bawah untuk melihat warganya:
            </div>

            <div style="max-height: 320px; overflow-y: auto; padding-right: 4px;">
              ${pekBarsHtml}
            </div>
          </div>
          <div class="performance-metrics-bar" style="background-color: var(--bg-light); font-size: 11px; padding: 10px 20px; color: var(--text-muted); border-top: 1px solid var(--border);">
            <span>Diklik untuk membuka daftar warga per profesi.</span>
          </div>
        </div>

        <!-- LEVEL 2: CITIZEN LIST BY PEKERJAAN -->
        <div class="pek-citizens-view" style="display: none;">
          <div class="result-card-header" style="background-color: var(--bg-light); border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; padding: 12px 20px;">
            <button class="btn-back-summary" style="background: none; border: none; color: var(--text-muted); font-size: 12px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 6px; padding: 4px 8px; border-radius: 6px; transition: background 0.2s;">
              <i data-lucide="arrow-left" style="width: 14px; height: 14px;"></i>
              Kembali
            </button>
            <div class="pek-list-title" style="font-size: 13px; font-weight: 700; color: var(--text-main);"></div>
          </div>
          <div style="padding: 16px 20px; max-height: 400px; overflow-y: auto;">
            <div class="pek-list-container" style="display: flex; flex-direction: column; gap: 10px;">
              <!-- Insert list here -->
            </div>
          </div>
        </div>

        <!-- LEVEL 3: BIODATA DETAIL -->
        <div class="pek-detail-view" style="display: none;">
          <div class="result-card-header" style="background-color: var(--bg-light); border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; padding: 12px 20px;">
            <button class="btn-back-citizens" style="background: none; border: none; color: var(--text-muted); font-size: 12px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 6px; padding: 4px 8px; border-radius: 6px; transition: background 0.2s;">
              <i data-lucide="arrow-left" style="width: 14px; height: 14px;"></i>
              Daftar Warga
            </button>
            <div style="font-size: 13px; font-weight: 700; color: var(--text-main);">Detail Biodata</div>
          </div>
          <div class="pek-detail-container" style="padding: 20px;">
            <!-- Biodata -->
          </div>
        </div>

      </div>
      <div class="chat-time">${timeStr}</div>
    `;

    chatBody.appendChild(wrapper);
    if (typeof lucide !== "undefined") lucide.createIcons();

    // Event Listeners for Interaction
    const cardEl = document.getElementById(cardId);
    if (cardEl) {
      const summaryView = cardEl.querySelector('.pek-summary-view');
      const citizensView = cardEl.querySelector('.pek-citizens-view');
      const detailView = cardEl.querySelector('.pek-detail-view');
      
      const btnBackSummary = cardEl.querySelector('.btn-back-summary');
      const listTitle = cardEl.querySelector('.pek-list-title');
      const listContainer = cardEl.querySelector('.pek-list-container');
      const detailContainer = cardEl.querySelector('.pek-detail-container');
      const btnBackCitizensList = cardEl.querySelectorAll('.btn-back-citizens');

      // 1. Click on RT/Pekerjaan Row -> Show Citizens
      cardEl.querySelectorAll('.pek-item-row').forEach(row => {
        row.addEventListener('click', () => {
          const pekKey = row.getAttribute('data-pek');
          const data = pekMap[pekKey];
          
          listTitle.innerHTML = `<div style="display: flex; align-items: center; gap: 8px;">
            <span>${pekKey} (${data.total} Warga)</span>
            <button class="btn-export-item-excel" style="background-color: #ECFDF5; border: 1px solid #A7F3D0; color: #047857; border-radius: 4px; padding: 2px 6px; font-size: 10px; font-weight: 700; cursor: pointer;">📥 Excel</button>
            <button class="btn-export-item-pdf" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; color: #1D4ED8; border-radius: 4px; padding: 2px 6px; font-size: 10px; font-weight: 700; cursor: pointer;">📄 PDF</button>
          </div>`;

          const btnExcelPek = listTitle.querySelector('.btn-export-item-excel');
          const btnPdfPek = listTitle.querySelector('.btn-export-item-pdf');
          if (btnExcelPek) btnExcelPek.addEventListener('click', (e) => { e.stopPropagation(); if (typeof XLSX === 'undefined') { alert('XLSX belum dimuat.'); return; } window.cipabotExportExcel(data.list, 'data_pekerjaan_' + pekKey + '.xlsx'); });
          if (btnPdfPek) btnPdfPek.addEventListener('click', (e) => { e.stopPropagation(); window.cipabotExportPdf('Data Pekerjaan: ' + pekKey, data.list); });
          
          listContainer.innerHTML = data.list.map((w, idx) => {
            const isMale = (w.jenis_kelamin || '').toLowerCase().includes("laki") || (w.jenis_kelamin || '').toLowerCase() === "l";
            const jkBadge = isMale 
              ? `<span style="background-color: #DBEAFE; color: #1D4ED8; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700;">Laki-laki</span>`
              : `<span style="background-color: #FCE7F3; color: #BE185D; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700;">Perempuan</span>`;
            
            return `
              <div class="citizen-list-item" data-idx="${idx}" data-pek="${pekKey}" style="border: 1px solid var(--border); border-radius: 8px; padding: 12px; display: flex; justify-content: space-between; align-items: center; background-color: var(--bg-white); cursor: pointer; transition: border-color 0.2s;">
                <div style="display: flex; gap: 12px; align-items: center;">
                  <div style="width: 24px; height: 24px; border-radius: 12px; background-color: var(--bg-light); color: var(--text-muted); font-size: 11px; font-weight: 700; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                    ${idx + 1}
                  </div>
                  <div>
                    <div style="font-size: 13px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">${w.nama} ${jkBadge}</div>
                    <div style="font-size: 11px; color: var(--text-muted); display: flex; gap: 8px; align-items: center;">
                      <span>NIK: ${w.nik || '-'}</span>
                      <span style="color: #cbd5e1;">|</span>
                      <span>RT ${w.rt || '01'}/RW ${w.rw || '02'}</span>
                    </div>
                  </div>
                </div>
                <i data-lucide="chevron-right" style="width: 16px; height: 16px; color: var(--text-muted);"></i>
              </div>
            `;
          }).join("");

          if (typeof lucide !== "undefined") lucide.createIcons();

          // 2. Click on Citizen -> Show Detail
          cardEl.querySelectorAll('.citizen-list-item').forEach(item => {
            item.addEventListener('click', () => {
              const clickedPek = item.getAttribute('data-pek');
              const clickedIdx = parseInt(item.getAttribute('data-idx'));
              const w = pekMap[clickedPek].list[clickedIdx];
              
              detailContainer.innerHTML = `
                <div style="display: flex; align-items: center; gap: 16px; margin-bottom: 20px; padding-bottom: 16px; border-bottom: 1px dashed var(--border);">
                  <div class="avatar-circle" style="width: 48px; height: 48px; background-color: var(--primary-light); color: var(--primary); border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 20px; font-weight: 700;">
                    ${w.nama ? w.nama.charAt(0).toUpperCase() : '?'}
                  </div>
                  <div>
                    <div style="font-size: 16px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">${w.nama}</div>
                    <div style="display: flex; gap: 8px;">
                      <span class="result-badge" style="background-color: var(--bg-light); color: var(--text-muted);">NIK: ${w.nik || '-'}</span>
                      ${w.status ? `<span class="result-badge" style="background-color: var(--success-light); color: var(--success);">${w.status}</span>` : ''}
                    </div>
                  </div>
                </div>

                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Nomor KK</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.kk || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Jenis Kelamin</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.jenis_kelamin || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tempat Lahir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tempat_lahir || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tanggal Lahir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tanggal_lahir || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Agama</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.agama || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Pendidikan Terakhir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.pendidikan || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Pekerjaan</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.pekerjaan || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Status Pernikahan</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.status_pernikahan || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Hub. Keluarga (SDHK)</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.sdhk || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">RT / RW</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">RT ${w.rt || '01'} / RW ${w.rw || '02'}</div>
                  </div>
                  ${w.alamat ? `<div class="result-item" style="grid-column: 1 / -1;">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Alamat Lengkap</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.alamat}</div>
                  </div>` : ''}
                </div>
              `;

              if (typeof lucide !== "undefined") lucide.createIcons();

              citizensView.style.display = 'none';
              detailView.style.display = 'block';
              scrollToBottom();
            });
          });

          summaryView.style.display = 'none';
          citizensView.style.display = 'block';
          scrollToBottom();
        });
      });

      // Back buttons
      btnBackSummary.addEventListener('click', () => {
        citizensView.style.display = 'none';
        summaryView.style.display = 'block';
        scrollToBottom();
      });

      btnBackCitizensList.forEach(b => {
        b.addEventListener('click', () => {
          detailView.style.display = 'none';
          citizensView.style.display = 'block';
          scrollToBottom();
        });
      });

      // Export - Pekerjaan
      const btnPekExcel = cardEl.querySelector('.btn-export-pek-excel');
      const btnPekPdf = cardEl.querySelector('.btn-export-pek-pdf');
      if (btnPekExcel) btnPekExcel.addEventListener('click', (e) => { e.stopPropagation(); if (typeof XLSX === 'undefined') { alert('Library XLSX belum dimuat.'); return; } window.cipabotExportExcel(filteredWarga, 'data_pekerjaan_warga.xlsx'); });
      if (btnPekPdf) btnPekPdf.addEventListener('click', (e) => { e.stopPropagation(); window.cipabotExportPdf('Statistik Pekerjaan Warga Kelurahan Cipaganti', filteredWarga); });
    }

    scrollToBottom();
  }


  function appendBotPendidikanCard(totalDataset, allWarga, genderFilter = "all") {
    const wrapper = document.createElement("div");
    wrapper.className = "chat-bubble-wrapper bot";
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    const cardId = 'pnd-card-' + Math.random().toString(36).substring(2, 9);
    
    // Filter by gender
    let filteredWarga = allWarga;
    if (genderFilter === "male") {
      filteredWarga = allWarga.filter(w => {
        const jk = (w.jenis_kelamin || '').toLowerCase();
        return jk.includes("laki") || jk === "pria" || jk === "l";
      });
    } else if (genderFilter === "female") {
      filteredWarga = allWarga.filter(w => {
        const jk = (w.jenis_kelamin || '').toLowerCase();
        return jk.includes("perempuan") || jk === "wanita" || jk === "p";
      });
    }

    // Group by Pendidikan (Categorized)
    const pndMap = {};
    filteredWarga.forEach(w => {
      let raw = w.pendidikan ? String(w.pendidikan).trim() : "Belum/Tidak Terdata";
      if (!raw || raw === "-" || /^\d+$/.test(raw) || raw.toLowerCase() === "pendidikan" || raw.toLowerCase() === "pendidikan terakhir") {
        raw = "Belum/Tidak Terdata";
      }
      
      const lw = raw.toLowerCase();
      let pndKey = "";
      if (lw.includes("belum") || lw.includes("tidak sekolah") || lw.includes("tidak tamat")) pndKey = "Tidak/Belum Sekolah";
      else if (lw.includes("sd") || lw.includes("sekolah dasar") || lw.includes("mi") || lw.includes("madrasah ibtidaiyah") || lw.includes("tamat sd")) pndKey = "SD / Sederajat";
      else if (lw.includes("smp") || lw.includes("mts") || lw.includes("madrasah tsanawiyah") || lw.includes("slpt")) pndKey = "SMP / Sederajat";
      else if (lw.includes("sma") || lw.includes("smk") || lw.includes("slta") || lw.includes("stm") || lw.includes("ma") || lw.includes("madrasah aliyah")) pndKey = "SMA / Sederajat";
      else if (lw.includes("d1") || lw.includes("d2") || lw.includes("d3") || lw.includes("diploma")) pndKey = "Diploma (D1/D2/D3)";
      else if (lw.includes("s1") || lw.includes("d4") || lw.includes("sarjana")) pndKey = "Sarjana (S1/D4)";
      else if (lw.includes("s2") || lw.includes("magister")) pndKey = "Magister (S2)";
      else if (lw.includes("s3") || lw.includes("doktor")) pndKey = "Doktor (S3)";
      else if (lw.includes("terdata") || lw.includes("kosong")) pndKey = "Belum/Tidak Terdata";
      else {
        pndKey = raw.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
      }

      if (!pndMap[pndKey]) {
        pndMap[pndKey] = { total: 0, laki: 0, perempuan: 0, list: [] };
      }
      pndMap[pndKey].total += 1;
      const jk = (w.jenis_kelamin || '').toLowerCase();
      if (jk.includes("laki") || jk === "pria" || jk === "l") {
        pndMap[pndKey].laki += 1;
      } else {
        pndMap[pndKey].perempuan += 1;
      }
      pndMap[pndKey].list.push(w);
    });

    // Custom sort logic for Education levels
    const eduOrder = {
      "Doktor (S3)": 8,
      "Magister (S2)": 7,
      "Sarjana (S1/D4)": 6,
      "Diploma (D1/D2/D3)": 5,
      "SMA / Sederajat": 4,
      "SMP / Sederajat": 3,
      "SD / Sederajat": 2,
      "Tidak/Belum Sekolah": 1,
      "Belum/Tidak Terdata": 0
    };

    const sortedPndKeys = Object.keys(pndMap).sort((a, b) => {
      const orderA = eduOrder[a] !== undefined ? eduOrder[a] : -1;
      const orderB = eduOrder[b] !== undefined ? eduOrder[b] : -1;
      if (orderA !== orderB) return orderB - orderA;
      return pndMap[b].total - pndMap[a].total;
    });

    let maxPndKey = "";
    let maxCount = 0;
    sortedPndKeys.forEach(k => {
      if (k !== "Belum/Tidak Terdata" && k !== "Tidak/Belum Sekolah" && pndMap[k].total > maxCount) {
        maxCount = pndMap[k].total;
        maxPndKey = k;
      }
    });

    const totalWarga = filteredWarga.length;

    const pndBarsHtml = sortedPndKeys.map(pndKey => {
      const data = pndMap[pndKey];
      const pct = totalWarga > 0 ? ((data.total / totalWarga) * 100).toFixed(1) : 0;
      return `
        <div class="pnd-item-row" data-pnd="${pndKey}" style="background-color: var(--bg-white); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px; margin-bottom: 8px; cursor: pointer; transition: all 0.2s ease;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
            <span style="font-size: 13px; font-weight: 700; color: var(--text-main); display: flex; align-items: center; gap: 6px;">
              <i data-lucide="graduation-cap" style="width: 14px; height: 14px; color: #8B5CF6;"></i>
              ${pndKey}
              <i data-lucide="chevron-right" style="width: 14px; height: 14px; color: var(--primary);"></i>
            </span>
            <span style="font-size: 12px; font-weight: 700; color: #6D28D9; background-color: #EDE9FE; padding: 2px 8px; border-radius: 100px;">
              ${data.total} Warga (${pct}%) →
            </span>
          </div>
          <div style="width: 100%; height: 6px; background-color: var(--bg-light); border-radius: 100px; overflow: hidden; margin-bottom: 6px;">
            <div style="width: ${pct}%; height: 100%; background-color: #8B5CF6; border-radius: 100px;"></div>
          </div>
          <div style="font-size: 11px; color: var(--text-muted); display: flex; gap: 12px;">
            <span>👨 <strong>${data.laki}</strong> Laki-laki</span>
            <span>👩 <strong>${data.perempuan}</strong> Perempuan</span>
          </div>
        </div>
      `;
    }).join("");

    let titleSuffix = "";
    if (genderFilter === "male") titleSuffix = " (Laki-laki)";
    if (genderFilter === "female") titleSuffix = " (Perempuan)";

    wrapper.innerHTML = `
      <div class="chat-bubble bot">
        Berikut adalah rekapitulasi sebaran <strong>Pendidikan Terakhir${titleSuffix}</strong> warga (Klik level pendidikan untuk melihat daftar nama warga):
      </div>
      <div class="result-card" id="${cardId}" style="border-color: #8B5CF6;">
        
        <!-- LEVEL 1: SUMMARY PER PENDIDIKAN -->
        <div class="pnd-summary-view">
          <div class="result-card-header" style="background-color: #EDE9FE;">
            <div class="result-card-title" style="color: #6D28D9;">
              <i data-lucide="graduation-cap" style="width: 16px; height: 16px; color: #8B5CF6;"></i>
              Statistik Pendidikan Terakhir
            </div>
            <div style="display: flex; gap: 6px; align-items: center;">
              <span class="result-badge" style="background-color: #8B5CF6; color: white;">${sortedPndKeys.length} Kategori</span>
              <button class="btn-export-pnd-excel" style="background-color: #ECFDF5; border: 1px solid #A7F3D0; color: #047857; border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 4px;">📥 Excel</button>
              <button class="btn-export-pnd-pdf" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; color: #1D4ED8; border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 4px;">📄 PDF</button>
            </div>
          </div>

          <div style="padding: 16px 20px;">
            <div style="display: flex; align-items: center; justify-content: space-between; padding-bottom: 12px; margin-bottom: 14px; border-bottom: 1px solid var(--border);">
              <div>
                <div style="font-size: 11px; font-weight: 700; color: var(--text-muted); text-transform: uppercase;">TOTAL DATA WARGA</div>
                <div style="font-size: 18px; font-weight: 700; color: #6D28D9;">${totalWarga} Warga</div>
              </div>
              ${maxPndKey ? `
                <div style="font-size: 11px; background-color: #EDE9FE; border: 1px solid #DDD6FE; padding: 6px 10px; border-radius: 6px; color: #6D28D9; text-align: right; max-width: 140px; line-height: 1.3;">
                  Terbanyak: <strong>${maxPndKey} (${maxCount})</strong>
                </div>
              ` : ''}
            </div>

            <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
              <i data-lucide="mouse-pointer-click" style="width: 14px; height: 14px; color: var(--primary);"></i>
              Klik kategori di bawah untuk melihat detail warganya:
            </div>

            <div style="max-height: 320px; overflow-y: auto; padding-right: 4px;">
              ${pndBarsHtml}
            </div>
          </div>
          <div class="performance-metrics-bar" style="background-color: var(--bg-light); font-size: 11px; padding: 10px 20px; color: var(--text-muted); border-top: 1px solid var(--border);">
            <span>Diklik untuk membuka daftar warga per jenjang pendidikan.</span>
          </div>
        </div>

        <!-- LEVEL 2: CITIZEN LIST BY PENDIDIKAN -->
        <div class="pnd-citizens-view" style="display: none;">
          <div class="result-card-header" style="background-color: var(--bg-light); border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; padding: 12px 20px;">
            <button class="btn-back-summary" style="background: none; border: none; color: var(--text-muted); font-size: 12px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 6px; padding: 4px 8px; border-radius: 6px; transition: background 0.2s;">
              <i data-lucide="arrow-left" style="width: 14px; height: 14px;"></i>
              Kembali
            </button>
            <div class="pnd-list-title" style="font-size: 13px; font-weight: 700; color: var(--text-main);"></div>
          </div>
          <div style="padding: 16px 20px; max-height: 400px; overflow-y: auto;">
            <div class="pnd-list-container" style="display: flex; flex-direction: column; gap: 10px;">
              <!-- Insert list here -->
            </div>
          </div>
        </div>

        <!-- LEVEL 3: BIODATA DETAIL -->
        <div class="pnd-detail-view" style="display: none;">
          <div class="result-card-header" style="background-color: var(--bg-light); border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; padding: 12px 20px;">
            <button class="btn-back-citizens" style="background: none; border: none; color: var(--text-muted); font-size: 12px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 6px; padding: 4px 8px; border-radius: 6px; transition: background 0.2s;">
              <i data-lucide="arrow-left" style="width: 14px; height: 14px;"></i>
              Daftar Warga
            </button>
            <div style="font-size: 13px; font-weight: 700; color: var(--text-main);">Detail Biodata</div>
          </div>
          <div class="pnd-detail-container" style="padding: 20px;">
            <!-- Biodata -->
          </div>
        </div>

      </div>
      <div class="chat-time">${timeStr}</div>
    `;

    chatBody.appendChild(wrapper);
    if (typeof lucide !== "undefined") lucide.createIcons();

    const cardEl = document.getElementById(cardId);
    if (cardEl) {
      const summaryView = cardEl.querySelector('.pnd-summary-view');
      const citizensView = cardEl.querySelector('.pnd-citizens-view');
      const detailView = cardEl.querySelector('.pnd-detail-view');
      
      const btnBackSummary = cardEl.querySelector('.btn-back-summary');
      const listTitle = cardEl.querySelector('.pnd-list-title');
      const listContainer = cardEl.querySelector('.pnd-list-container');
      const detailContainer = cardEl.querySelector('.pnd-detail-container');
      const btnBackCitizensList = cardEl.querySelectorAll('.btn-back-citizens');

      // 1. Click on Row -> Show Citizens
      cardEl.querySelectorAll('.pnd-item-row').forEach(row => {
        row.addEventListener('click', () => {
          const pndKey = row.getAttribute('data-pnd');
          const data = pndMap[pndKey];
          
          listTitle.innerHTML = `<div style="display: flex; align-items: center; gap: 8px;">
            <span>${pndKey} (${data.total} Warga)</span>
            <button class="btn-export-item-excel" style="background-color: #ECFDF5; border: 1px solid #A7F3D0; color: #047857; border-radius: 4px; padding: 2px 6px; font-size: 10px; font-weight: 700; cursor: pointer;">📥 Excel</button>
            <button class="btn-export-item-pdf" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; color: #1D4ED8; border-radius: 4px; padding: 2px 6px; font-size: 10px; font-weight: 700; cursor: pointer;">📄 PDF</button>
          </div>`;

          const btnExcelPnd = listTitle.querySelector('.btn-export-item-excel');
          const btnPdfPnd = listTitle.querySelector('.btn-export-item-pdf');
          if (btnExcelPnd) btnExcelPnd.addEventListener('click', (e) => { e.stopPropagation(); if (typeof XLSX === 'undefined') { alert('XLSX belum dimuat.'); return; } window.cipabotExportExcel(data.list, 'data_pendidikan_' + pndKey + '.xlsx'); });
          if (btnPdfPnd) btnPdfPnd.addEventListener('click', (e) => { e.stopPropagation(); window.cipabotExportPdf('Data Pendidikan: ' + pndKey, data.list); });
          
          listContainer.innerHTML = data.list.map((w, idx) => {
            const isMale = (w.jenis_kelamin || '').toLowerCase().includes("laki") || (w.jenis_kelamin || '').toLowerCase() === "l";
            const jkBadge = isMale 
              ? `<span style="background-color: #DBEAFE; color: #1D4ED8; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700;">Laki-laki</span>`
              : `<span style="background-color: #FCE7F3; color: #BE185D; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700;">Perempuan</span>`;
            
            return `
              <div class="citizen-list-item" data-idx="${idx}" data-pnd="${pndKey}" style="border: 1px solid var(--border); border-radius: 8px; padding: 12px; display: flex; justify-content: space-between; align-items: center; background-color: var(--bg-white); cursor: pointer; transition: border-color 0.2s;">
                <div style="display: flex; gap: 12px; align-items: center;">
                  <div style="width: 24px; height: 24px; border-radius: 12px; background-color: var(--bg-light); color: var(--text-muted); font-size: 11px; font-weight: 700; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                    ${idx + 1}
                  </div>
                  <div>
                    <div style="font-size: 13px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">${w.nama} ${jkBadge}</div>
                    <div style="font-size: 11px; color: var(--text-muted); display: flex; gap: 8px; align-items: center;">
                      <span>NIK: ${w.nik || '-'}</span>
                      <span style="color: #cbd5e1;">|</span>
                      <span>RT ${w.rt || '01'}/RW ${w.rw || '02'}</span>
                    </div>
                  </div>
                </div>
                <i data-lucide="chevron-right" style="width: 16px; height: 16px; color: var(--text-muted);"></i>
              </div>
            `;
          }).join("");

          if (typeof lucide !== "undefined") lucide.createIcons();

          // 2. Click on Citizen -> Show Detail
          cardEl.querySelectorAll('.citizen-list-item').forEach(item => {
            item.addEventListener('click', () => {
              const clickedPnd = item.getAttribute('data-pnd');
              const clickedIdx = parseInt(item.getAttribute('data-idx'));
              const w = pndMap[clickedPnd].list[clickedIdx];
              
              detailContainer.innerHTML = `
                <div style="display: flex; align-items: center; gap: 16px; margin-bottom: 20px; padding-bottom: 16px; border-bottom: 1px dashed var(--border);">
                  <div class="avatar-circle" style="width: 48px; height: 48px; background-color: #EDE9FE; color: #6D28D9; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 20px; font-weight: 700;">
                    ${w.nama ? w.nama.charAt(0).toUpperCase() : '?'}
                  </div>
                  <div>
                    <div style="font-size: 16px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">${w.nama}</div>
                    <div style="display: flex; gap: 8px;">
                      <span class="result-badge" style="background-color: var(--bg-light); color: var(--text-muted);">NIK: ${w.nik || '-'}</span>
                      ${w.status ? `<span class="result-badge" style="background-color: var(--success-light); color: var(--success);">${w.status}</span>` : ''}
                    </div>
                  </div>
                </div>

                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Nomor KK</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.kk || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Jenis Kelamin</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.jenis_kelamin || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tempat Lahir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tempat_lahir || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tanggal Lahir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tanggal_lahir || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Agama</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.agama || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Pendidikan Terakhir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.pendidikan || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Pekerjaan</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.pekerjaan || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Status Pernikahan</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.status_pernikahan || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Hub. Keluarga (SDHK)</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.sdhk || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">RT / RW</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">RT ${w.rt || '01'} / RW ${w.rw || '02'}</div>
                  </div>
                  ${w.alamat ? `<div class="result-item" style="grid-column: 1 / -1;">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Alamat Lengkap</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.alamat}</div>
                  </div>` : ''}
                </div>
              `;

              if (typeof lucide !== "undefined") lucide.createIcons();

              citizensView.style.display = 'none';
              detailView.style.display = 'block';
              scrollToBottom();
            });
          });

          summaryView.style.display = 'none';
          citizensView.style.display = 'block';
          scrollToBottom();
        });
      });

      btnBackSummary.addEventListener('click', () => {
        citizensView.style.display = 'none';
        summaryView.style.display = 'block';
        scrollToBottom();
      });

      btnBackCitizensList.forEach(b => {
        b.addEventListener('click', () => {
          detailView.style.display = 'none';
          citizensView.style.display = 'block';
          scrollToBottom();
        });
      });

      // Export - Pendidikan
      const btnPndExcel = cardEl.querySelector('.btn-export-pnd-excel');
      const btnPndPdf = cardEl.querySelector('.btn-export-pnd-pdf');
      if (btnPndExcel) btnPndExcel.addEventListener('click', (e) => { e.stopPropagation(); if (typeof XLSX === 'undefined') { alert('Library XLSX belum dimuat.'); return; } window.cipabotExportExcel(filteredWarga, 'data_pendidikan_warga.xlsx'); });
      if (btnPndPdf) btnPndPdf.addEventListener('click', (e) => { e.stopPropagation(); window.cipabotExportPdf('Statistik Pendidikan Warga Kelurahan Cipaganti', filteredWarga); });
    }

    scrollToBottom();
  }

  function appendBotAgamaCard(totalDataset, allWarga) {
    const wrapper = document.createElement("div");
    wrapper.className = "chat-bubble-wrapper bot";
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    const cardId = 'agm-card-' + Math.random().toString(36).substring(2, 9);
    
    // Group by Agama
    const agmMap = {};
    allWarga.forEach(w => {
      let agmKey = w.agama ? String(w.agama).trim() : "Belum/Tidak Terdata";
      if (!agmKey || agmKey === "-" || /^\d+$/.test(agmKey) || agmKey.toLowerCase() === "agama") {
        agmKey = "Belum/Tidak Terdata";
      }
      if (!agmMap[agmKey]) {
        agmMap[agmKey] = { total: 0, laki: 0, perempuan: 0, list: [] };
      }
      agmMap[agmKey].total += 1;
      const jk = (w.jenis_kelamin || '').toLowerCase();
      if (jk.includes("laki") || jk === "pria" || jk === "l") {
        agmMap[agmKey].laki += 1;
      } else {
        agmMap[agmKey].perempuan += 1;
      }
      agmMap[agmKey].list.push(w);
    });

    const sortedAgmKeys = Object.keys(agmMap).sort((a, b) => agmMap[b].total - agmMap[a].total);

    let maxAgmKey = sortedAgmKeys.length > 0 ? sortedAgmKeys[0] : "";
    let maxCount = maxAgmKey ? agmMap[maxAgmKey].total : 0;
    const totalWarga = allWarga.length;

    const agmBarsHtml = sortedAgmKeys.map(agmKey => {
      const data = agmMap[agmKey];
      const pct = totalWarga > 0 ? ((data.total / totalWarga) * 100).toFixed(1) : 0;
      return `
        <div class="agm-item-row" data-agm="${agmKey}" style="background-color: var(--bg-white); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px; margin-bottom: 8px; cursor: pointer; transition: all 0.2s ease;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
            <span style="font-size: 13px; font-weight: 700; color: var(--text-main); display: flex; align-items: center; gap: 6px;">
              <i data-lucide="book" style="width: 14px; height: 14px; color: #6366F1;"></i>
              ${agmKey}
              <i data-lucide="chevron-right" style="width: 14px; height: 14px; color: var(--primary);"></i>
            </span>
            <span style="font-size: 12px; font-weight: 700; color: #4338CA; background-color: #E0E7FF; padding: 2px 8px; border-radius: 100px;">
              ${data.total} Warga (${pct}%) →
            </span>
          </div>
          <div style="width: 100%; height: 6px; background-color: var(--bg-light); border-radius: 100px; overflow: hidden; margin-bottom: 6px;">
            <div style="width: ${pct}%; height: 100%; background-color: #6366F1; border-radius: 100px;"></div>
          </div>
          <div style="font-size: 11px; color: var(--text-muted); display: flex; gap: 12px;">
            <span>👨 <strong>${data.laki}</strong> Laki-laki</span>
            <span>👩 <strong>${data.perempuan}</strong> Perempuan</span>
          </div>
        </div>
      `;
    }).join("");

    wrapper.innerHTML = `
      <div class="chat-bubble bot">
        Berikut adalah rekapitulasi sebaran <strong>Status Agama</strong> warga di Kelurahan Cipaganti (Klik agama untuk melihat daftar nama warga):
      </div>
      <div class="result-card" id="${cardId}" style="border-color: #8B5CF6;">
        
        <!-- LEVEL 1: SUMMARY PER PEKERJAAN -->
        <div class="agm-summary-view">
          <div class="result-card-header" style="background-color: #E0E7FF;">
            <div class="result-card-title" style="color: #4338CA;">
              <i data-lucide="book" style="width: 16px; height: 16px; color: #6366F1;"></i>
              Statistik Agama Warga
            </div>
            <div style="display: flex; gap: 6px; align-items: center;">
              <span class="result-badge" style="background-color: #6366F1; color: white;">${sortedAgmKeys.length} Agama</span>
              <button class="btn-export-agm-excel" style="background-color: #ECFDF5; border: 1px solid #A7F3D0; color: #047857; border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 4px;">📥 Excel</button>
              <button class="btn-export-agm-pdf" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; color: #1D4ED8; border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 4px;">📄 PDF</button>
            </div>
          </div>

          <div style="padding: 16px 20px;">
            <div style="display: flex; align-items: center; justify-content: space-between; padding-bottom: 12px; margin-bottom: 14px; border-bottom: 1px solid var(--border);">
              <div>
                <div style="font-size: 11px; font-weight: 700; color: var(--text-muted); text-transform: uppercase;">TOTAL DATA WARGA</div>
                <div style="font-size: 18px; font-weight: 700; color: #4338CA;">${totalWarga} Warga</div>
              </div>
              ${maxAgmKey ? `
                <div style="font-size: 11px; background-color: #E0E7FF; border: 1px solid #C7D2FE; padding: 6px 10px; border-radius: 6px; color: #4338CA; text-align: right; max-width: 140px; line-height: 1.3;">
                  Terbanyak: <strong>${maxAgmKey} (${maxCount})</strong>
                </div>
              ` : ''}
            </div>

            <div style="font-size: 12px; font-weight: 700; color: var(--text-main); margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
              <i data-lucide="mouse-pointer-click" style="width: 14px; height: 14px; color: var(--primary);"></i>
              Klik salah satu agama di bawah untuk melihat warganya:
            </div>

            <div style="max-height: 320px; overflow-y: auto; padding-right: 4px;">
              ${agmBarsHtml}
            </div>
          </div>
          <div class="performance-metrics-bar" style="background-color: var(--bg-light); font-size: 11px; padding: 10px 20px; color: var(--text-muted); border-top: 1px solid var(--border);">
            <span>Diklik untuk membuka daftar warga per agama.</span>
          </div>
        </div>

        <!-- LEVEL 2: CITIZEN LIST BY PEKERJAAN -->
        <div class="agm-citizens-view" style="display: none;">
          <div class="result-card-header" style="background-color: var(--bg-light); border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; padding: 12px 20px;">
            <button class="btn-back-summary" style="background: none; border: none; color: var(--text-muted); font-size: 12px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 6px; padding: 4px 8px; border-radius: 6px; transition: background 0.2s;">
              <i data-lucide="arrow-left" style="width: 14px; height: 14px;"></i>
              Kembali
            </button>
            <div class="agm-list-title" style="font-size: 13px; font-weight: 700; color: var(--text-main);"></div>
          </div>
          <div style="padding: 16px 20px; max-height: 400px; overflow-y: auto;">
            <div class="agm-list-container" style="display: flex; flex-direction: column; gap: 10px;">
              <!-- Insert list here -->
            </div>
          </div>
        </div>

        <!-- LEVEL 3: BIODATA DETAIL -->
        <div class="agm-detail-view" style="display: none;">
          <div class="result-card-header" style="background-color: var(--bg-light); border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; padding: 12px 20px;">
            <button class="btn-back-citizens" style="background: none; border: none; color: var(--text-muted); font-size: 12px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 6px; padding: 4px 8px; border-radius: 6px; transition: background 0.2s;">
              <i data-lucide="arrow-left" style="width: 14px; height: 14px;"></i>
              Daftar Warga
            </button>
            <div style="font-size: 13px; font-weight: 700; color: var(--text-main);">Detail Biodata</div>
          </div>
          <div class="agm-detail-container" style="padding: 20px;">
            <!-- Biodata -->
          </div>
        </div>

      </div>
      <div class="chat-time">${timeStr}</div>
    `;

    chatBody.appendChild(wrapper);
    if (typeof lucide !== "undefined") lucide.createIcons();

    // Event Listeners for Interaction
    const cardEl = document.getElementById(cardId);
    if (cardEl) {
      const summaryView = cardEl.querySelector('.agm-summary-view');
      const citizensView = cardEl.querySelector('.agm-citizens-view');
      const detailView = cardEl.querySelector('.agm-detail-view');
      
      const btnBackSummary = cardEl.querySelector('.btn-back-summary');
      const listTitle = cardEl.querySelector('.agm-list-title');
      const listContainer = cardEl.querySelector('.agm-list-container');
      const detailContainer = cardEl.querySelector('.agm-detail-container');
      const btnBackCitizensList = cardEl.querySelectorAll('.btn-back-citizens');

      // 1. Click on RT/Agama Row -> Show Citizens
      cardEl.querySelectorAll('.agm-item-row').forEach(row => {
        row.addEventListener('click', () => {
          const agmKey = row.getAttribute('data-agm');
          const data = agmMap[agmKey];
          
          listTitle.innerHTML = `<div style="display: flex; align-items: center; gap: 8px;">
            <span>${agmKey} (${data.total} Warga)</span>
            <button class="btn-export-item-excel" style="background-color: #ECFDF5; border: 1px solid #A7F3D0; color: #047857; border-radius: 4px; padding: 2px 6px; font-size: 10px; font-weight: 700; cursor: pointer;">📥 Excel</button>
            <button class="btn-export-item-pdf" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; color: #1D4ED8; border-radius: 4px; padding: 2px 6px; font-size: 10px; font-weight: 700; cursor: pointer;">📄 PDF</button>
          </div>`;

          const btnExcelAgm = listTitle.querySelector('.btn-export-item-excel');
          const btnPdfAgm = listTitle.querySelector('.btn-export-item-pdf');
          if (btnExcelAgm) btnExcelAgm.addEventListener('click', (e) => { e.stopPropagation(); if (typeof XLSX === 'undefined') { alert('XLSX belum dimuat.'); return; } window.cipabotExportExcel(data.list, 'data_agama_' + agmKey + '.xlsx'); });
          if (btnPdfAgm) btnPdfAgm.addEventListener('click', (e) => { e.stopPropagation(); window.cipabotExportPdf('Data Agama: ' + agmKey, data.list); });
          
          listContainer.innerHTML = data.list.map((w, idx) => {
            const isMale = (w.jenis_kelamin || '').toLowerCase().includes("laki") || (w.jenis_kelamin || '').toLowerCase() === "l";
            const jkBadge = isMale 
              ? `<span style="background-color: #DBEAFE; color: #1D4ED8; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700;">Laki-laki</span>`
              : `<span style="background-color: #FCE7F3; color: #BE185D; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700;">Perempuan</span>`;
            
            return `
              <div class="citizen-list-item" data-idx="${idx}" data-agm="${agmKey}" style="border: 1px solid var(--border); border-radius: 8px; padding: 12px; display: flex; justify-content: space-between; align-items: center; background-color: var(--bg-white); cursor: pointer; transition: border-color 0.2s;">
                <div style="display: flex; gap: 12px; align-items: center;">
                  <div style="width: 24px; height: 24px; border-radius: 12px; background-color: var(--bg-light); color: var(--text-muted); font-size: 11px; font-weight: 700; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                    ${idx + 1}
                  </div>
                  <div>
                    <div style="font-size: 13px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">${w.nama} ${jkBadge}</div>
                    <div style="font-size: 11px; color: var(--text-muted); display: flex; gap: 8px; align-items: center;">
                      <span>NIK: ${w.nik || '-'}</span>
                      <span style="color: #cbd5e1;">|</span>
                      <span>RT ${w.rt || '01'}/RW ${w.rw || '02'}</span>
                    </div>
                  </div>
                </div>
                <i data-lucide="chevron-right" style="width: 16px; height: 16px; color: var(--text-muted);"></i>
              </div>
            `;
          }).join("");

          if (typeof lucide !== "undefined") lucide.createIcons();

          // 2. Click on Citizen -> Show Detail
          cardEl.querySelectorAll('.citizen-list-item').forEach(item => {
            item.addEventListener('click', () => {
              const clickedPek = item.getAttribute('data-agm');
              const clickedIdx = parseInt(item.getAttribute('data-idx'));
              const w = agmMap[clickedPek].list[clickedIdx];
              
              detailContainer.innerHTML = `
                <div style="display: flex; align-items: center; gap: 16px; margin-bottom: 20px; padding-bottom: 16px; border-bottom: 1px dashed var(--border);">
                  <div class="avatar-circle" style="width: 48px; height: 48px; background-color: var(--primary-light); color: var(--primary); border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 20px; font-weight: 700;">
                    ${w.nama ? w.nama.charAt(0).toUpperCase() : '?'}
                  </div>
                  <div>
                    <div style="font-size: 16px; font-weight: 700; color: var(--text-main); margin-bottom: 4px;">${w.nama}</div>
                    <div style="display: flex; gap: 8px;">
                      <span class="result-badge" style="background-color: var(--bg-light); color: var(--text-muted);">NIK: ${w.nik || '-'}</span>
                      ${w.status ? `<span class="result-badge" style="background-color: var(--success-light); color: var(--success);">${w.status}</span>` : ''}
                    </div>
                  </div>
                </div>

                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Nomor KK</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.kk || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Jenis Kelamin</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.jenis_kelamin || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tempat Lahir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tempat_lahir || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Tanggal Lahir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.tanggal_lahir || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Agama</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.agama || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Pendidikan Terakhir</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.pendidikan || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Agama</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.agama || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Status Pernikahan</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.status_pernikahan || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Hub. Keluarga (SDHK)</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.sdhk || '-'}</div>
                  </div>
                  <div class="result-item">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">RT / RW</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">RT ${w.rt || '01'} / RW ${w.rw || '02'}</div>
                  </div>
                  ${w.alamat ? `<div class="result-item" style="grid-column: 1 / -1;">
                    <div class="result-label" style="font-size: 11px; color: var(--text-muted);">Alamat Lengkap</div>
                    <div class="result-value" style="font-size: 13px; font-weight: 600; color: var(--text-main);">${w.alamat}</div>
                  </div>` : ''}
                </div>
              `;

              if (typeof lucide !== "undefined") lucide.createIcons();

              citizensView.style.display = 'none';
              detailView.style.display = 'block';
              scrollToBottom();
            });
          });

          summaryView.style.display = 'none';
          citizensView.style.display = 'block';
          scrollToBottom();
        });
      });

      // Back buttons
      btnBackSummary.addEventListener('click', () => {
        citizensView.style.display = 'none';
        summaryView.style.display = 'block';
        scrollToBottom();
      });

      btnBackCitizensList.forEach(b => {
        b.addEventListener('click', () => {
          detailView.style.display = 'none';
          citizensView.style.display = 'block';
          scrollToBottom();
        });
      });

      // Export - Agama
      const btnAgmExcel = cardEl.querySelector('.btn-export-agm-excel');
      const btnAgmPdf = cardEl.querySelector('.btn-export-agm-pdf');
      if (btnAgmExcel) btnAgmExcel.addEventListener('click', (e) => { e.stopPropagation(); if (typeof XLSX === 'undefined') { alert('Library XLSX belum dimuat.'); return; } window.cipabotExportExcel(allWarga, 'data_agama_warga.xlsx'); });
      if (btnAgmPdf) btnAgmPdf.addEventListener('click', (e) => { e.stopPropagation(); window.cipabotExportPdf('Statistik Agama Warga Kelurahan Cipaganti', allWarga); });
    }

    scrollToBottom();
  }



  function handleSearch(query, algo) {
    const trimmedQuery = query.trim();
    if (trimmedQuery === "") return;
    
    // Hide welcome card on first search
    if (welcomeCard) {
      welcomeCard.style.display = "none";
    }

    appendUserMessage(trimmedQuery);
    const typing = appendBotTyping();

    setTimeout(() => {
      typing.remove();
      
      const lowerQuery = trimmedQuery.toLowerCase();
      
      // 1. Check for Greetings
      const greetings = ["halo", "hai", "hello", "hey", "assalamualaikum", "pagi", "siang", "sore", "malam", "permisi"];
      const isGreeting = greetings.some(g => lowerQuery === g || lowerQuery.startsWith(g + " ") || lowerQuery.endsWith(" " + g));
      
      if (isGreeting) {
        appendBotMessage(`Halo! Selamat datang di CIPABOT. Saya asisten digital Kelurahan Cipaganti. Silakan masukkan Nama atau NIK warga untuk melakukan pencarian data.<br><br>Anda juga bisa menanyakan statistik penduduk seperti:
        <br>• <strong>"berapa jumlah laki-laki"</strong> / <strong>"perempuan"</strong>
        <br>• <strong>"berapa jumlah lansia"</strong> / <strong>"jumlah lansia per RT"</strong> (usia ≥ 60 tahun)`);
        return;
      }
      
      // 2. Check for Lansia queries
      const isLansia = lowerQuery.includes("lansia") || lowerQuery.includes("lanjut usia") || lowerQuery.includes("manula") || lowerQuery.includes("60 tahun") || lowerQuery.includes("diatas 60") || lowerQuery.includes("di atas 60") || lowerQuery.includes("tua");

      if (isLansia) {
        const total = DATASET_WARGA.length;
        let allLansia = DATASET_WARGA.map(w => {
          const age = calculateAge(w.tanggal_lahir);
          return { ...w, umur: age };
        }).filter(w => w.umur >= 60);

        allLansia.sort((a, b) => b.umur - a.umur); // Oldest first

        // Check if query asks for per RT breakdown
        const isPerRt = lowerQuery.includes("rt") || lowerQuery.includes("r.t") || lowerQuery.includes("tiap rt") || lowerQuery.includes("per rt") || lowerQuery.includes("sebaran");

        if (isPerRt) {
          appendBotLansiaPerRtCard(total, allLansia);
          saveToHistory(trimmedQuery, algo);
          return;
        }

        const hasMale = lowerQuery.includes("laki") || lowerQuery.includes("pria") || lowerQuery.includes("cowok") || lowerQuery.includes("lk");
        const hasFemale = lowerQuery.includes("perempuan") || lowerQuery.includes("wanita") || lowerQuery.includes("cewek") || lowerQuery.includes("pr");

        let genderFilter = "all";
        let lansiaList = allLansia;
        let messageText = "";

        if (hasMale && !hasFemale) {
          genderFilter = "male";
          lansiaList = allLansia.filter(w => {
            const jk = w.jenis_kelamin.toLowerCase();
            return jk.includes("laki") || jk === "pria" || jk === "l";
          });
          messageText = `Berikut adalah rincian data dan daftar warga lansia <strong>Laki-laki (Pria, Usia ≥ 60 Tahun)</strong> di Kelurahan Cipaganti:`;
        } else if (hasFemale && !hasMale) {
          genderFilter = "female";
          lansiaList = allLansia.filter(w => {
            const jk = w.jenis_kelamin.toLowerCase();
            return jk.includes("perempuan") || jk === "wanita" || jk === "p";
          });
          messageText = `Berikut meperoleh rincian data dan daftar warga lansia <strong>Perempuan (Wanita, Usia ≥ 60 Tahun)</strong> di Kelurahan Cipaganti:`;
        } else {
          const countLaki = allLansia.filter(w => {
            const jk = w.jenis_kelamin.toLowerCase();
            return jk.includes("laki") || jk === "pria" || jk === "l";
          }).length;
          const countPerempuan = allLansia.filter(w => {
            const jk = w.jenis_kelamin.toLowerCase();
            return jk.includes("perempuan") || jk === "wanita" || jk === "p";
          }).length;

          messageText = `Berikut adalah rincian statistik dan daftar warga lansia <strong>(Usia ≥ 60 Tahun)</strong> di Kelurahan Cipaganti (Total <strong>${allLansia.length}</strong> lansia: <strong>${countLaki}</strong> Laki-laki, <strong>${countPerempuan}</strong> Perempuan):`;
        }

        appendBotMessage(messageText);
        appendBotLansiaCard(total, lansiaList, genderFilter, allLansia.length);
        saveToHistory(trimmedQuery, algo);
        return;
      }

      // 3. Check for Statistics / Gender queries
      const isStatsQuery = lowerQuery.includes("jumlah") || lowerQuery.includes("berapa") || lowerQuery.includes("total") || lowerQuery.includes("statistik");
      const hasGender = lowerQuery.includes("laki") || lowerQuery.includes("pria") || lowerQuery.includes("perempuan") || lowerQuery.includes("wanita") || lowerQuery.includes("gender");
      
      if (isStatsQuery || hasGender) {
        const total = DATASET_WARGA.length;
        const laki = DATASET_WARGA.filter(w => {
          const jk = w.jenis_kelamin.toLowerCase();
          return jk.includes("laki") || jk === "pria" || jk === "l";
        }).length;
        const perempuan = DATASET_WARGA.filter(w => {
          const jk = w.jenis_kelamin.toLowerCase();
          return jk.includes("perempuan") || jk === "wanita" || jk === "p";
        }).length;
        
        const pctLaki = total > 0 ? ((laki / total) * 100).toFixed(1) : 0;
        const pctPerempuan = total > 0 ? ((perempuan / total) * 100).toFixed(1) : 0;
        
        appendBotMessage(`Berikut adalah statistik gender dari dataset aktif kelurahan (total <strong>${total}</strong> warga):`);
        appendBotStatsCard(total, laki, pctLaki, perempuan, pctPerempuan);
        saveToHistory(trimmedQuery, algo);
        return;
      }
      // 4. Check for Pekerjaan queries
      const isPekerjaan = lowerQuery.includes("pekerjaan") || lowerQuery.includes("profesi") || lowerQuery.includes("status pekerjaan") || lowerQuery.includes("pegawai") || lowerQuery.includes("kerja");
      if (isPekerjaan) {
        const hasMale = lowerQuery.includes("laki") || lowerQuery.includes("pria");
        const hasFemale = lowerQuery.includes("perempuan") || lowerQuery.includes("wanita");
        let genderFilter = "all";
        if (hasMale && !hasFemale) genderFilter = "male";
        if (hasFemale && !hasMale) genderFilter = "female";
        
        appendBotPekerjaanCard(DATASET_WARGA.length, DATASET_WARGA, genderFilter);
        saveToHistory(trimmedQuery, algo);
        return;
      }

      
      // Check for Pendidikan queries
      const isPendidikan = lowerQuery.includes("pendidikan") || lowerQuery.includes("sekolah") || lowerQuery.includes("kuliah") || lowerQuery.includes("sarjana") || lowerQuery.includes("lulusan");
      if (isPendidikan) {
        const hasMale = lowerQuery.includes("laki") || lowerQuery.includes("pria");
        const hasFemale = lowerQuery.includes("perempuan") || lowerQuery.includes("wanita");
        let genderFilter = "all";
        if (hasMale && !hasFemale) genderFilter = "male";
        if (hasFemale && !hasMale) genderFilter = "female";
        
        appendBotPendidikanCard(DATASET_WARGA.length, DATASET_WARGA, genderFilter);
        saveToHistory(trimmedQuery, algo);
        return;
      }

      // 5. Check for Agama queries
      const isAgama = lowerQuery.includes("agama") || lowerQuery.includes("kepercayaan") || lowerQuery.includes("keyakinan") || lowerQuery.includes("religi");
      if (isAgama) {
        appendBotAgamaCard(DATASET_WARGA.length, DATASET_WARGA);
        saveToHistory(trimmedQuery, algo);
        return;
      }


      
      function extractSearchKeyword(qStr) {
        let q = qStr.trim();
        const prefixes = [
          /^carikan\s+nama\s+/i,
          /^carikan\s+warga\s+nama\s+/i,
          /^carikan\s+warga\s+/i,
          /^carikan\s+/i,
          /^cari\s+nama\s+/i,
          /^cari\s+warga\s+nama\s+/i,
          /^cari\s+warga\s+/i,
          /^cari\s+/i,
          /^tolong\s+carikan\s+nama\s+/i,
          /^tolong\s+carikan\s+/i,
          /^tolong\s+cari\s+/i,
          /^siapa\s+nama\s+/i,
          /^siapa\s+warga\s+nama\s+/i,
          /^siapa\s+warga\s+/i,
          /^siapa\s+/i,
          /^tampilkan\s+nama\s+/i,
          /^tampilkan\s+warga\s+/i,
          /^tampilkan\s+/i,
          /^data\s+warga\s+nama\s+/i,
          /^data\s+warga\s+/i,
          /^informasi\s+warga\s+/i
        ];

        for (const prefix of prefixes) {
          if (prefix.test(q)) {
            const cleaned = q.replace(prefix, '').trim();
            if (cleaned.length > 0) {
              return cleaned;
            }
          }
        }
        return q;
      }

      // 4. Regular record search using KMP & BM (with natural language prefix removal)
      const cleanKeyword = extractSearchKeyword(trimmedQuery);
      const result = searchWarga(cleanKeyword, algo);
      appendBotResult(result, algo);
      if (result.warga) {
        saveToHistory(trimmedQuery, algo);
      }
    }, 800);
  }

  function scrollToBottom() {
    chatBody.scrollTop = chatBody.scrollHeight;
  }

  // Excel File Parsing with SheetJS
  if (excelFileInput) {
    excelFileInput.addEventListener("change", (e) => {
      const file = e.target.files[0];
      if (!file) return;

      excelUploadLabel.textContent = "Membaca...";

      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          
          // Helper to handle multi-row / merged Excel headers (e.g., Row 1: Header, Row 2: Subheader, Row 3: Column numbers 1..12)
          function parseSheetSmart(worksheet) {
            const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, cellDates: true, defval: "" });
            if (!rawRows || rawRows.length === 0) return [];

            // 1. Identify row index where actual citizen data starts
            let dataStartIdx = -1;
            for (let r = 0; r < Math.min(rawRows.length, 12); r++) {
              const row = rawRows[r];
              if (!row || row.length === 0) continue;
              
              const hasNikOrKk = row.some(cell => {
                const s = String(cell || '').trim();
                return /^\d{9,}$/.test(s);
              });

              const col0 = String(row[0] || '').trim();
              const isRowNumber = /^\d+\.?$/.test(col0);
              const hasDataCells = row.some((cell, cIdx) => cIdx > 0 && String(cell || '').trim().length > 2);

              if (hasNikOrKk || (isRowNumber && hasDataCells && r >= 2)) {
                dataStartIdx = r;
                break;
              }
            }

            if (dataStartIdx === -1) {
              for (let r = 0; r < Math.min(rawRows.length, 12); r++) {
                const firstCells = rawRows[r].slice(0, 4).map(c => String(c || '').toLowerCase().trim());
                const isHeaderRow = firstCells.some(c => c === 'no' || c === 'kk' || c === 'nik' || c === 'nama' || c === 'l / p' || c === 'l/p');
                const isNumberIdxRow = firstCells.every(c => c === '' || /^\d{1,2}$/.test(c));
                if (!isHeaderRow && !isNumberIdxRow && rawRows[r].some(c => String(c || '').trim() !== '')) {
                  dataStartIdx = r;
                  break;
                }
              }
            }
            if (dataStartIdx === -1) dataStartIdx = 1;

            // 2. Build column headers by combining all header rows before dataStartIdx
            const maxCols = Math.max(...rawRows.slice(0, dataStartIdx + 1).map(r => r.length));
            const headers = [];

            for (let c = 0; c < maxCols; c++) {
              let parts = [];
              for (let r = 0; r < dataStartIdx; r++) {
                const val = String(rawRows[r][c] || '').trim();
                if (/^\d{1,2}$/.test(val)) continue; // skip 1..12 index numbers row
                if (val && !parts.includes(val)) {
                  parts.push(val);
                }
              }
              headers.push(parts.join(' ').trim() || `KOLOM_${c+1}`);
            }

            // 3. Convert data rows into objects
            const resultObjects = [];
            for (let r = dataStartIdx; r < rawRows.length; r++) {
              const row = rawRows[r];
              if (!row || row.every(c => String(c || '').trim() === '')) continue;

              const lineText = row.map(c => String(c || '')).join(' ').toLowerCase();
              if (lineText.includes('jumlah') || lineText.includes('total')) continue;

              const rowObj = {};
              headers.forEach((h, c) => {
                rowObj[h] = row[c] !== undefined ? row[c] : "";
              });
              resultObjects.push(rowObj);
            }

            return resultObjects;
          }

          let allRows = [];
          workbook.SheetNames.forEach((sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            const rows = parseSheetSmart(worksheet);
            if (rows && rows.length > 0) {
              rows.forEach(r => r._sheetName = sheetName);
              allRows.push(...rows);
            }
          });

          if (allRows.length > 0) {
            // Detect all column headers
            const detectedColumns = Object.keys(allRows[0]).filter(k => k !== '_sheetName');
            console.log("[CIPABOT] Kolom Excel yang terdeteksi:", detectedColumns);
            
            // Show detected columns in chat for transparency
            const colListHtml = detectedColumns.map(c => `<code style="background:#EFF6FF;padding:2px 6px;border-radius:4px;font-size:11px;color:#2563EB;">${c}</code>`).join(', ');
            appendBotMessage(`📋 Kolom yang terdeteksi dari Excel: ${colListHtml}`);
            
            const mappedDataset = allRows.map((row, idx) => {
              const keys = Object.keys(row);
              
              const cleanStr = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');

              // Robust field finder with safe matching:
              const getField = (aliases) => {
                // Pass 1: Exact match on raw string
                for (const alias of aliases) {
                  const la = alias.toLowerCase().trim();
                  for (const key of keys) {
                    if (key === '_sheetName') continue;
                    const lk = key.toLowerCase().trim();
                    if (lk === la) {
                      const v = row[key];
                      return (v !== undefined && v !== null) ? String(v).trim() : "";
                    }
                  }
                }
                // Pass 2: Cleaned exact match (strip spaces, slashes, punctuation)
                for (const alias of aliases) {
                  const ca = cleanStr(alias);
                  if (!ca) continue;
                  for (const key of keys) {
                    if (key === '_sheetName') continue;
                    const ck = cleanStr(key);
                    if (ck === ca) {
                      const v = row[key];
                      return (v !== undefined && v !== null) ? String(v).trim() : "";
                    }
                  }
                }
                // Pass 3: Cleaned includes match (only for aliases >= 3 chars)
                for (const alias of aliases) {
                  const ca = cleanStr(alias);
                  if (ca.length < 3) continue;
                  for (const key of keys) {
                    if (key === '_sheetName') continue;
                    const ck = cleanStr(key);
                    if (ck.includes(ca)) {
                      const v = row[key];
                      return (v !== undefined && v !== null) ? String(v).trim() : "";
                    }
                  }
                }
                // Pass 4: Word-level match
                for (const alias of aliases) {
                  const aliasWords = alias.toLowerCase().split(/\s+/);
                  if (aliasWords.length < 2) continue;
                  for (const key of keys) {
                    if (key === '_sheetName') continue;
                    const lk = key.toLowerCase().trim();
                    if (aliasWords.every(w => lk.includes(w))) {
                      const v = row[key];
                      return (v !== undefined && v !== null) ? String(v).trim() : "";
                    }
                  }
                }
                return "";
              };

              // --- Handle Excel dates ---
              const formatDate = (val) => {
                if (!val) return "";
                if (val instanceof Date) {
                  const dd = String(val.getDate()).padStart(2, '0');
                  const mm = String(val.getMonth() + 1).padStart(2, '0');
                  const yyyy = val.getFullYear();
                  return `${dd}-${mm}-${yyyy}`;
                }
                
                const strVal = String(val).trim();
                if (!strVal) return "";

                const num = Number(strVal);
                if (!isNaN(num) && num > 1000 && num < 100000) {
                  const dateObj = new Date(Math.round((num - 25569) * 86400 * 1000));
                  const dd = String(dateObj.getUTCDate()).padStart(2, '0');
                  const mm = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
                  const yyyy = dateObj.getUTCFullYear();
                  return `${dd}-${mm}-${yyyy}`;
                }

                const isoMatch = strVal.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
                if (isoMatch) {
                  const yyyy = isoMatch[1];
                  const mm = String(isoMatch[2]).padStart(2, '0');
                  const dd = String(isoMatch[3]).padStart(2, '0');
                  return `${dd}-${mm}-${yyyy}`;
                }

                const dmyMatch = strVal.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
                if (dmyMatch) {
                  const dd = String(dmyMatch[1]).padStart(2, '0');
                  const mm = String(dmyMatch[2]).padStart(2, '0');
                  const yyyy = dmyMatch[3];
                  return `${dd}-${mm}-${yyyy}`;
                }

                return strVal;
              };

              // --- Map each field ---
              const nama = getField(["nama lengkap", "nama_lengkap", "nama warga", "nama", "name"]);
              const nik = getField(["nik", "no nik", "no. nik", "nomor induk kependudukan", "nomor induk", "no ktp", "no. ktp"]);
              const kk = getField(["no kk", "no. kk", "nomor kk", "nomor kartu keluarga", "kk", "kartu keluarga"]);
              
              const rawJk = getField(["jenis kelamin", "jenis_kelamin", "jns kelamin", "jns. kelamin", "l/p", "lk/pr", "kelamin", "gender", "jk", "j/k", "sex", "lp"]);
              let jenis_kelamin = rawJk;
              if (rawJk) {
                const jkUpper = rawJk.toUpperCase().trim();
                if (jkUpper === "L" || jkUpper === "LK" || jkUpper === "LAKI-LAKI" || jkUpper === "LAKILAKI" || jkUpper === "PRIA") {
                  jenis_kelamin = "Laki-laki";
                } else if (jkUpper === "P" || jkUpper === "PR" || jkUpper === "PEREMPUAN" || jkUpper === "WANITA") {
                  jenis_kelamin = "Perempuan";
                }
              }
              
              let tempat_lahir = "";
              let tanggal_lahir = "";
              
              const tempatSeparate = getField(["tempat lahir", "tempat_lahir", "tmp lahir", "tmpt lahir", "kota lahir", "tempat", "tmp"]);
              const tanggalSeparate = getField(["tgl lahir", "tgl. lahir", "tgl_lahir", "tgl.lahir", "tgllahir", "tgl lhr", "tgl. lhr", "tanggal lahir", "tanggal_lahir", "tanggal.lahir", "tgl", "tanggal", "lahir"]);
              
              if (tempatSeparate || tanggalSeparate) {
                tempat_lahir = tempatSeparate;
                tanggal_lahir = formatDate(tanggalSeparate);
              } else {
                const ttlCombined = getField(["tempat tanggal lahir", "tempat, tanggal lahir", "tempat/tanggal lahir", "ttl"]);
                if (ttlCombined) {
                  const commaIdx = ttlCombined.indexOf(",");
                  if (commaIdx > 0) {
                    tempat_lahir = ttlCombined.substring(0, commaIdx).trim();
                    tanggal_lahir = formatDate(ttlCombined.substring(commaIdx + 1).trim());
                  } else {
                    const maybeDate = Number(ttlCombined);
                    if (!isNaN(maybeDate) && maybeDate > 1000) {
                      tanggal_lahir = formatDate(ttlCombined);
                    } else {
                      tempat_lahir = ttlCombined;
                    }
                  }
                }
              }

              const agama = getField(["agama", "religion", "agm"]);
              const pendidikan = getField(["pendidikan", "pendidikan terakhir", "pend. terakhir", "pend", "sekolah"]);
              const pekerjaan = getField(["pekerjaan", "kerja", "profesi", "job", "pkj"]);
              const alamat = getField(["alamat", "alamat lengkap", "alamat tinggal", "address", "almt"]);
              
              let rtVal = getField(["rt", "r.t", "no rt", "no. rt"]);
              let rwVal = getField(["rw", "r.w", "no rw", "no. rw"]);

              // Extract RT from Sheet Name if empty (e.g. "RT 03" -> "03")
              if (!rtVal && row._sheetName) {
                const sName = String(row._sheetName).trim();
                const m = sName.match(/\d+/);
                if (m) {
                  rtVal = String(parseInt(m[0])).padStart(2, '0');
                } else {
                  rtVal = sName;
                }
              }
              if (!rtVal) rtVal = "01";
              if (!rwVal) rwVal = "02"; // Sample RW 02

              const statusPernikahan = getField(["status perkawinan", "status pernikahan", "status_pernikahan", "status kawin", "perkawinan", "kawin"]) || getField(["status"]);
              const sdhk = getField(["sdhk", "status dalam hubungan keluarga", "hubungan keluarga", "hub keluarga", "shdk", "status hubungan"]);
              const keterangan = getField(["keterangan", "ket", "ket."]);

              return {
                nama: nama || `WARGA ${idx+1}`,
                nik: nik || "",
                kk: kk || "",
                jenis_kelamin: jenis_kelamin || "",
                tempat_lahir: tempat_lahir,
                tanggal_lahir: tanggal_lahir,
                agama: agama || "",
                pendidikan: pendidikan || "",
                pekerjaan: pekerjaan || "",
                status_pernikahan: statusPernikahan || "",
                rt: rtVal,
                rw: rwVal,
                alamat: alamat || `RT ${rtVal} / RW ${rwVal}, Kelurahan Cipaganti`,
                sdhk: sdhk || "",
                status: keterangan || "Aktif"
              };
            });

            // Debug: log first mapped record
            console.log("[CIPABOT] Contoh data pertama setelah mapping:", mappedDataset[0]);

            // Update global dataset
            DATASET_WARGA.length = 0;
            DATASET_WARGA.push(...mappedDataset);

            
            // Add new file to Multi-File History in localStorage
            try {
              const history = getFileHistory();
              const now = new Date();
              const timeStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth()+1).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
              const fileId = "file_" + Date.now();
              const newFileObj = {
                id: fileId,
                name: file.name,
                time: timeStr,
                count: mappedDataset.length,
                data: mappedDataset
              };

              // Filter out duplicate filename if re-uploaded
              const filteredHistory = history.filter(f => f.name !== file.name);
              filteredHistory.unshift(newFileObj);
              if (filteredHistory.length > 8) filteredHistory.pop(); // Max 8 files history

              saveFileHistory(filteredHistory);
              switchActiveDataset(fileId, false);
            } catch (errSave) {
              console.warn("[CIPABOT] Gagal menyimpan file ke riwayat:", errSave);
            }

            excelUploadLabel.textContent = `📁 ${file.name}`;

            
            // Hide welcome card on upload success to show notification
            if (welcomeCard) {
              welcomeCard.style.display = "none";
            }
            appendExcelLoadSuccess(file.name, mappedDataset.length, workbook.SheetNames.length);
          } else {
            alert("File Excel kosong atau format kolom tidak dikenali.");
            excelUploadLabel.textContent = "Gagal Mengunggah";
          }
        } catch (err) {
          console.error(err);
          alert("Gagal mengurai file Excel. Pastikan file berformat .xlsx atau .xls.");
          excelUploadLabel.textContent = "Gagal Mengunggah";
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  // Event Listeners
  
  // Reset Excel Dataset button handler
  const btnResetExcel = document.getElementById("btn-reset-excel");
  if (btnResetExcel) {
    btnResetExcel.addEventListener("click", () => {
      const activeId = localStorage.getItem("cipabot_active_file_id") || "default";
      if (activeId === "default") {
        alert("Saat ini Anda sedang menggunakan Dataset Bawaan.");
        return;
      }
      const history = getFileHistory();
      const activeFile = history.find(f => f.id === activeId);
      const nameStr = activeFile ? activeFile.name : "file aktif";
      if (confirm(`Apakah Anda yakin ingin menghapus file "${nameStr}" dari riwayat terunggah?`)) {
        const updatedHistory = history.filter(f => f.id !== activeId);
        saveFileHistory(updatedHistory);
        switchActiveDataset("default", true);
        alert(`File "${nameStr}" berhasil dihapus dari riwayat.`);
      }
    });
  }

  sendBtn.addEventListener("click", () => {
    const query = chatInput.value;
    const algo = algoSelect.value;
    handleSearch(query, algo);
    chatInput.value = "";
  });

  chatInput.addEventListener("keypress", (e) => {
    if (e.key === "Enter") {
      const query = chatInput.value;
      const algo = algoSelect.value;
      handleSearch(query, algo);
      chatInput.value = "";
    }
  });

  newSearchBtn.addEventListener("click", () => {
    chatBody.innerHTML = "";
    if (welcomeCard) {
      welcomeCard.style.display = "flex";
      chatBody.appendChild(welcomeCard);
    }
    chatInput.value = "";
  });

  // Suggestion tags in welcome card
  document.querySelectorAll(".welcome-suggestions .tag-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const txt = btn.textContent.replace("Contoh: ", "");
      if (txt.startsWith("NIK ")) {
        chatInput.value = txt.replace("NIK ", "");
      } else {
        chatInput.value = txt;
      }
      chatInput.focus();
    });
  });

  updateHistoryUI();

  
  // Copy default dataset once loaded
  if (DEFAULT_WARGA.length === 0) {
    DEFAULT_WARGA = JSON.parse(JSON.stringify(DATASET_WARGA));
  }

  const excelFileSelector = document.getElementById("excel-file-selector");

  function getFileHistory() {
    try {
      return JSON.parse(localStorage.getItem("cipabot_file_history")) || [];
    } catch (e) {
      return [];
    }
  }

  function saveFileHistory(history) {
    try {
      localStorage.setItem("cipabot_file_history", JSON.stringify(history));
    } catch (e) {
      console.warn("[CIPABOT] Gagal menyimpan riwayat file ke localStorage:", e);
    }
  }

  function updateFileSelectorUI() {
    if (!excelFileSelector) return;
    const history = getFileHistory();
    const activeId = localStorage.getItem("cipabot_active_file_id") || "default";

    excelFileSelector.innerHTML = '<option value="default">🌐 Dataset Bawaan (RW 02)</option>';

    history.forEach((fileObj) => {
      const opt = document.createElement("option");
      opt.value = fileObj.id;
      opt.textContent = `📁 ${fileObj.name} (${fileObj.count} Warga - ${fileObj.time})`;
      excelFileSelector.appendChild(opt);
    });

    excelFileSelector.value = activeId;
  }

  function switchActiveDataset(fileId, isUserAction = false) {
    const history = getFileHistory();
    if (fileId === "default" || !fileId) {
      DATASET_WARGA.length = 0;
      DATASET_WARGA.push(...DEFAULT_WARGA);
      localStorage.setItem("cipabot_active_file_id", "default");
      if (excelUploadLabel) excelUploadLabel.textContent = "+ Unggah Excel Baru";
      updateFileSelectorUI();
      if (isUserAction) {
        appendBotMessage(`Berhasil beralih ke <strong>Dataset Bawaan Kelurahan Cipaganti</strong> (total <strong>${DATASET_WARGA.length}</strong> warga).`);
      }
    } else {
      const fileObj = history.find(f => f.id === fileId);
      if (fileObj && Array.isArray(fileObj.data)) {
        DATASET_WARGA.length = 0;
        DATASET_WARGA.push(...fileObj.data);
        localStorage.setItem("cipabot_active_file_id", fileObj.id);
        if (excelUploadLabel) excelUploadLabel.textContent = `📁 ${fileObj.name}`;
        updateFileSelectorUI();
        if (isUserAction) {
          appendBotMessage(`Berhasil beralih ke file Excel <strong>"${fileObj.name}"</strong> (total <strong>${fileObj.count}</strong> warga).`);
        }
      } else {
        // Fallback if file not found
        switchActiveDataset("default");
      }
    }
  }

  if (excelFileSelector) {
    excelFileSelector.addEventListener("change", (e) => {
      switchActiveDataset(e.target.value, true);
    });
  }

  // Restore active dataset on initialization
  const initialActiveId = localStorage.getItem("cipabot_active_file_id") || "default";
  switchActiveDataset(initialActiveId, false);

});

// ----------------------------------------------------
// Hero Mini Chatbot Simulator Logic
// ----------------------------------------------------
function initHeroSimulator() {
  const heroChatBody = document.getElementById("hero-chat-body");
  const heroChatInput = document.getElementById("hero-chat-input");
  const heroSendBtn = document.getElementById("hero-send-btn");

  function appendHeroMsg(text, isUser = false) {
    const bubble = document.createElement("div");
    bubble.className = `message-bubble ${isUser ? 'message-user' : 'message-bot'}`;
    bubble.textContent = text;
    heroChatBody.appendChild(bubble);
    
    // Auto-scroll
    heroChatBody.scrollTop = heroChatBody.scrollHeight;
  }

  function runHeroSearch(query) {
    if (!query || query.trim() === "") return;
    appendHeroMsg(query, true);

    const typing = document.createElement("div");
    typing.className = "message-bubble message-bot";
    typing.textContent = "...";
    heroChatBody.appendChild(typing);
    heroChatBody.scrollTop = heroChatBody.scrollHeight;

    setTimeout(() => {
      typing.remove();
      const res = searchWarga(query.trim(), "kmp"); // Default to KMP for quick display
      if (res.warga) {
        appendHeroMsg(`Ditemukan! NIK: ${res.warga.nik}, Nama: ${res.warga.nama}, RT/RW: ${res.warga.rt}/${res.warga.rw}.`);
        appendHeroMsg(`Algoritma KMP membandingkan ${res.comparisons} karakter dalam ${res.duration} ms.`);
      } else {
        appendHeroMsg(`Data "${query}" tidak ditemukan di file Excel Kelurahan Cipaganti.`);
      }
    }, 600);
  }

  // Pill clicks
  document.querySelectorAll("#hero-chat-body .tag-btn").forEach(btn => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const txt = btn.textContent;
      if (txt.includes("Andri")) {
        runHeroSearch("Andri Rustandi");
      } else if (txt.includes("Siti")) {
        runHeroSearch("Siti Aminah");
      } else {
        runHeroSearch("3273011210010002");
      }
    });
  });

  heroSendBtn.addEventListener("click", () => {
    const val = heroChatInput.value;
    runHeroSearch(val);
    heroChatInput.value = "";
  });

  heroChatInput.addEventListener("keypress", (e) => {
    if (e.key === "Enter") {
      const val = heroChatInput.value;
      runHeroSearch(val);
      heroChatInput.value = "";
    }
  });
}

// ----------------------------------------------------
// How It Works: String Matching Visualizer
// ----------------------------------------------------
function initTimelineVisualizer() {
  const textCharsContainer = document.getElementById("text-chars");
  const patternCharsContainer = document.getElementById("pattern-chars");
  const compCountEl = document.getElementById("visualizer-comp");
  const shiftCountEl = document.getElementById("visualizer-shift");
  const stepBtn = document.getElementById("btn-visualizer-next");
  const resetBtn = document.getElementById("btn-visualizer-reset");
  const kmpPill = document.getElementById("pill-kmp");
  const bmPill = document.getElementById("pill-bm");

  const text = "CIPAGANTI";
  const pattern = "PAGAN";
  
  let currentAlgo = "kmp";
  let searchStates = [];
  let stateIndex = 0;

  // Pre-generate states for visualization
  function generateKmpStates() {
    searchStates = [];
    const n = text.length;
    const m = pattern.length;
    
    // Simulating KMP checks step-by-step
    // Let's manually feed standard steps for clarity of visualization
    
    // Step 0: Initial alignment
    // C I P A G A N T I
    // P A G A N
    // Compare C and P (Mismatch) -> Shift pattern to index 1
    searchStates.push({
      shift: 0,
      checkingIdx: 0,
      matches: [],
      mismatch: 0,
      comparisons: 1,
      shiftCount: 0
    });

    // Step 1: compare text[1] 'I' and pattern[0] 'P' (Mismatch) -> Shift pattern to index 2
    searchStates.push({
      shift: 1,
      checkingIdx: 0,
      matches: [],
      mismatch: 0,
      comparisons: 2,
      shiftCount: 1
    });

    // Step 2: compare text[2] 'P' and pattern[0] 'P' (Match!)
    searchStates.push({
      shift: 2,
      checkingIdx: 0,
      matches: [0],
      mismatch: -1,
      comparisons: 3,
      shiftCount: 2
    });

    // Step 3: compare text[3] 'A' and pattern[1] 'A' (Match!)
    searchStates.push({
      shift: 2,
      checkingIdx: 1,
      matches: [0, 1],
      mismatch: -1,
      comparisons: 4,
      shiftCount: 2
    });

    // Step 4: compare text[4] 'G' and pattern[2] 'G' (Match!)
    searchStates.push({
      shift: 2,
      checkingIdx: 2,
      matches: [0, 1, 2],
      mismatch: -1,
      comparisons: 5,
      shiftCount: 2
    });

    // Step 5: compare text[5] 'A' and pattern[3] 'A' (Match!)
    searchStates.push({
      shift: 2,
      checkingIdx: 3,
      matches: [0, 1, 2, 3],
      mismatch: -1,
      comparisons: 6,
      shiftCount: 2
    });

    // Step 6: compare text[6] 'N' and pattern[4] 'N' (Match! Full pattern matches!)
    searchStates.push({
      shift: 2,
      checkingIdx: 4,
      matches: [0, 1, 2, 3, 4],
      mismatch: -1,
      comparisons: 7,
      shiftCount: 2,
      completed: true
    });
  }

  function generateBmStates() {
    searchStates = [];
    
    // Boyer-Moore searches from right to left
    // Text:    C I P A G A N T I
    // Pattern: P A G A N
    // Step 0: pattern at shift 0. Compare text[4]='G' with pattern[4]='N' (Mismatch).
    // Bad character rule: 'G' is in pattern at index 2. Shift pattern so 'G' aligns with text[4].
    // New shift = 4 - 2 = 2.
    searchStates.push({
      shift: 0,
      checkingIdx: 4,
      matches: [],
      mismatch: 4,
      comparisons: 1,
      shiftCount: 0
    });

    // Step 1: pattern at shift 2. Compare text[6]='N' with pattern[4]='N' (Match!)
    searchStates.push({
      shift: 2,
      checkingIdx: 4,
      matches: [4],
      mismatch: -1,
      comparisons: 2,
      shiftCount: 1
    });

    // Step 2: Compare text[5]='A' with pattern[3]='A' (Match!)
    searchStates.push({
      shift: 2,
      checkingIdx: 3,
      matches: [3, 4],
      mismatch: -1,
      comparisons: 3,
      shiftCount: 1
    });

    // Step 3: Compare text[4]='G' with pattern[2]='G' (Match!)
    searchStates.push({
      shift: 2,
      checkingIdx: 2,
      matches: [2, 3, 4],
      mismatch: -1,
      comparisons: 4,
      shiftCount: 1
    });

    // Step 4: Compare text[3]='A' with pattern[1]='A' (Match!)
    searchStates.push({
      shift: 2,
      checkingIdx: 1,
      matches: [1, 2, 3, 4],
      mismatch: -1,
      comparisons: 5,
      shiftCount: 1
    });

    // Step 5: Compare text[2]='P' with pattern[0]='P' (Match! Full match!)
    searchStates.push({
      shift: 2,
      checkingIdx: 0,
      matches: [0, 1, 2, 3, 4],
      mismatch: -1,
      comparisons: 6,
      shiftCount: 1,
      completed: true
    });
  }

  function renderState() {
    if (searchStates.length === 0) return;
    const state = searchStates[stateIndex];

    // Render Text Row
    textCharsContainer.innerHTML = "";
    for (let idx = 0; idx < text.length; idx++) {
      const char = text[idx];
      const box = document.createElement("div");
      box.className = "char-box";
      box.textContent = char;

      // Highlight logic
      const patAlignedIdx = idx - state.shift;
      if (patAlignedIdx >= 0 && patAlignedIdx < pattern.length) {
        if (state.matches.includes(patAlignedIdx)) {
          box.classList.add("match");
        } else if (state.mismatch === patAlignedIdx) {
          box.classList.add("mismatch");
        } else if (state.checkingIdx === patAlignedIdx) {
          box.classList.add("pointer");
        }
      }
      textCharsContainer.appendChild(box);
    }

    // Render Pattern Row (with spaces/offsets for shift alignment)
    patternCharsContainer.innerHTML = "";
    // Insert spacing offsets
    for (let offset = 0; offset < state.shift; offset++) {
      const spacer = document.createElement("div");
      spacer.className = "char-box";
      spacer.style.border = "none";
      spacer.style.backgroundColor = "transparent";
      patternCharsContainer.appendChild(spacer);
    }
    // Pattern chars
    for (let idx = 0; idx < pattern.length; idx++) {
      const char = pattern[idx];
      const box = document.createElement("div");
      box.className = "char-box";
      box.textContent = char;

      if (state.matches.includes(idx)) {
        box.classList.add("match");
      } else if (state.mismatch === idx) {
        box.classList.add("mismatch");
      } else if (state.checkingIdx === idx) {
        box.classList.add("pointer");
      }
      patternCharsContainer.appendChild(box);
    }

    // Update stats
    compCountEl.textContent = state.comparisons;
    shiftCountEl.textContent = state.shiftCount;

    // Button states
    if (state.completed || stateIndex === searchStates.length - 1) {
      stepBtn.disabled = true;
      stepBtn.textContent = "Selesai";
      stepBtn.style.opacity = "0.5";
    } else {
      stepBtn.disabled = false;
      stepBtn.textContent = "Langkah Berikutnya";
      stepBtn.style.opacity = "1";
    }
  }

  function selectAlgo(algo) {
    currentAlgo = algo;
    stateIndex = 0;
    if (algo === "kmp") {
      kmpPill.classList.add("active");
      bmPill.classList.remove("active");
      generateKmpStates();
    } else {
      kmpPill.classList.remove("active");
      bmPill.classList.add("active");
      generateBmStates();
    }
    renderState();
  }

  stepBtn.addEventListener("click", () => {
    if (stateIndex < searchStates.length - 1) {
      stateIndex++;
      renderState();
    }
  });

  resetBtn.addEventListener("click", () => {
    stateIndex = 0;
    renderState();
  });

  kmpPill.addEventListener("click", () => selectAlgo("kmp"));
  bmPill.addEventListener("click", () => selectAlgo("bm"));

  // Initial load
  selectAlgo("kmp");
}
