const fs = require('fs');
const path = require('path');

const firstNames = ["Andri","Siti","Gilang","Dedi","Ratna","Budi","Cecep","Eka","Iis","Yayan","Andreas","Rizky","Agus","Rina","Wawan","Ika","Tina","Rudi","Fajar","Maya","Rina","Hendra","Rini","Suryo","Rama","Indra","Nina","Putri","Bayu","Dita","Adi","Sari","Lili","Slamet","Rosa","Dewi","Iwan","Ricky","Anton","Linda","Wira","Rian","Mega","Aris","Ria","Dina","Hermansyah","Rifa","Sinta","Yogi","Gita"];
const lastNames = ["Rustandi","Aminah","Ramadhan","Kusnadi","Ratnasari","Santoso","Suryana","Setiawan","Dahliawati","Ruhian","Wijaya","Amalia","Pratama","Wijayanto","Hidayat","Saputra","Wibowo","Prasetyo","Kurniawan","Sukma","Purnama","Siregar","Harahap","Nugroho","Putra","Putri","Siregar","Santika","Fauzi","Rahma","Sari"];

function randomInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function randomName() {
  const f = firstNames[Math.floor(Math.random() * firstNames.length)];
  const l = lastNames[Math.floor(Math.random() * lastNames.length)];
  return `${f} ${l}`.toUpperCase();
}

function randomNIK(seed) {
  // 16 digit pseudo-unique number
  const base = Date.now().toString().slice(-6) + String(seed).padStart(6, '0');
  const rand = String(Math.floor(Math.random() * 1e4)).padStart(4, '0');
  return (base + rand).slice(0, 16);
}

function generateDataset(n, outPath) {
  const arr = [];
  for (let i = 0; i < n; i++) {
    arr.push({
      nama: randomName(),
      nik: randomNIK(i),
      kk: randomNIK(i + 1000000),
      jenis_kelamin: Math.random() < 0.5 ? 'Laki-laki' : 'Perempuan',
      tempat_lahir: 'Bandung',
      tanggal_lahir: `${String(randomInt(1,28)).padStart(2,'0')}-${String(randomInt(1,12)).padStart(2,'0')}-${randomInt(1950,2010)}`,
      agama: 'Islam',
      pendidikan: 'SMA',
      pekerjaan: 'Pekerjaan',
      status_pernikahan: Math.random() < 0.5 ? 'Kawin' : 'Belum Kawin',
      alamat: 'Jl. Contoh No. ' + randomInt(1,200),
      rt: String(randomInt(1,10)).padStart(2,'0'),
      rw: '02',
      status: 'Aktif'
    });
  }
  fs.writeFileSync(outPath, JSON.stringify(arr, null, 2), 'utf8');
  console.log(`Generated ${n} records -> ${outPath}`);
}

if (require.main === module) {
  const sizes = [1000, 10000, 100000];
  const outDir = path.join(__dirname, 'data');
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  for (const s of sizes) {
    const outPath = path.join(outDir, `dataset_warga_${s}.json`);
    generateDataset(s, outPath);
  }
}

module.exports = { generateDataset };
