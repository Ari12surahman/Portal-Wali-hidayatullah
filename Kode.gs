// --- ID SPREADSHEET (Sesuai Chat Bapak) ---
var SHEET_ID = "1dD_1TzjLht36J_H7P6d1fovHhMH31ZsVJZCPHlkiAjs"; 
var SHEET_DB = "Data_Santri"; // Ganti target login ke Data_Santri
var SHEET_LAPORAN = "Laporan";
var SHEET_CONFIG = "Pengaturan";

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle("Portal Wali Santri")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppConfig() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_CONFIG);
    var config = { appName: "Portal Wali Santri", logoUrl: "" };

    if (sheet) {
      var valName = sheet.getRange("B2").getValue();
      var valLogo = sheet.getRange("B3").getValue();
      if(valName) config.appName = valName;
      if(valLogo) config.logoUrl = valLogo;
    }
    return config;
  } catch (e) {
    return { appName: "Portal Wali Santri", logoUrl: "" };
  }
}

// Helper Format Nomor WA ke +62
function formatNomorWA(no) {
  if (!no) return "";
  var n = String(no).trim().replace(/[^0-9+]/g, "");
  if (!n) return "";
  if (n.startsWith("+62")) return n;
  if (n.startsWith("62")) return "+" + n;
  if (n.startsWith("0")) return "+62" + n.substring(1);
  if (n.startsWith("8")) return "+62" + n;
  return n;
}

// --- LOGIN (CEK SHEET DATA_SANTRI: BISA NAMA, USERNAME, ATAU NO WA) ---
function loginWali(u, p) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_DB); // Data_Santri
    
    if(!sheet) return {status: 'GAGAL', msg: 'Database tidak ditemukan!'};
    
    // Ambil Data: Nama(0), Kelas(1), Musyrif(2), Password(3), Target(4), No WA(5), Username(6)
    var data = sheet.getDataRange().getValues(); 
    
    var inputUser = String(u).trim().toLowerCase();
    var inputWA = formatNomorWA(u);
    var inputPass = String(p).trim();

    // Loop mulai baris 1 (Header dilewati)
    for(var i=1; i<data.length; i++) {
      var dbNama = String(data[i][0]).trim();         // Kolom A: Nama
      var dbPass = String(data[i][3]).trim();         // Kolom D: Password (Index 3)
      var dbNoWa = formatNomorWA(data[i][5] || '');   // Kolom F: No WA Wali (Index 5)
      var dbUserWali = String(data[i][6] || '').trim().toLowerCase(); // Kolom G: Username Wali (Index 6)

      // Cek apakah ada kecocokan Nama, Username Wali, atau No WA Wali & Password benar
      var matchUser = (dbNama.toLowerCase() === inputUser) || 
                      (dbUserWali !== '' && dbUserWali === inputUser) || 
                      (dbNoWa !== '' && (dbNoWa === inputWA || dbNoWa === String(u).trim()));

      if(matchUser && dbPass === inputPass) {
        return { status: 'SUKSES', namaSantri: dbNama };
      }
    }
    return {status: 'GAGAL', msg: 'Username / No. WA / Nama Santri atau Password Salah.'};
  } catch(e) {
    return {status: 'ERROR', msg: 'Error: ' + e.toString()};
  }
}

// --- RIWAYAT (DENGAN ROW INDEX UNTUK EDIT/HAPUS) ---
function getRiwayatAnak(namaSantri, tglFilter) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_LAPORAN);
    if (!sheet) return [];
    
    // Pakai DisplayValues biar tanggal terbaca sebagai teks (dd/MM/yyyy)
    var data = sheet.getDataRange().getDisplayValues();
    var result = [];
    
    // Loop dari bawah (Terbaru)
    for(var i = data.length - 1; i >= 1; i--) {
      if(result.length >= 100) break; // Max 100 data
      
      var row = data[i];
      // STRUKTUR BARU (Sesuai Admin):
      // 0:Tanggal, 1:Mulai, 2:Selesai, 3:Nama, 4:Jenis, 
      // 5:Surat, 6:HalAwal, 7:HalAkhir, 8:AyatAwal, 9:AyatAkhir,
      // 10:Nilai, 11:Catatan, 12:Musyrif
      
      var dbNama = String(row[3]).trim().toLowerCase(); // Kolom D (Index 3)
      var targetNama = String(namaSantri).trim().toLowerCase();
      
      if(dbNama === targetNama) {
        
        // Cek Filter Tanggal
        // row[0] formatnya "30/12/2025" -> Ubah ke "2025-12-30" utk filter
        var tglSheet = row[0];
        var parts = tglSheet.split('/');
        var tglISO = "";
        if(parts.length === 3) tglISO = parts[2] + "-" + parts[1] + "-" + parts[0];
        else tglISO = tglSheet;

        if(tglFilter && tglISO !== tglFilter) continue;

        // Push Data ke Frontend (dengan rowIndex untuk edit/hapus)
        result.push({
          rowIndex: i + 1, // 1-indexed untuk sheet
          waktu: row[0], // Tampilkan dd/MM/yyyy
          jam: row[1] + " - " + row[2], // Tampilkan jam
          nama: row[3],
          jenis: row[4],
          surat: row[5],
          halAwal: row[6],
          halAkhir: row[7],
          ayatAwal: row[8],
          ayatAkhir: row[9],
          nilai: row[10],
          catatan: row[11],
          pelapor: row[12] // Musyrif ada di kolom M (Index 12)
        });
      }
    }
    return result;
  } catch (e) {
    return [];
  }
}

// --- GET ABSENSI SANTRI ---
function getAbsensiSantri(namaSantri, tglFilter) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Absensi");
    if (!sheet) return [];
    
    // Ambil Data: Tanggal(0), Waktu(1), Nama(2), Status(3), Musyrif(4)
    var data = sheet.getDataRange().getDisplayValues();
    var result = [];
    
    // Loop dari bawah (Terbaru)
    for(var i = data.length - 1; i >= 1; i--) {
      // Limit 50 data terakhir yang ditampilkan
      if(result.length >= 50) break; 
      
      var row = data[i];
      var dbNama = String(row[2]).trim().toLowerCase();
      var targetNama = String(namaSantri).trim().toLowerCase();
      
      if(dbNama === targetNama) {
        
        // Filter Tanggal (jika user milih tanggal tertentu di frontend)
        if(tglFilter) {
           var tglSheet = row[0];
           var parts = tglSheet.split('/');
           var tglISO = "";
           // Convert dd/MM/yyyy -> yyyy-MM-dd
           if(parts.length === 3) tglISO = parts[2] + "-" + parts[1] + "-" + parts[0];
           else tglISO = tglSheet;
           
           if(tglISO !== tglFilter) continue;
        }

        result.push({
          tanggal: row[0],
          waktu: row[1],
          status: row[3], // H, S, I, A
          musyrif: row[4]
        });
      }
    }
    return result;
  } catch (e) {
    return [];
  }
}

// --- GET PROGRES HAFALAN & STATISTIK ABSENSI REAL-TIME ---
function getProgresDanTargetWali(namaSantri) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var targetNama = String(namaSantri).trim().toLowerCase();
    
    // 1. Ambil data santri dari Data_Santri (Nama, Kelas, Musyrif, Target)
    var shSantri = ss.getSheetByName(SHEET_DB);
    var kelasSantri = "";
    var musyrifSantri = "";
    var targetIndividual = "";
    if (shSantri) {
      var dSantri = shSantri.getDataRange().getValues();
      for (var i = 1; i < dSantri.length; i++) {
        var dbNm = String(dSantri[i][0]).trim().toLowerCase();
        if (dbNm === targetNama) {
          kelasSantri = String(dSantri[i][1] || '').trim();
          musyrifSantri = String(dSantri[i][2] || '').trim();
          targetIndividual = String(dSantri[i][4] || '').trim();
          break;
        }
      }
    }

    // 2. Cari Target Default dari Target_Hafalan atau Master_Kelas jika targetIndividual kosong
    var targetStr = targetIndividual;
    if (!targetStr) {
      try {
        var shTarget = ss.getSheetByName("Target_Hafalan");
        if (shTarget && shTarget.getLastRow() >= 2) {
          var tData = shTarget.getDataRange().getValues();
          for (var t = 1; t < tData.length; t++) {
            var tTipe = String(tData[t][1]).trim().toLowerCase();
            var tUntuk = String(tData[t][2]).trim().toLowerCase();
            var tDetail = String(tData[t][4] || tData[t][3] || '').trim();
            if (tTipe === 'santri' && tUntuk === targetNama) {
              targetStr = tDetail; break;
            } else if (tTipe === 'kelas' && tUntuk === kelasSantri.toLowerCase() && !targetStr) {
              targetStr = tDetail;
            }
          }
        }
      } catch(e){}
    }
    if (!targetStr && kelasSantri) {
      try {
        var shKelas = ss.getSheetByName("Master_Kelas");
        if (shKelas && shKelas.getLastRow() >= 2) {
          var kData = shKelas.getDataRange().getValues();
          for (var k = 1; k < kData.length; k++) {
            if (String(kData[k][0]).trim().toLowerCase() === kelasSantri.toLowerCase()) {
              targetStr = String(kData[k][1] || '').trim();
              break;
            }
          }
        }
      } catch(e){}
    }

    // 3. Ambil Master Quran untuk kalkulasi
    var mqList = [];
    try {
      var shMq = ss.getSheetByName("Master_Quran");
      if (shMq && shMq.getLastRow() >= 2) {
        var mqData = shMq.getRange(2, 1, shMq.getLastRow() - 1, 5).getValues();
        mqData.forEach(function(r) {
          if (r[1]) mqList.push({ juz: Number(r[0]), surat: String(r[1]).trim(), halaman: Number(r[3]) || 0 });
        });
      }
    } catch(e){}

    // 4. Kalkulasi statistik riwayat setoran dari Laporan
    var stats = { surahs: [], hal: 0, ayat: 0, count: 0, rangesPerSurat: {}, uniquePages: {}, setoranList: [] };
    try {
      var shLapor = ss.getSheetByName(SHEET_LAPORAN);
      if (shLapor && shLapor.getLastRow() >= 2) {
        var lapData = shLapor.getRange(2, 4, shLapor.getLastRow() - 1, 7).getValues();
        lapData.forEach(function(r) {
          var nm = String(r[0]).trim().toLowerCase();
          if (nm !== targetNama) return;
          stats.count++;
          var srt = String(r[2]).trim();
          var rowJenis = String(r[1]).trim().toUpperCase();
          if (srt && srt !== '-' && srt !== '') {
            if (stats.surahs.indexOf(srt) === -1) stats.surahs.push(srt);
          }
          var hAwal = parseInt(r[3]) || 0;
          var hAkhir = parseInt(r[4]) || hAwal;
          if (hAwal === 0 && hAkhir === 0 && srt.toLowerCase().indexOf('juz') !== -1) {
            var mJ = srt.match(/juz\s*(\d+)/i);
            if (mJ && mJ[1]) {
              var jR = getJuzRange(parseInt(mJ[1]));
              if (jR) { hAwal = jR.start; hAkhir = jR.end; }
            }
          }
          if (hAwal > 0 || hAkhir > 0) {
            for (var pg = Math.min(hAwal, hAkhir); pg <= Math.max(hAwal, hAkhir); pg++) {
              stats.uniquePages[pg] = true;
            }
          }
          var aAwal = parseInt(r[5]) || 0;
          var aAkhir = parseInt(r[6]) || aAwal;
          stats.setoranList.push({ jenis: rowJenis, surat: srt, halAwal: hAwal, halAkhir: hAkhir, ayatAwal: aAwal, ayatAkhir: aAkhir });

          if (rowJenis === 'Z' && srt && aAwal > 0 && aAkhir > 0) {
            if (!stats.rangesPerSurat[srt]) stats.rangesPerSurat[srt] = [];
            stats.rangesPerSurat[srt].push([aAwal, aAkhir]);
          } else {
            if (aAwal || aAkhir) stats.ayat += Math.abs(aAkhir - aAwal) + 1;
            else if (stats.count) stats.ayat += 10;
          }
        });
        stats.hal = Object.keys(stats.uniquePages || {}).length;
        for (var surat in stats.rangesPerSurat) {
          var merged = mergeRanges(stats.rangesPerSurat[surat]);
          stats.ayat += hitungTotalAyatDariRanges(merged);
        }
      }
    } catch(e){}

    var prog = kalkulasiProgressTarget(targetStr, stats, mqList);
    var stStatus = "Baru Mulai";
    if (prog.persen >= 100 || String(prog.sisa).toLowerCase().includes('selesai') || String(prog.sisa).toLowerCase().includes('tercapai')) stStatus = "Tercapai!";
    else if (prog.persen >= 75) stStatus = "Sesuai Target";
    else if (prog.persen >= 40) stStatus = "Sudah Setengahnya";
    else stStatus = "Baru Mulai";

    prog.status = stStatus;
    prog.targetDisplay = targetStr || "Target Reguler / Standar";
    prog.kelas = kelasSantri || "-";
    prog.musyrif = musyrifSantri || "-";

    // 5. Statistik Absensi
    var absenStats = { H: 0, S: 0, I: 0, A: 0, total: 0 };
    try {
      var shAbs = ss.getSheetByName("Absensi");
      if (shAbs && shAbs.getLastRow() >= 2) {
        var absData = shAbs.getDataRange().getValues();
        for (var i = 1; i < absData.length; i++) {
          var absNm = String(absData[i][2]).trim().toLowerCase();
          if (absNm === targetNama) {
            var st = String(absData[i][3]).trim().toUpperCase();
            if (st === 'H' || st === 'S' || st === 'I' || st === 'A') {
              absenStats[st] = (absenStats[st] || 0) + 1;
              absenStats.total++;
            }
          }
        }
      }
    } catch(e){}

    prog.absensi = absenStats;
    return prog;
  } catch(e) {
    return { persen: 0, tercapai: "0 Surat", sisa: "Gagal memuat progres", targetDisplay: "-", status: "Baru Mulai", absensi: { H: 0, S: 0, I: 0, A: 0 } };
  }
}

// --- HELPER KALKULASI PROGRES TERSTRUKTUR ---
function getJuzRange(juzNum) {
  var j = parseInt(juzNum);
  if (j === 1) return { start: 1, end: 21, totalHal: 21, nama: "Juz 1" };
  if (j === 30) return { start: 582, end: 604, totalHal: 23, nama: "Juz 30" };
  if (j >= 2 && j <= 29) {
    var s = ((j - 1) * 20) + 2;
    var e = (j * 20) + 1;
    return { start: s, end: e, totalHal: 20, nama: "Juz " + j };
  }
  return null;
}

function kalkulasiProgressTarget(targetStr, stats, mqList) {
  var srtCount = stats.surahs ? stats.surahs.length : 0;
  var halCount = stats.hal || 0;
  var ayatCount = stats.ayat || 0;
  var setoranList = stats.setoranList || [];
  
  var persen = 0;
  var tercapai = "";
  var sisa = "";
  var tTrim = String(targetStr || '').trim();
  var tLower = tTrim.toLowerCase();
  
  if (!tTrim) {
    persen = Math.min(100, Math.round(((stats.count || 0) / 20) * 100));
    tercapai = srtCount + " Surat (" + halCount + " Hal)";
    sisa = "Belum ada target khusus dirinci";
    return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
  }

  // 1. TERSTRUKTUR / HEURISTIK JUZ
  if (tTrim.indexOf("JUZ:") === 0 || tLower.indexOf("juz") !== -1) {
    var matchJuz = tTrim.match(/JUZ:(\d+)/i) || tLower.match(/juz\s*(\d+)/i);
    if (matchJuz && matchJuz[1]) {
      var numJuz = parseInt(matchJuz[1]);
      var juzInfo = getJuzRange(numJuz);
      if (juzInfo) {
        var pagesDone = {};
        setoranList.forEach(function(item) {
          if (item.halAwal > 0 || item.halAkhir > 0) {
            var ha = item.halAwal || item.halAkhir;
            var hk = item.halAkhir || item.halAwal;
            for (var p = Math.min(ha, hk); p <= Math.max(ha, hk); p++) {
              if (p >= juzInfo.start && p <= juzInfo.end) {
                pagesDone[p] = true;
              }
            }
          }
        });
        var completedHal = Object.keys(pagesDone).length;
        if (completedHal === 0 && srtCount > 0 && (numJuz === 30 || numJuz === 29)) {
          var maxS = numJuz === 30 ? 37 : 11;
          persen = Math.min(100, Math.round((srtCount / maxS) * 100));
          tercapai = srtCount + " dari " + maxS + " Surat (" + completedHal + " / " + juzInfo.totalHal + " Hal)";
          var sisaS = Math.max(0, maxS - srtCount);
          sisa = sisaS > 0 ? ("Sisa " + sisaS + " surat lagi di " + juzInfo.nama) : ("Target " + juzInfo.nama + " Selesai! 🎉");
        } else {
          persen = Math.min(100, Math.round((completedHal / juzInfo.totalHal) * 100));
          tercapai = completedHal + " dari " + juzInfo.totalHal + " Halaman (" + Math.round(persen) + "%)";
          var sisaH = Math.max(0, juzInfo.totalHal - completedHal);
          sisa = sisaH > 0 ? ("Sisa " + sisaH + " halaman lagi di " + juzInfo.nama) : ("Target " + juzInfo.nama + " Selesai! 🎉");
        }
        return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: completedHal, ayat: ayatCount };
      }
    }
    var maxHal = 600;
    persen = Math.min(100, Math.round((halCount / maxHal) * 100));
    tercapai = halCount + " dari " + maxHal + " Halaman (" + srtCount + " Surat)";
    sisa = halCount >= maxHal ? "Target 30 Juz Selesai! 🎉" : ("Sisa " + Math.max(0, maxHal - halCount) + " halaman lagi");
    return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
  }

  // 2. TERSTRUKTUR SURAT / RENTANG SURAT
  if (tTrim.indexOf("SURAT:") === 0 || tLower.indexOf("surat") !== -1 || tLower.indexOf("surah") !== -1) {
    var parts = tTrim.split("|");
    if (tTrim.indexOf("SURAT:") === 0 && parts.length >= 3 && mqList && mqList.length > 0) {
      var sStart = parts[1].trim();
      var sEnd = parts[2].split(" ")[0].trim();
      var pStart = 604, pEnd = 1;
      mqList.forEach(function(m) {
        if (isSuratMatch(m.surat, sStart)) { if (m.halaman < pStart) pStart = m.halaman; }
        if (isSuratMatch(m.surat, sEnd)) { if (m.halaman > pEnd) pEnd = m.halaman; }
      });
      if (pStart <= pEnd) {
        var totalH = (pEnd - pStart) + 1;
        var pagesDone = {};
        setoranList.forEach(function(item) {
          if (item.halAwal > 0 || item.halAkhir > 0) {
            for (var p = Math.min(item.halAwal, item.halAkhir); p <= Math.max(item.halAwal, item.halAkhir); p++) {
              if (p >= pStart && p <= pEnd) pagesDone[p] = true;
            }
          }
        });
        var completed = Object.keys(pagesDone).length;
        persen = Math.min(100, Math.round((completed / totalH) * 100));
        tercapai = completed + " dari " + totalH + " Halaman (" + sStart + " s/d " + sEnd + ")";
        var sisaH = Math.max(0, totalH - completed);
        sisa = sisaH > 0 ? ("Sisa " + sisaH + " halaman lagi untuk rentang surah ini") : "Target Rentang Surah Selesai! 🎉";
        return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: completed, ayat: ayatCount };
      }
    }
    var matchS = tLower.match(/(\d+)/);
    var maxS = matchS ? parseInt(matchS[1]) : 10;
    persen = Math.min(100, Math.round((srtCount / maxS) * 100));
    tercapai = srtCount + " dari " + maxS + " Surat";
    sisa = srtCount >= maxS ? "Target Surat Selesai! 🎉" : ("Sisa " + (maxS - srtCount) + " surat lagi");
    return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
  }

  // 3. TERSTRUKTUR HALAMAN / TAHAP
  if (tTrim.indexOf("HAL:") === 0 || tLower.indexOf("hal") !== -1 || tLower.indexOf("lembar") !== -1 || tLower.indexOf("hlm") !== -1) {
    var matchHal = tTrim.match(/HAL:(\d+)/i) || tLower.match(/(\d+)/);
    var maxH = matchHal ? parseInt(matchHal[1]) : 10;
    if (tLower.indexOf("lembar") !== -1) maxH = maxH * 2;
    persen = Math.min(100, Math.round((halCount / maxH) * 100));
    tercapai = halCount + " dari " + maxH + " Halaman";
    sisa = halCount >= maxH ? "Target Halaman Selesai! 🎉" : ("Sisa " + (maxH - halCount) + " halaman lagi");
    return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
  }

  if (tTrim.indexOf("TAHAP:") === 0 || tLower.indexOf("tahap") !== -1 || tLower.indexOf("jilid") !== -1 || tLower.indexOf("iqro") !== -1) {
    var maxJ = 30;
    persen = Math.min(100, Math.round((halCount / maxJ) * 100));
    tercapai = halCount + " Halaman / Bab tuntas";
    sisa = halCount >= maxJ ? "Target Jilid/Tahap Selesai! 🎉" : ("Sisa " + (maxJ - halCount) + " halaman/latihan lagi");
    return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
  }

  // 4. GENERAL / LAINNYA
  var maxH = 20;
  var maxA = 300;
  persen = halCount > 0 ? Math.min(100, Math.round((halCount / maxH) * 100)) : Math.min(100, Math.round((ayatCount / maxA) * 100));
  tercapai = srtCount + " Surat (" + halCount + " Hal / " + ayatCount + " Ayat)";
  var sisaH = Math.max(0, maxH - halCount);
  var sisaA = Math.max(0, maxA - ayatCount);
  sisa = (halCount >= maxH || ayatCount >= maxA) ? "Target Selesai / Tercapai! 🎉" : ("Sisa " + sisaH + " Halaman (" + sisaA + " Ayat) lagi");
  return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
}

function mergeRanges(ranges) {
  if (!ranges || ranges.length === 0) return [];
  var sorted = ranges.slice().sort(function(a, b) { return a[0] - b[0]; });
  var merged = [sorted[0].slice()];
  for (var i = 1; i < sorted.length; i++) {
    var last = merged[merged.length - 1];
    if (sorted[i][0] <= last[1] + 1) {
      last[1] = Math.max(last[1], sorted[i][1]);
    } else {
      merged.push(sorted[i].slice());
    }
  }
  return merged;
}

function hitungTotalAyatDariRanges(mergedRanges) {
  var total = 0;
  for (var i = 0; i < mergedRanges.length; i++) {
    total += mergedRanges[i][1] - mergedRanges[i][0] + 1;
  }
  return total;
}

function normalizeSurat(str) {
  if (!str) return "";
  var s = String(str).toLowerCase().trim();
  s = s.replace(/^(?:q\.?s\.?|sura[th]?|quran sura[th]?)\s*(?:-|_|\s)*\s*/i, "").trim();
  if (s.indexOf("lahab") !== -1 || s.indexOf("masad") !== -1) return "almasad";
  if (s.indexOf("inshirah") !== -1 || s.indexOf("insyirah") !== -1 || s.indexOf("syarh") !== -1 || s.indexOf("alam nas") !== -1) return "alinsyirah";
  if (s.indexOf("isra") !== -1 || s.indexOf("bani israil") !== -1) return "alisra";
  if (s.indexOf("kahf") !== -1) return "alkahf";
  if (s.indexOf("ghafir") !== -1 || s.indexOf("mumin") !== -1) return "almumin";
  if (s.indexOf("fushilat") !== -1 || s.indexOf("fussilat") !== -1 || s.indexOf("ha mim") !== -1) return "fushilat";
  if (s.indexOf("imran") !== -1) return "aliimran";
  s = s.replace(/sh/g, "sy").replace(/th/g, "t").replace(/ts/g, "t").replace(/dz/g, "z").replace(/dh/g, "z").replace(/oo/g, "u").replace(/ee/g, "i");
  s = s.replace(/[^a-z0-9]/g, "").replace(/([a-z0-9])\1+/g, "$1").replace(/([aiueo])h$/i, "$1");
  return s;
}

function isSuratMatch(str1, str2) {
  if (!str1 || !str2) return false;
  var n1 = normalizeSurat(str1);
  var n2 = normalizeSurat(str2);
  if (!n1 || !n2) return false;
  if (n1 === n2) return true;
  if (n1.length >= 4 && n2.length >= 4) {
    if (n1.indexOf(n2) !== -1 || n2.indexOf(n1) !== -1) return true;
  }
  return false;
}


// --- API BRIDGE UNTUK VERCEL ---
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var funcName = payload.func;
    var args = payload.args || [];
    
    if (typeof this[funcName] !== 'function') {
      throw new Error('Fungsi ' + funcName + ' tidak ditemukan di backend.');
    }
    
    var result = this[funcName].apply(this, args);
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
// --- END API BRIDGE ---
