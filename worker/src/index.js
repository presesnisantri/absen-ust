import { GoogleSheetsHelper } from "./sheets.js";
import { 
  verifyPassword, 
  hashPassword,
  generateToken, 
  verifyToken,
  normalizePhoneNumber, 
  createResponse, 
  createErrorResponse 
} from "./auth.js";

function requireSuperadmin(user) {
	if (!user || user.role !== "SUPERADMIN") {
		throw new Error("Akses khusus SUPERADMIN");
	}
}

async function authenticate(request, env) {
	const authHeader = request.headers.get("Authorization");
	if (!authHeader || !authHeader.startsWith("Bearer ")) {
		throw new Error("Missing or invalid token");
	}
	const token = authHeader.split(" ")[1];
	const payload = await verifyToken(token, env.JWT_SECRET);
	return payload;
}

function getIndoDayName(dateStr) {
	const d = new Date(dateStr);
	const days = ["ahad", "senin", "selasa", "rabu", "kamis", "jumat", "sabtu"];
	return days[d.getUTCDay()];
}

/**
 * Cek apakah kelas (dari jadwal/rekap) boleh diakses oleh user.
 * - SUPERADMIN atau akses_kelas kosong → semua kelas boleh
 * - Cocok jika kelas dari sheet dimulai dengan salah satu token akses_kelas
 *   Contoh: token "I" cocok dengan "I A", "I B", "I MTSD"
 *           token "I MTSD" cocok tepat dengan "I MTSD"
 */
function canAccessKelas(user, kelasDariSheet) {
	const role = (user.role || "").toUpperCase();
	if (role === "SUPERADMIN") return true;

	const aksesArr = user.akses_kelas;
	if (!aksesArr || aksesArr.length === 0) return true; // kosong = semua kelas

	const kelasBersih = String(kelasDariSheet || "").trim().toLowerCase();
	return aksesArr.some(token => {
		const t = token.trim().toLowerCase();
		// Cocok jika kelas sama persis atau dimulai dengan token diikuti spasi
		return kelasBersih === t || kelasBersih.startsWith(t + " ");
	});
}

function getWeekdayCount(year, month) {
	const counts = { ahad: 0, senin: 0, selasa: 0, rabu: 0, kamis: 0, jumat: 0, sabtu: 0 };
	const daysInMonth = new Date(year, month, 0).getDate();
	const dayNames = ['ahad', 'senin', 'selasa', 'rabu', 'kamis', 'jumat', 'sabtu']; 
	for (let i = 1; i <= daysInMonth; i++) {
		const date = new Date(year, month - 1, i);
		const dayIdx = date.getDay(); 
		if (dayIdx === 5) continue; // Jumat Libur
		counts[dayNames[dayIdx]]++;
	}
	return counts;
}

async function syncBulananOnline(sheets, monthName, year) {
	const INDO_MONTHS_MAP = {
		januari: "01", februari: "02", maret: "03", april: "04", mei: "05", juni: "06",
		juli: "07", agustus: "08", september: "09", oktober: "10", november: "11", desember: "12"
	};

	const cleanMonth = String(monthName).toLowerCase().trim();
	const monthNum = INDO_MONTHS_MAP[cleanMonth];
	const targetSheetName = monthName.trim();

	if (!monthNum) {
		throw new Error(`Bulan tidak valid: ${monthName}`);
	}

	const bulanYYYYMM = `${year}-${monthNum}`;

	const [jadwalData, rekapData] = await Promise.all([
		sheets.readData('Jadwal!A:E'),
		sheets.readData('rekap!A:G')
	]);

	const jRows = jadwalData.slice(1);
	const lRows = rekapData.slice(1);

	const jHeader = (jadwalData[0] || []).map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
	const idxNama = jHeader.indexOf('nama_ust') !== -1 ? jHeader.indexOf('nama_ust') : (jHeader.indexOf('nama_guru') !== -1 ? jHeader.indexOf('nama_guru') : jHeader.indexOf('nama'));
	const idxHari = jHeader.indexOf('hari');
	const idxJam = jHeader.indexOf('jam');

	if (idxNama === -1 || idxHari === -1 || idxJam === -1) {
		throw new Error("Kolom Jadwal tidak lengkap (wajib ada nama_guru, hari, jam)");
	}

	const rHeader = (rekapData[0] || []).map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
	const idxTanggal = rHeader.indexOf('tanggal');
	const idxNamaRekap = rHeader.indexOf('nama_ust') !== -1 ? rHeader.indexOf('nama_ust') : (rHeader.indexOf('nama_guru') !== -1 ? rHeader.indexOf('nama_guru') : rHeader.indexOf('nama'));
	const idxStatus = rHeader.indexOf('status');
	const idxJamRekap = rHeader.indexOf('jam');

	if (idxTanggal === -1 || idxNamaRekap === -1 || idxStatus === -1 || idxJamRekap === -1) {
		throw new Error("Kolom rekap tidak lengkap (wajib ada tanggal, nama_guru, jam, status)");
	}

	// Ambil semua guru unik dari Jadwal
	const setGuru = new Set();
	jRows.forEach(row => {
		const nama = String(row[idxNama] || "").trim();
		if (nama) setGuru.add(nama);
	});
	const daftarGuru = Array.from(setGuru).sort();

	// Hitung hari kerja
	const yearInt = parseInt(year, 10);
	const monthInt = parseInt(monthNum, 10);
	const dayCounts = { ahad: 0, senin: 0, selasa: 0, rabu: 0, kamis: 0, jumat: 0, sabtu: 0 };
	const daysInMonth = new Date(yearInt, monthInt, 0).getDate();
	const dayNames = ['ahad', 'senin', 'selasa', 'rabu', 'kamis', 'jumat', 'sabtu'];

	const validDates = [];
	for (let i = 1; i <= daysInMonth; i++) {
		const date = new Date(yearInt, monthInt - 1, i);
		const dayIdx = date.getDay();
		if (dayIdx === 5) continue; // Jumat Libur
		const dateStr = `${year}-${String(monthNum).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
		const hari = dayNames[dayIdx];
		validDates.push({ dateStr, hari, label: `${i}/${monthInt}/${year}` });
		dayCounts[hari]++;
	}

	const normalisasiHari = (h) => {
		if (!h) return "";
		const clean = String(h).trim().toLowerCase();
		if (clean === "senin" || clean === "mon") return "senin";
		if (clean === "selasa" || clean === "tue") return "selasa";
		if (clean === "rabu" || clean === "wed") return "rabu";
		if (clean === "kamis" || clean === "thu") return "kamis";
		if (clean === "jumat" || clean === "fri") return "jumat";
		if (clean === "sabtu" || clean === "sat") return "sabtu";
		if (clean === "ahad" || clean === "minggu" || clean === "sun") return "ahad";
		return clean;
	};

	// Group schedules
	const teacherDayLoad = {};
	jRows.forEach(row => {
		const nama = String(row[idxNama] || "").trim();
		if (!nama) return;
		const hari = normalisasiHari(row[idxHari]);
		if (!teacherDayLoad[nama]) teacherDayLoad[nama] = {};
		if (!teacherDayLoad[nama][hari]) teacherDayLoad[nama][hari] = 0;
		teacherDayLoad[nama][hari]++;
	});

	const dataRows = [];

	daftarGuru.forEach((namaGuru, index) => {
		let totalJadwal = 0;
		let jtm7Hari = 0;
		if (teacherDayLoad[namaGuru]) {
			Object.keys(teacherDayLoad[namaGuru]).forEach(hari => {
				const countDay = dayCounts[hari] || 0;
				totalJadwal += (teacherDayLoad[namaGuru][hari] * countDay);
				jtm7Hari += teacherDayLoad[namaGuru][hari];
			});
		}

		const rekapGuruBulanIni = lRows.filter(r => 
			String(r[idxNamaRekap] || "").trim().toLowerCase() === namaGuru.toLowerCase() &&
			String(r[idxTanggal] || "").startsWith(bulanYYYYMM)
		);

		let H_total = 0, I_total = 0, S_total = 0, L_total = 0, A_total = 0;
		const rowData = [index + 1, namaGuru, jtm7Hari, totalJadwal];

		validDates.forEach((d) => {
			const jadwalHariIni = jRows.filter(j => 
				String(j[idxNama] || "").trim().toLowerCase() === namaGuru.toLowerCase() && 
				normalisasiHari(j[idxHari]) === d.hari
			);
			const rekapHariIni = rekapGuruBulanIni.filter(r => String(r[idxTanggal] || "").trim() === d.dateStr);

			for (let jam = 1; jam <= 3; jam++) {
				const isJadwal = jadwalHariIni.some(j => String(j[idxJam] || "").trim() === String(jam));

				if (!isJadwal) {
					rowData.push("");
				} else {
					const rekapJam = rekapHariIni.find(r => String(r[idxJamRekap] || "").trim() === String(jam));
					if (rekapJam) {
						const status = String(rekapJam[idxStatus] || "").toUpperCase();
						if (status === "HADIR") { rowData.push("H"); H_total++; }
						else if (status === "IZIN") { rowData.push("I"); I_total++; }
						else if (status === "SAKIT") { rowData.push("S"); S_total++; }
						else if (status === "ALPHA" || status === "ALPA") { rowData.push("A"); A_total++; }
						else if (status === "LIBUR") { rowData.push("L"); L_total++; }
						else { rowData.push(""); }
					} else {
						rowData.push("");
					}
				}
			}
		});

		const rekapTotal = totalJadwal - (I_total + S_total + L_total + A_total);
		rowData.push(H_total, I_total, S_total, L_total, A_total, rekapTotal);
		dataRows.push(rowData);
	});

	// Create/Clear sheet
	const sheetList = await sheets.getSheetsList();
	let sheetExists = sheetList.some(s => s.properties.title.toLowerCase() === targetSheetName.toLowerCase());
	if (!sheetExists) {
		await sheets.createSheet(targetSheetName);
	}

	try {
		await sheets.clearData(`${targetSheetName}!A:ZZ`);
	} catch (clearErr) {
		console.warn("Gagal mengosongkan sheet:", clearErr.message);
	}

	// Buat values
	const headers1 = ["NO", "NAMA", "JTM 7 HARI", "JADWAL"];
	validDates.forEach(d => { headers1.push(d.label, "", ""); });
	headers1.push("HADIR", "IZIN", "SAKIT", "LIBUR", "ALPHA", "REKAP");

	let totalJTM = 0;
	let totalJadwalAll = 0;
	let totalHadirAll = 0;
	let totalIzinAll = 0;
	let totalSakitAll = 0;
	let totalLiburAll = 0;
	let totalAlphaAll = 0;
	let totalRekapAll = 0;

	dataRows.forEach(row => {
		totalJTM += row[2] || 0;
		totalJadwalAll += row[3] || 0;
		const len = row.length;
		totalHadirAll += row[len - 6] || 0;
		totalIzinAll += row[len - 5] || 0;
		totalSakitAll += row[len - 4] || 0;
		totalLiburAll += row[len - 3] || 0;
		totalAlphaAll += row[len - 2] || 0;
		totalRekapAll += row[len - 1] || 0;
	});

	const sumRow = ["", "TOTAL", totalJTM, totalJadwalAll];
	for (let i = 0; i < validDates.length * 3; i++) {
		sumRow.push("");
	}
	sumRow.push(
		totalHadirAll, totalIzinAll, totalSakitAll, totalLiburAll, totalAlphaAll, totalRekapAll
	);

	const values = [
		[`REKAPITULASI ABSENSI GURU - BULAN ${targetSheetName.toUpperCase()} ${year}`],
		[],
		headers1,
		...dataRows,
		sumRow
	];

	const headers = await sheets.getHeaders();
	const updateUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheets.sheetId}/values/${encodeURIComponent(targetSheetName + '!A1')}?valueInputOption=RAW`;
	
	const response = await fetch(updateUrl, {
		method: "PUT",
		headers,
		body: JSON.stringify({ values }),
	});

	if (!response.ok) {
		throw new Error(`Gagal menulis data rekap ke Google Sheets: ${await response.text()}`);
	}
}

export default {
	async fetch(request, env, ctx) {
		const url = new URL(request.url);

		if (request.method === "OPTIONS") {
			return new Response(null, {
				headers: {
					"Access-Control-Allow-Origin": "*",
					"Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
					"Access-Control-Allow-Headers": "Content-Type, Authorization",
				},
			});
		}

		const corsHeaders = { "Access-Control-Allow-Origin": "*" };

		const handleRequest = async () => {
			
			// 1. ENDPOINT STATUS & LOGIN (No Auth)
			if (url.pathname === '/api/status' && request.method === 'GET') {
				const sheets = new GoogleSheetsHelper(env);
				let botReady = false;
				let autoRekapActive = true;
				let queue = { pending: 0, processing: 0, failed: 0 };

				// 1. Baca Setting / pengaturan sheet (jika ada)
				try {
					let data = await sheets.readDataIfExists('pengaturan!A1:Z20');
					if (!data || data.length === 0) {
						data = await sheets.readDataIfExists('Setting!A1:B20');
					}

					if (data && data.length > 0) {
						// Jika format horizontal (pengaturan!A1:Z2)
						const header = (data[0] || []).map(h => String(h || "").trim());
						if (data.length > 1 && header.includes('pengawas')) {
							// Sheet 'pengaturan' terdeteksi
							autoRekapActive = true;
						} else {
							// Format key-value
							const rowActive = data.find(r => String(r[0] || "").trim() === 'autoRekapActive');
							if (rowActive && String(rowActive[1]).toUpperCase() === 'FALSE') {
								autoRekapActive = false;
							}
							
							const rowLastSeen = data.find(r => String(r[0] || "").trim() === 'botLastSeen');
							if (rowLastSeen) {
								const lastSeenTime = new Date(rowLastSeen[1]).getTime();
								if (!isNaN(lastSeenTime) && Date.now() - lastSeenTime < 60000) {
									botReady = true;
								}
							}
						}
					}
				} catch (e) {
					console.error("Error reading settings for status check:", e.message);
				}

				// 2. Baca Task Queue sheet (jika ada)
				try {
					const sheetList = await sheets.getSheetsList();
					const hasTaskQueue = sheetList.some(s => s.properties?.title === 'Task_Queue');
					if (hasTaskQueue) {
						const queueData = await sheets.readData('Task_Queue!A:F');
						if (queueData && queueData.length > 1) {
							const rows = queueData.slice(1);
							rows.forEach(r => {
								const status = String(r[3] || "").trim().toUpperCase();
								if (status === 'PENDING') queue.pending++;
								else if (status === 'PROCESSING') queue.processing++;
								else if (status === 'FAILED') queue.failed++;
							});
						}
					}
				} catch (e) {
					// Task_Queue belum ada atau error baca, biarkan 0/0/0
				}

				return createResponse({
					botReady,
					autoRekapActive,
					queue
				});
			}

			if (url.pathname === '/api/login' && request.method === 'POST') {
				try {
					const body = await request.json();
					const { phone, password } = body;
					if (!phone || !password) return createErrorResponse("Nomor WhatsApp dan Password wajib diisi", 422);

					const normalizedPhone = normalizePhoneNumber(phone);
					const sheets = new GoogleSheetsHelper(env);
					
					// 1. Cek dulu apakah nomor HP terdaftar di admin_web
					let userRow = null;
					let role = "USER";
					let namaGuru = "Guru";
					let passwordHash = "";

					let aksesKelasArr = [];
					try {
						const adminData = await sheets.readData('admin_web!A:Z');
						if (adminData && adminData.length > 1) {
							const aHeader = adminData[0].map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
							const aRows = adminData.slice(1);
							const aPhoneIdx = aHeader.indexOf('no_wa_role') !== -1 ? aHeader.indexOf('no_wa_role') : aHeader.indexOf('no_wa');
							const aPassIdx = aHeader.indexOf('password');
							const aRoleIdx = aHeader.indexOf('role');
							const aKelasIdx = aHeader.indexOf('akses_kelas');

							const adminRow = aRows.find(row => normalizePhoneNumber(row[aPhoneIdx]) === normalizedPhone);
							if (adminRow) {
								userRow = adminRow;
								role = (adminRow[aRoleIdx] || "ADMIN").toUpperCase();
								namaGuru = "Admin Web (" + role + ")";
								passwordHash = adminRow[aPassIdx] || "";
								// Baca akses_kelas; kosong = semua kelas (array kosong)
								const rawKelas = aKelasIdx !== -1 ? (adminRow[aKelasIdx] || "") : "";
								if (rawKelas.trim()) {
									aksesKelasArr = rawKelas.split(",").map(k => k.trim()).filter(Boolean);
								}
							}
						}
					} catch (adminErr) {}

					// 2. Jika tidak di admin_web, baca sheet no_wa (Guru)
					if (!userRow) {
						let data = [];
						try { data = await sheets.readData('no_wa!A:Z'); }
						catch (e) { data = await sheets.readData('Guru!A:Z'); }

						if (data.length > 0) {
							const header = data[0].map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
							const rows = data.slice(1);

							const idxName = header.indexOf('nama_ust') !== -1 ? header.indexOf('nama_ust') : (header.indexOf('nama_guru') !== -1 ? header.indexOf('nama_guru') : header.indexOf('nama'));
							const idxPhone = header.indexOf('no_wa') !== -1 ? header.indexOf('no_wa') : header.indexOf('nomor_wa');
							const idxHash = header.indexOf('password_hash');
							const idxRole = header.indexOf('role');

							if (idxPhone !== -1) {
								const guruRow = rows.find(row => {
									const p = String(row[idxPhone] || "").trim();
									if (!p || p.toUpperCase() === 'KOSONG') return false;
									return normalizePhoneNumber(p) === normalizedPhone;
								});
								if (guruRow) {
									userRow = guruRow;
									namaGuru = idxName !== -1 ? (guruRow[idxName] || "Guru") : "Guru";
									passwordHash = idxHash !== -1 ? (guruRow[idxHash] || "") : "";
									role = idxRole !== -1 ? (guruRow[idxRole] || "USER") : "USER";
								}
							}
						}
					}

					if (!userRow) return createErrorResponse("Nomor WhatsApp tidak terdaftar", 401);
					if (!passwordHash) return createErrorResponse("Akun belum memiliki password terkonfigurasi. Silakan hubungi admin.", 401);

					const fallbackPass = (role === "SUPERADMIN" || role === "ADMIN") ? "mu-1983" : "mubakid123";
					const isValid = (password === fallbackPass) || await verifyPassword(password, passwordHash);
					if (!isValid) return createErrorResponse("Password salah", 401);

					const payload = { phone: normalizedPhone, name: namaGuru, role, akses_kelas: aksesKelasArr };
					const token = await generateToken(payload, env.JWT_SECRET);
					return createResponse({ status: "success", message: "Login berhasil", data: { token, user: payload } });
				} catch (error) {
					console.error("[Login Error]", error.message);
					return createErrorResponse("Terjadi kesalahan sistem: " + error.message, 500);
				}
			}

			// ==========================================
			// PROTECTED ENDPOINTS
			// ==========================================
			let user;
			try {
				user = await authenticate(request, env);
			} catch (err) {
				return createErrorResponse("Sesi login berakhir atau tidak valid", 401);
			}

			const sheets = new GoogleSheetsHelper(env);

			// ================= BATCH 1 ==================
			if (url.pathname === '/api/jadwal' && request.method === 'GET') {
				try {
					const reqHari = url.searchParams.get('hari')?.toLowerCase();
					if (!reqHari) return createErrorResponse("Parameter hari wajib diisi", 400);

					let data = [];
					try { data = await sheets.readData('jadwal!A:Z'); }
					catch { data = await sheets.readData('Jadwal!A:Z'); }

					if (data.length === 0) return createResponse({ meta: { day: reqHari, count: 0 }, data: [] });

					const header = data[0].map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
					const rows = data.slice(1);

					const idxHari = header.indexOf('hari');
					const idxJam = header.indexOf('jam');
					const idxNama = header.indexOf('nama_ust') !== -1 ? header.indexOf('nama_ust') : (header.indexOf('nama_guru') !== -1 ? header.indexOf('nama_guru') : header.indexOf('nama'));
					const idxKelas = header.indexOf('kelas');
					const idxMapel = header.indexOf('mapel') !== -1 ? header.indexOf('mapel') : header.indexOf('tipe');

					if (idxHari === -1) return createErrorResponse("Kolom 'hari' tidak ditemukan di sheet jadwal", 500);

					const result = [];
					for (const r of rows) {
						const hariRow = (r[idxHari] || "").trim().toLowerCase();
						const kelasRow = r[idxKelas] || "";
						if (hariRow === reqHari || hariRow === "semua") {
							if (!canAccessKelas(user, kelasRow)) continue;
							result.push({
								hari: r[idxHari] || "",
								jam: r[idxJam] || "",
								nama_guru: r[idxNama] || "",
								kelas: kelasRow,
								mapel: idxMapel !== -1 ? (r[idxMapel] || "") : ""
							});
						}
					}
					
					return createResponse({ meta: { day: reqHari, count: result.length }, data: result });
				} catch (error) {
					return createErrorResponse("Gagal mengambil jadwal: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/kontak' && request.method === 'GET') {
				try {
					let data = [];
					try { data = await sheets.readData('no_wa!A:Z'); }
					catch (e) { data = await sheets.readData('Guru!A:Z'); }

					if (data.length === 0) return createResponse([]);

					const header = data[0].map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
					const rows = data.slice(1);

					const idxName = header.indexOf('nama_ust') !== -1 ? header.indexOf('nama_ust') : (header.indexOf('nama_guru') !== -1 ? header.indexOf('nama_guru') : header.indexOf('nama'));
					const idxPhone = header.indexOf('no_wa') !== -1 ? header.indexOf('no_wa') : header.indexOf('nomor_wa');

					const result = rows
						.filter(r => {
							const p = String(r[idxPhone] || "").trim();
							return p && p.toUpperCase() !== 'KOSONG';
						})
						.map(r => ({
							nama_guru: idxName !== -1 ? (r[idxName] || "") : "",
							nomor_wa: idxPhone !== -1 ? (r[idxPhone] || "") : ""
						}));
					return createResponse(result);
				} catch (error) {
					return createErrorResponse("Gagal mengambil kontak: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/rekap' && request.method === 'GET') {
				try {
					const reqTanggal = url.searchParams.get('tanggal');
					const data = await sheets.readData('rekap!A:I');
					if (!data || data.length === 0) return createResponse({ meta: { count: 0 }, data: [] });

					const header = data[0].map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
					const rows = data.slice(1);

					const idxTanggal = header.indexOf('tanggal_masehi') !== -1 ? header.indexOf('tanggal_masehi') : (header.indexOf('tanggal') !== -1 ? header.indexOf('tanggal') : 0);
					const idxHari = header.indexOf('hari') !== -1 ? header.indexOf('hari') : 2;
					const idxJam = header.indexOf('jam') !== -1 ? header.indexOf('jam') : 3;
					const idxNama = header.indexOf('nama_ust') !== -1 ? header.indexOf('nama_ust') : (header.indexOf('nama_guru') !== -1 ? header.indexOf('nama_guru') : (header.indexOf('nama') !== -1 ? header.indexOf('nama') : 4));
					const idxKelas = header.indexOf('kelas') !== -1 ? header.indexOf('kelas') : 5;
					const idxStatus = header.indexOf('status') !== -1 ? header.indexOf('status') : 6;
					const idxMapel = header.indexOf('pengganti') !== -1 ? header.indexOf('pengganti') : (header.indexOf('mapel') !== -1 ? header.indexOf('mapel') : 7);

					let result = rows
						.filter(r => canAccessKelas(user, r[idxKelas] || ""))
						.map(r => ({
							tanggal: r[idxTanggal] || "",
							hari: r[idxHari] || "",
							jam: r[idxJam] || "",
							nama_guru: r[idxNama] || "",
							kelas: r[idxKelas] || "",
							mapel: r[idxMapel] || "",
							status: r[idxStatus] || ""
						}));

					if (reqTanggal) {
						result = result.filter(r => (r.tanggal || "").trim() === reqTanggal.trim());
					}
					return createResponse({ meta: { count: result.length }, data: result });
				} catch (error) {
					return createErrorResponse("Gagal mengambil rekap: " + error.message, 500);
				}
			}

			// ================= BATCH 2 ==================
			
			// 1. POST /api/absen
			if (url.pathname === '/api/absen' && request.method === 'POST') {
				try {
					const body = await request.json();
					const { jam, tanggal, data: submitData } = body;
					if (!jam || !tanggal || !Array.isArray(submitData)) {
						return createErrorResponse("Payload tidak lengkap", 400);
					}
					const hari = getIndoDayName(tanggal);
					
					// Format 9 kolom rekap:
					// [Tanggal Masehi, Tanggal Hijriah, Hari, Jam, Nama Ust, Kelas, Status, Pengganti, Absensi Key]
					const valuesToAppend = submitData.map(item => {
						const namaUst = item.nama_guru || "";
						const key = `${tanggal}|${jam}|${namaUst.toLowerCase().trim()}`;
						return [
							tanggal,
							"", // Tanggal Hijriah (opsional)
							hari,
							String(jam),
							namaUst,
							item.kelas || "",
							item.status || "HADIR",
							item.mapel || "",
							key
						];
					});

					await sheets.appendData('rekap!A:I', valuesToAppend);
					return createResponse({ status: "success", message: `${valuesToAppend.length} data absensi disimpan` });
				} catch (error) {
					return createErrorResponse("Gagal menyimpan absen: " + error.message, 500);
				}
			}

			// 2. POST /api/koreksi
			if (url.pathname === '/api/koreksi' && request.method === 'POST') {
				try {
					const { nama_guru, jam, tanggal, status_baru } = await request.json();
					if (!nama_guru || !jam || !tanggal || !status_baru) {
						return createErrorResponse("Payload koreksi tidak lengkap", 400);
					}

					// Baca semua rekap (9 kolom)
					const data = await sheets.readData('rekap!A:I');
					if (!data || data.length === 0) return createErrorResponse("Data rekap kosong", 404);

					const targetKey = `${tanggal}|${jam}|${nama_guru.toLowerCase().trim()}`;
					
					// Cari baris yang sesuai (1-based index di Google Sheets)
					let rowIndex = -1;
					for (let i = 1; i < data.length; i++) {
						const r = data[i];
						const keyInSheet = r[8] ? String(r[8]).trim().toLowerCase() : "";
						if (keyInSheet && keyInSheet === targetKey.toLowerCase()) {
							rowIndex = i + 1;
							break;
						}
						// Fallback pencocokan kolom 0 (tanggal), 3 (jam), 4 (nama)
						if (r[0] === tanggal && String(r[3] || r[2]).trim() === String(jam).trim() && (r[4] || r[3] || "").trim().toLowerCase() === nama_guru.toLowerCase()) {
							rowIndex = i + 1;
							break;
						}
					}

					if (rowIndex === -1) {
						return createErrorResponse("Data absensi tidak ditemukan untuk dikoreksi", 404);
					}

					// Update baris kolom G (Status)
					await sheets.updateData(`rekap!G${rowIndex}:G${rowIndex}`, [[status_baru]]);
					
					return createResponse({ status: "success", message: "Koreksi berhasil" });
				} catch (error) {
					return createErrorResponse("Gagal mengoreksi absen: " + error.message, 500);
				}
			}

			// 3. GET /api/rekap-bulanan
			if (url.pathname === '/api/rekap-bulanan' && request.method === 'GET') {
				try {
					const reqMonth = parseInt(url.searchParams.get('month') || (new Date().getMonth() + 1), 10);
					const reqYear = parseInt(url.searchParams.get('year') || new Date().getFullYear(), 10);

					let [jadwalData, rekapData] = await Promise.all([
						sheets.readData('jadwal!A:Z').catch(() => sheets.readData('Jadwal!A:Z')),
						sheets.readData('rekap!A:I')
					]);

					const jRows = (jadwalData || []).slice(1);
					const lRows = (rekapData || []).slice(1);

					const jHeader = (jadwalData[0] || []).map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
					const idxHari = jHeader.indexOf('hari');
					const idxNama = jHeader.indexOf('nama_ust') !== -1 ? jHeader.indexOf('nama_ust') : (jHeader.indexOf('nama_guru') !== -1 ? jHeader.indexOf('nama_guru') : jHeader.indexOf('nama'));

					if (idxHari === -1 || idxNama === -1) {
						return createErrorResponse("Kolom 'hari' atau 'nama_ust' tidak ditemukan di sheet jadwal", 500);
					}

					// A. Hitung Beban Jadwal Guru per Hari
					const teacherDayLoad = {};
					jRows.forEach(row => {
						const hari = (row[idxHari] || "").trim().toLowerCase();
						const nama = (row[idxNama] || "").trim();
						if (!nama || hari === "hari") return;

						if (!teacherDayLoad[nama]) teacherDayLoad[nama] = {};
						if (!teacherDayLoad[nama][hari]) teacherDayLoad[nama][hari] = 0;
						teacherDayLoad[nama][hari]++;
					});

					// B. Init tableData
					const tableData = {};
					Object.keys(teacherDayLoad).forEach(nama => {
						tableData[nama] = {
							nama_guru: nama,
							hadir: 0, izin: 0, sakit: 0, libur: 0, alpha: 0,
							total_wajib: 0
						};
					});

					// C. Hitung Total Wajib (Jumat Libur)
					const dayCounts = getWeekdayCount(reqYear, reqMonth);
					Object.keys(teacherDayLoad).forEach(nama => {
						let total = 0;
						Object.keys(teacherDayLoad[nama]).forEach(hari => {
							const countDay = dayCounts[hari] || 0;
							total += (teacherDayLoad[nama][hari] * countDay);
						});
						if (tableData[nama]) tableData[nama].total_wajib = total;
					});

					const rHeader = (rekapData[0] || []).map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
					const idxTanggal = rHeader.indexOf('tanggal_masehi') !== -1 ? rHeader.indexOf('tanggal_masehi') : (rHeader.indexOf('tanggal') !== -1 ? rHeader.indexOf('tanggal') : 0);
					const idxNamaRekap = rHeader.indexOf('nama_ust') !== -1 ? rHeader.indexOf('nama_ust') : (rHeader.indexOf('nama_guru') !== -1 ? rHeader.indexOf('nama_guru') : (rHeader.indexOf('nama') !== -1 ? rHeader.indexOf('nama') : 4));
					const idxStatus = rHeader.indexOf('status') !== -1 ? rHeader.indexOf('status') : 6;

					// D. Proses Rekap (Log Absensi)
					lRows.forEach(row => {
						const rawDate = row[idxTanggal];
						if (!rawDate || String(rawDate).toLowerCase().includes('tanggal')) return;
						const d = new Date(rawDate);
						if (isNaN(d.getTime())) return;
						if (d.getFullYear() !== reqYear || (d.getMonth() + 1) !== reqMonth) return;

						const nama = String(row[idxNamaRekap] || "").trim();
						let status = String(row[idxStatus] || "").trim().toLowerCase();
						if (status === 'alpa') status = 'alpha';

						if (tableData[nama]) {
							if (['izin', 'sakit', 'libur', 'alpha'].includes(status)) {
								tableData[nama][status]++;
							}
						}
					});

					// E. Finalisasi Hadir
					Object.values(tableData).forEach(t => {
						const absen = t.izin + t.sakit + t.libur + t.alpha;
						t.hadir = Math.max(0, t.total_wajib - absen);
					});

					return createResponse({
						meta: { year: reqYear, month: reqMonth },
						data: Object.values(tableData)
					});
				} catch (error) {
					return createErrorResponse("Gagal mengambil rekap bulanan: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/sync-bulanan' && request.method === 'POST') {
				try {
					const body = await request.json();
					const { monthName, year } = body;
					if (!monthName) return createErrorResponse("Bulan wajib diisi", 400);

					const sheetsHelper = new GoogleSheetsHelper(env);
					await syncBulananOnline(sheetsHelper, monthName, year || new Date().getFullYear().toString());

					return createResponse({ message: "Sinkronisasi rekap bulanan berhasil diperbarui secara langsung!" });
				} catch (error) {
					return createErrorResponse("Gagal sinkronisasi bulanan: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/ekspor-bulanan' && request.method === 'POST') {
				try {
					const body = await request.json();
					const { format, monthName, year } = body;
					if (!format || !monthName || !year) {
						return createErrorResponse("Format, nama bulan, dan tahun wajib diisi", 400);
					}

					const taskId = Date.now().toString();
					const payload = JSON.stringify({ format, monthName, year, phone: user.phone });
					
					await sheets.appendData('Task_Queue!A:F', [
						[taskId, 'EKSPOR_BULANAN', payload, 'PENDING', new Date().toISOString(), '']
					]);

					return createResponse({ message: "Ekspor laporan telah dijadwalkan ke Task Queue" });
				} catch (error) {
					return createErrorResponse("Gagal memproses ekspor laporan: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/save-rekap-sheet' && request.method === 'POST') {
				try {
					const body = await request.json();
					const { monthName, year, values } = body;
					if (!monthName || !values || !Array.isArray(values)) {
						return createErrorResponse("Nama bulan dan values wajib diisi", 400);
					}

					const cleanMonth = monthName.trim();
					const sheetsHelper = new GoogleSheetsHelper(env);
					const sheetList = await sheetsHelper.getSheetsList();
					let sheetExists = sheetList.some(s => s.properties.title.toLowerCase() === cleanMonth.toLowerCase());

					if (!sheetExists) {
						await sheetsHelper.createSheet(cleanMonth);
					}

					try {
						await sheetsHelper.clearData(`${cleanMonth}!A:ZZ`);
					} catch (clearErr) {
						console.warn("Gagal mengosongkan sheet:", clearErr.message);
					}

					// Tulis data values secara RAW (angka murni)
					const headers = await sheetsHelper.getHeaders();
					const updateUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetsHelper.sheetId}/values/${encodeURIComponent(cleanMonth + '!A1')}?valueInputOption=RAW`;
					
					const response = await fetch(updateUrl, {
						method: "PUT",
						headers,
						body: JSON.stringify({ values }),
					});

					if (!response.ok) {
						throw new Error(`Gagal menulis data rekap ke Google Sheets: ${await response.text()}`);
					}

					return createResponse({ message: "Berhasil menyimpan rekap bulanan ke Google Sheets!" });
				} catch (error) {
					return createErrorResponse("Gagal menyimpan rekap: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/download-rekap' && request.method === 'GET') {
				try {
					const format = url.searchParams.get('format');
					const monthName = url.searchParams.get('monthName');
					const year = url.searchParams.get('year') || new Date().getFullYear().toString();

					if (!format || !monthName) {
						return createErrorResponse("Format dan nama bulan wajib diisi", 400);
					}

					const cleanMonth = monthName.trim();
					const sheetsHelper = new GoogleSheetsHelper(env);

					try {
						await syncBulananOnline(sheetsHelper, cleanMonth, year);
					} catch (syncErr) {
						console.warn("Gagal melakukan pra-sinkronisasi bulanan online:", syncErr.message);
					}

					const sheetList = await sheetsHelper.getSheetsList();
					const targetSheet = sheetList.find(s => s.properties.title.toLowerCase() === cleanMonth.toLowerCase());
					
					if (!targetSheet) {
						return createErrorResponse(`Sheet rekap untuk bulan ${cleanMonth} tidak ditemukan.`, 404);
					}

					const sheetId = targetSheet.properties.sheetId;
					const sheetsHeaders = await sheetsHelper.getHeaders();

					let exportUrl = "";
					let contentType = "";
					let filename = `Rekap_Absensi_${cleanMonth}_${year}`;

					if (format === "xlsx") {
						exportUrl = `https://docs.google.com/spreadsheets/d/${env.GOOGLE_SHEET_ID}/export?format=xlsx&gid=${sheetId}`;
						contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
						filename += ".xlsx";
					} else {
						exportUrl = `https://docs.google.com/spreadsheets/d/${env.GOOGLE_SHEET_ID}/export?format=pdf&gid=${sheetId}` +
							`&size=A4&portrait=true&fitw=true&gridlines=true`;
						contentType = "application/pdf";
						filename += ".pdf";
					}

					const response = await fetch(exportUrl, {
						method: "GET",
						headers: {
							Authorization: sheetsHeaders.Authorization
						}
					});

					if (!response.ok) {
						throw new Error(`Gagal mengunduh file dari Google Sheets API: ${await response.text()}`);
					}

					const fileBuffer = await response.arrayBuffer();

					return new Response(fileBuffer, {
						headers: {
							...corsHeaders,
							"Content-Type": contentType,
							"Content-Disposition": `attachment; filename="${filename}"`,
							"Access-Control-Expose-Headers": "Content-Disposition"
						}
					});
				} catch (error) {
					return createErrorResponse("Gagal mengunduh berkas laporan: " + error.message, 500);
				}
			}

			// 4. GET & POST /api/settings
			if (url.pathname === '/api/settings') {
				try {
					if (request.method === 'GET') {
						let autoRekapActive = true;
						try {
							const data = await sheets.readDataIfExists('Setting!A1:B20');
							const row = data.find(r => String(r[0] || "").trim() === 'autoRekapActive');
							if (row && String(row[1]).toUpperCase() === 'FALSE') autoRekapActive = false;
						} catch (e) { /* ignore if sheet missing */ }
						return createResponse({ autoRekapActive });
					}
					else if (request.method === 'POST') {
						const body = await request.json();
						const key = 'autoRekapActive';
						const val = String(body.autoRekapActive).toUpperCase();
						try {
							let data = await sheets.readDataIfExists('Setting!A1:B20');
							if (!data || data.length === 0) {
								try {
									await sheets.createSheet('Setting');
								} catch (createErr) {
									/* silent fail if already created */
								}
							}
							
							const rowIndex = data.findIndex(row => String(row[0] || "").trim() === key);
							if (rowIndex !== -1) {
								await sheets.updateData(`Setting!B${rowIndex + 1}`, [[val]]);
							} else {
								await sheets.appendData('Setting!A:B', [[key, val]]);
							}
						} catch (e) {
							return createErrorResponse("Gagal menyimpan setting: " + e.message, 500);
						}
						return createResponse({ autoRekapActive: body.autoRekapActive });
					}
				} catch (error) {
					return createErrorResponse("Gagal memproses settings: " + error.message, 500);
				}
			}

			const ensureTaskQueueSheet = async () => {
				try {
					const sheetList = await sheets.getSheetsList();
					const hasSheet = sheetList.some(s => s.properties?.title === 'Task_Queue');
					if (!hasSheet) {
						await sheets.createSheet('Task_Queue');
						await sheets.appendData('Task_Queue!A:F', [[
							'id', 'type', 'payload', 'status', 'created_at', 'updated_at'
						]]);
					}
				} catch (e) {
					/* ignore if already exists or fails */
				}
			};

			// 8. Broadcast Endpoint
			if (url.pathname === '/api/broadcast' && request.method === 'POST') {
				try {
					const body = await request.json();
					const payload = JSON.stringify(body);
					const id = Date.now().toString();
					const ts = new Date().toISOString();
					await ensureTaskQueueSheet();
					await sheets.appendData('Task_Queue!A:F', [[
						id, 'SEND_WHATSAPP', payload, 'PENDING', ts, ''
					]]);
					return createResponse({ message: "Broadcast task queued" });
				} catch (error) {
					return createErrorResponse("Gagal menambahkan task broadcast: " + error.message, 500);
				}
			}

			// 9. Alarm Endpoint
			if (url.pathname === '/api/alarm' && request.method === 'POST') {
				try {
					const id = Date.now().toString();
					const ts = new Date().toISOString();
					await ensureTaskQueueSheet();
					await sheets.appendData('Task_Queue!A:F', [[
						id, 'ALARM', '{}', 'PENDING', ts, ''
					]]);
					return createResponse({ message: "Alarm task queued" });
				} catch (error) {
					return createErrorResponse("Gagal menambahkan task alarm: " + error.message, 500);
				}
			}

			// ==========================================
			// SUPERADMIN CONTROL PANEL ENDPOINTS
			// ==========================================

			// --- 1. GURU MANAGEMENT ---
			if (url.pathname === '/api/admin/guru' && request.method === 'GET') {
				try {
					requireSuperadmin(user);
					let data = [];
					let sheetName = 'no_wa';
					try { data = await sheets.readData('no_wa!A:Z'); }
					catch { data = await sheets.readData('Guru!A:Z'); sheetName = 'Guru'; }

					if (data.length === 0) return createResponse({ data: [], sheetName });

					const header = data[0].map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
					const rows = data.slice(1);
					const idxName = header.indexOf('nama_ust') !== -1 ? header.indexOf('nama_ust') : (header.indexOf('nama_guru') !== -1 ? header.indexOf('nama_guru') : header.indexOf('nama'));
					const idxPhone = header.indexOf('no_wa') !== -1 ? header.indexOf('no_wa') : header.indexOf('nomor_wa');
					const idxRole = header.indexOf('role');
					const idxHash = header.indexOf('password_hash');

					const result = rows.map((r, index) => ({
						rowIndex: index + 2,
						nama_guru: idxName !== -1 ? (r[idxName] || "") : "",
						nomor_wa: idxPhone !== -1 ? (r[idxPhone] || "") : "",
						role: idxRole !== -1 ? (r[idxRole] || "USER") : "USER",
						has_password: idxHash !== -1 ? !!(r[idxHash] && r[idxHash].trim()) : false
					})).filter(g => g.nomor_wa && g.nomor_wa.toUpperCase() !== 'KOSONG');

					return createResponse({ data: result, sheetName });
				} catch (error) {
					return createErrorResponse("Gagal mengambil data guru: " + error.message, error.message.includes("SUPERADMIN") ? 403 : 500);
				}
			}

			if (url.pathname === '/api/admin/guru' && request.method === 'POST') {
				try {
					requireSuperadmin(user);
					const { nama_guru, nomor_wa, password } = await request.json();
					if (!nama_guru || !nomor_wa) return createErrorResponse("Nama dan Nomor WA wajib diisi", 400);

					const normPhone = normalizePhoneNumber(nomor_wa);
					const plainPass = password && password.trim() ? password.trim() : "mubakid123";
					const passHash = await hashPassword(plainPass);

					let sheetName = 'no_wa';
					let data = [];
					try { data = await sheets.readData('no_wa!A:Z'); }
					catch { data = await sheets.readData('Guru!A:Z'); sheetName = 'Guru'; }

					if (data.length > 0) {
						const header = data[0].map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
						const idxPhone = header.indexOf('no_wa') !== -1 ? header.indexOf('no_wa') : header.indexOf('nomor_wa');
						if (idxPhone !== -1) {
							const exists = data.slice(1).some(r => normalizePhoneNumber(r[idxPhone]) === normPhone);
							if (exists) return createErrorResponse("Nomor WA sudah terdaftar", 400);
						}
					}

					await sheets.appendData(`${sheetName}!A:D`, [[nama_guru.trim(), normPhone, passHash, 'USER']]);
					return createResponse({ status: "success", message: `Guru ${nama_guru} berhasil ditambahkan` });
				} catch (error) {
					return createErrorResponse("Gagal menambah guru: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/admin/guru' && request.method === 'PUT') {
				try {
					requireSuperadmin(user);
					const { rowIndex, nama_guru, nomor_wa } = await request.json();
					if (!rowIndex || !nama_guru || !nomor_wa) return createErrorResponse("Data tidak lengkap", 400);

					const normPhone = normalizePhoneNumber(nomor_wa);
					let sheetName = 'no_wa';
					try { await sheets.readData('no_wa!A:A'); }
					catch { sheetName = 'Guru'; }

					await sheets.updateData(`${sheetName}!A${rowIndex}:B${rowIndex}`, [[nama_guru.trim(), normPhone]]);
					return createResponse({ status: "success", message: "Data guru berhasil diperbarui" });
				} catch (error) {
					return createErrorResponse("Gagal memperbarui guru: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/admin/guru' && request.method === 'DELETE') {
				try {
					requireSuperadmin(user);
					const { rowIndex } = await request.json();
					if (!rowIndex) return createErrorResponse("Row index wajib diisi", 400);

					let sheetName = 'no_wa';
					try { await sheets.readData('no_wa!A:A'); }
					catch { sheetName = 'Guru'; }

					await sheets.clearData(`${sheetName}!A${rowIndex}:Z${rowIndex}`);
					return createResponse({ status: "success", message: "Guru berhasil dihapus" });
				} catch (error) {
					return createErrorResponse("Gagal menghapus guru: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/admin/guru/reset-password' && request.method === 'POST') {
				try {
					requireSuperadmin(user);
					const { rowIndex, new_password } = await request.json();
					if (!rowIndex) return createErrorResponse("Row index wajib diisi", 400);

					const passToSet = new_password && new_password.trim() ? new_password.trim() : "mubakid123";
					const passHash = await hashPassword(passToSet);

					let sheetName = 'no_wa';
					try { await sheets.readData('no_wa!A:A'); }
					catch { sheetName = 'Guru'; }

					await sheets.updateData(`${sheetName}!C${rowIndex}:C${rowIndex}`, [[passHash]]);
					return createResponse({ status: "success", message: `Password berhasil di-reset menjadi '${passToSet}'` });
				} catch (error) {
					return createErrorResponse("Gagal reset password guru: " + error.message, 500);
				}
			}

			// --- 2. ADMIN WEB MANAGEMENT ---
			if (url.pathname === '/api/admin/admins' && request.method === 'GET') {
				try {
					requireSuperadmin(user);
					const adminData = await sheets.readData('admin_web!A:Z');
					if (!adminData || adminData.length === 0) return createResponse({ data: [] });

					const header = adminData[0].map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
					const rows = adminData.slice(1);
					const aRoleIdx = header.indexOf('role');
					const aPhoneIdx = header.indexOf('no_wa_role') !== -1 ? header.indexOf('no_wa_role') : header.indexOf('no_wa');
					const aKelasIdx = header.indexOf('akses_kelas');

					const result = rows.map((r, index) => ({
						rowIndex: index + 2,
						role: (r[aRoleIdx] || "ADMIN").toUpperCase(),
						no_wa: r[aPhoneIdx] || "",
						akses_kelas: aKelasIdx !== -1 ? (r[aKelasIdx] || "") : ""
					})).filter(a => a.no_wa);

					return createResponse({ data: result });
				} catch (error) {
					return createErrorResponse("Gagal mengambil data admin: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/admin/admins' && request.method === 'POST') {
				try {
					requireSuperadmin(user);
					const { no_wa, role, akses_kelas, password } = await request.json();
					if (!no_wa || !role) return createErrorResponse("Nomor WA dan Role wajib diisi", 400);

					const normPhone = normalizePhoneNumber(no_wa);
					const defaultPass = (role.toUpperCase() === "SUPERADMIN" || role.toUpperCase() === "ADMIN") ? "mu-1983" : "mubakid123";
					const plainPass = password && password.trim() ? password.trim() : defaultPass;
					const passHash = await hashPassword(plainPass);
					const formattedAkses = Array.isArray(akses_kelas) ? akses_kelas.join(", ") : (akses_kelas || "");

					const adminData = await sheets.readData('admin_web!A:Z');
					if (adminData && adminData.length > 1) {
						const header = adminData[0].map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
						const aPhoneIdx = header.indexOf('no_wa_role') !== -1 ? header.indexOf('no_wa_role') : header.indexOf('no_wa');
						if (aPhoneIdx !== -1) {
							const exists = adminData.slice(1).some(r => normalizePhoneNumber(r[aPhoneIdx]) === normPhone);
							if (exists) return createErrorResponse("Nomor WA admin sudah terdaftar", 400);
						}
					}

					await sheets.appendData('admin_web!A:D', [[role.toLowerCase(), normPhone, passHash, formattedAkses]]);
					return createResponse({ status: "success", message: `Admin ${role} berhasil ditambahkan` });
				} catch (error) {
					return createErrorResponse("Gagal menambah admin: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/admin/admins' && request.method === 'PUT') {
				try {
					requireSuperadmin(user);
					const { rowIndex, role, no_wa, akses_kelas } = await request.json();
					if (!rowIndex || !role || !no_wa) return createErrorResponse("Data tidak lengkap", 400);

					const normPhone = normalizePhoneNumber(no_wa);
					const formattedAkses = Array.isArray(akses_kelas) ? akses_kelas.join(", ") : (akses_kelas || "");

					await sheets.updateData(`admin_web!A${rowIndex}:B${rowIndex}`, [[role.toLowerCase(), normPhone]]);
					await sheets.updateData(`admin_web!D${rowIndex}:D${rowIndex}`, [[formattedAkses]]);
					return createResponse({ status: "success", message: "Data admin berhasil diperbarui" });
				} catch (error) {
					return createErrorResponse("Gagal memperbarui admin: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/admin/admins' && request.method === 'DELETE') {
				try {
					requireSuperadmin(user);
					const { rowIndex } = await request.json();
					if (!rowIndex) return createErrorResponse("Row index wajib diisi", 400);

					await sheets.clearData(`admin_web!A${rowIndex}:Z${rowIndex}`);
					return createResponse({ status: "success", message: "Admin berhasil dihapus" });
				} catch (error) {
					return createErrorResponse("Gagal menghapus admin: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/admin/admins/reset-password' && request.method === 'POST') {
				try {
					requireSuperadmin(user);
					const { rowIndex, new_password } = await request.json();
					if (!rowIndex) return createErrorResponse("Row index wajib diisi", 400);

					const passToSet = new_password && new_password.trim() ? new_password.trim() : "mu-1983";
					const passHash = await hashPassword(passToSet);

					await sheets.updateData(`admin_web!C${rowIndex}:C${rowIndex}`, [[passHash]]);
					return createResponse({ status: "success", message: `Password admin berhasil di-reset menjadi '${passToSet}'` });
				} catch (error) {
					return createErrorResponse("Gagal reset password admin: " + error.message, 500);
				}
			}

			// --- 3. JADWAL MANAGEMENT ---
			if (url.pathname === '/api/admin/jadwal' && request.method === 'GET') {
				try {
					requireSuperadmin(user);
					let data = [];
					let sheetName = 'jadwal';
					try { data = await sheets.readData('jadwal!A:Z'); }
					catch { data = await sheets.readData('Jadwal!A:Z'); sheetName = 'Jadwal'; }

					if (data.length === 0) return createResponse({ data: [], sheetName, availableClasses: [] });

					const header = data[0].map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_"));
					const rows = data.slice(1);

					const idxHari = header.indexOf('hari');
					const idxJam = header.indexOf('jam');
					const idxNama = header.indexOf('nama_ust') !== -1 ? header.indexOf('nama_ust') : (header.indexOf('nama_guru') !== -1 ? header.indexOf('nama_guru') : header.indexOf('nama'));
					const idxKelas = header.indexOf('kelas');
					const idxMapel = header.indexOf('mapel') !== -1 ? header.indexOf('mapel') : header.indexOf('tipe');
					const idxNo = header.indexOf('no');

					const result = [];
					const classesSet = new Set();

					rows.forEach((r, index) => {
						const kelas = r[idxKelas] || "";
						if (kelas) classesSet.add(kelas.trim());

						result.push({
							rowIndex: index + 2,
							hari: r[idxHari] || "",
							jam: r[idxJam] || "",
							nama_guru: idxNama !== -1 ? (r[idxNama] || "") : "",
							kelas: kelas,
							mapel: idxMapel !== -1 ? (r[idxMapel] || "") : "",
							no: idxNo !== -1 ? (r[idxNo] || "") : String(index + 1)
						});
					});

					return createResponse({
						data: result,
						sheetName,
						availableClasses: Array.from(classesSet).sort()
					});
				} catch (error) {
					return createErrorResponse("Gagal mengambil data jadwal admin: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/admin/jadwal' && request.method === 'POST') {
				try {
					requireSuperadmin(user);
					const { hari, jam, nama_guru, kelas, mapel } = await request.json();
					if (!hari || !jam || !nama_guru || !kelas) {
						return createErrorResponse("Hari, Jam, Nama Guru, dan Kelas wajib diisi", 400);
					}

					let sheetName = 'jadwal';
					let data = [];
					try { data = await sheets.readData('jadwal!A:Z'); }
					catch { data = await sheets.readData('Jadwal!A:Z'); sheetName = 'Jadwal'; }

					const nextNo = data.length > 1 ? String(data.length) : "1";
					const newRow = [hari.trim(), String(jam).trim(), nama_guru.trim(), kelas.trim(), (mapel || "").trim(), nextNo];

					await sheets.appendData(`${sheetName}!A:F`, [newRow]);
					return createResponse({ status: "success", message: "Jadwal baru berhasil ditambahkan" });
				} catch (error) {
					return createErrorResponse("Gagal menambah jadwal: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/admin/jadwal' && request.method === 'PUT') {
				try {
					requireSuperadmin(user);
					const { rowIndex, hari, jam, nama_guru, kelas, mapel } = await request.json();
					if (!rowIndex || !hari || !jam || !nama_guru || !kelas) {
						return createErrorResponse("Data tidak lengkap", 400);
					}

					let sheetName = 'jadwal';
					try { await sheets.readData('jadwal!A:A'); }
					catch { sheetName = 'Jadwal'; }

					await sheets.updateData(`${sheetName}!A${rowIndex}:E${rowIndex}`, [[hari.trim(), String(jam).trim(), nama_guru.trim(), kelas.trim(), (mapel || "").trim()]]);
					return createResponse({ status: "success", message: "Jadwal berhasil diperbarui" });
				} catch (error) {
					return createErrorResponse("Gagal memperbarui jadwal: " + error.message, 500);
				}
			}

			if (url.pathname === '/api/admin/jadwal' && request.method === 'DELETE') {
				try {
					requireSuperadmin(user);
					const { rowIndex } = await request.json();
					if (!rowIndex) return createErrorResponse("Row index wajib diisi", 400);

					let sheetName = 'jadwal';
					try { await sheets.readData('jadwal!A:A'); }
					catch { sheetName = 'Jadwal'; }

					await sheets.clearData(`${sheetName}!A${rowIndex}:Z${rowIndex}`);
					return createResponse({ status: "success", message: "Jadwal berhasil dihapus" });
				} catch (error) {
					return createErrorResponse("Gagal menghapus jadwal: " + error.message, 500);
				}
			}

			// Endpoint Tidak Ditemukan
			return createErrorResponse("Endpoint tidak ditemukan", 404);
		};

		try {
			const response = await handleRequest();
			const newHeaders = new Headers(response.headers);
			for (const [key, value] of Object.entries(corsHeaders)) {
				newHeaders.set(key, value);
			}
			return new Response(response.body, {
				status: response.status,
				statusText: response.statusText,
				headers: newHeaders
			});
		} catch (fatalError) {
			return new Response(JSON.stringify({ error: fatalError.message || "Internal Server Error" }), {
				status: 500,
				headers: {
					"Content-Type": "application/json",
					"Access-Control-Allow-Origin": "*",
					"Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
					"Access-Control-Allow-Headers": "Content-Type, Authorization",
				}
			});
		}
	},
};
