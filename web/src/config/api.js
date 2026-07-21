export const BASE_URL = (import.meta.env.VITE_API_URL || import.meta.env.VITE_API_BASE_URL || "").replace(/\/$/, "");

export const API_ENDPOINTS = {
  STATUS: `${BASE_URL}/api/status`,
  SETTINGS: `${BASE_URL}/api/settings`,
  LOGIN: `${BASE_URL}/api/login`,
  JADWAL: `${BASE_URL}/api/jadwal`,
  REKAP: `${BASE_URL}/api/rekap`,
  KONTAK: `${BASE_URL}/api/kontak`,
  ABSEN: `${BASE_URL}/api/absen`,
  KOREKSI: `${BASE_URL}/api/koreksi`,
  REKAP_BULANAN: `${BASE_URL}/api/rekap-bulanan`,
  SYNC_BULANAN: `${BASE_URL}/api/sync-bulanan`,
  EKSPOR_BULANAN: `${BASE_URL}/api/ekspor-bulanan`,
  ALARM: `${BASE_URL}/api/alarm`,
  BROADCAST: `${BASE_URL}/api/broadcast`,
  ADMIN_GURU: `${BASE_URL}/api/admin/guru`,
  ADMIN_GURU_RESET: `${BASE_URL}/api/admin/guru/reset-password`,
  ADMIN_ADMINS: `${BASE_URL}/api/admin/admins`,
  ADMIN_ADMINS_RESET: `${BASE_URL}/api/admin/admins/reset-password`,
  ADMIN_JADWAL: `${BASE_URL}/api/admin/jadwal`,
};
