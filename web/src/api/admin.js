import { API_ENDPOINTS } from "../config/api.js";
import { get, post, put, del } from "./client.js";


// Guru API
export const getGuru = () => get(API_ENDPOINTS.ADMIN_GURU);
export const addGuru = (data) => post(API_ENDPOINTS.ADMIN_GURU, data);
export const updateGuru = (data) => put(API_ENDPOINTS.ADMIN_GURU, data);
export const deleteGuru = (rowIndex) => del(API_ENDPOINTS.ADMIN_GURU, { rowIndex });
export const resetGuruPassword = (data) => post(API_ENDPOINTS.ADMIN_GURU_RESET, data);

// Admin Web API
export const getAdmins = () => get(API_ENDPOINTS.ADMIN_ADMINS);
export const addAdmin = (data) => post(API_ENDPOINTS.ADMIN_ADMINS, data);
export const updateAdmin = (data) => put(API_ENDPOINTS.ADMIN_ADMINS, data);
export const deleteAdmin = (rowIndex) => del(API_ENDPOINTS.ADMIN_ADMINS, { rowIndex });
export const resetAdminPassword = (data) => post(API_ENDPOINTS.ADMIN_ADMINS_RESET, data);

// Jadwal API
export const getJadwal = () => get(API_ENDPOINTS.ADMIN_JADWAL);
export const addJadwal = (data) => post(API_ENDPOINTS.ADMIN_JADWAL, data);
export const updateJadwal = (data) => put(API_ENDPOINTS.ADMIN_JADWAL, data);
export const deleteJadwal = (rowIndex) => del(API_ENDPOINTS.ADMIN_JADWAL, { rowIndex });
