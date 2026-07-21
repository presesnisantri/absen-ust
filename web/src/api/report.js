import { API_ENDPOINTS } from "../config/api.js";
import { get, post } from "./client.js";

export const getRekapBulanan = (month, year) =>
  get(`${API_ENDPOINTS.REKAP_BULANAN}?month=${month}&year=${year}`);

export const syncBulanan = (payload) =>
  post(API_ENDPOINTS.SYNC_BULANAN, payload);

export const eksporBulanan = (payload) =>
  post(API_ENDPOINTS.EKSPOR_BULANAN, payload);

export const downloadRekapFile = async (format, monthName, year) => {
  const token = localStorage.getItem("token");
  const headers = {};
  if (token) {
    headers["Authorization"] = `Bearer ${token}`;
  }
  const url = `${API_ENDPOINTS.EKSPOR_BULANAN.replace("ekspor-bulanan", "download-rekap")}?format=${format}&monthName=${encodeURIComponent(monthName)}&year=${year}`;
  const response = await fetch(url, {
    method: "GET",
    headers
  });
  if (!response.ok) {
    let errText = "Gagal mengunduh berkas.";
    try {
      const errObj = await response.json();
      errText = errObj.error || errText;
    } catch {}
    throw new Error(errText);
  }
  return response.blob();
};

export const saveRekapSheet = (payload) =>
  post(API_ENDPOINTS.EKSPOR_BULANAN.replace("ekspor-bulanan", "save-rekap-sheet"), payload);
