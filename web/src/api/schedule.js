import { API_ENDPOINTS } from "../config/api.js";
import { get } from "./client.js";

export const getJadwal = (hari) =>
  get(`${API_ENDPOINTS.JADWAL}?hari=${hari}`);
