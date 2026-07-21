import { API_ENDPOINTS } from "../config/api.js";
import { get, post } from "./client.js";

export const getRekapHarian = (tanggal) =>
  get(`${API_ENDPOINTS.REKAP}?tanggal=${tanggal}`);

export const getAllRekap = () =>
  get(`${API_ENDPOINTS.REKAP}?tanggal=`);

export const submitAbsen = (payload) => post(API_ENDPOINTS.ABSEN, payload);

export const koreksiAbsen = (payload) => post(API_ENDPOINTS.KOREKSI, payload);
