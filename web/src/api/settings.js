import { API_ENDPOINTS } from "../config/api.js";
import { get, post } from "./client.js";

export const getSettings = () => get(API_ENDPOINTS.SETTINGS);

export const updateSettings = (payload) =>
  post(API_ENDPOINTS.SETTINGS, payload);

export const sendAlarm = () => post(API_ENDPOINTS.ALARM);

export const sendBroadcast = (payload) =>
  post(API_ENDPOINTS.BROADCAST, payload);

export const getKontak = () => get(API_ENDPOINTS.KONTAK);
