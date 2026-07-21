import { API_ENDPOINTS } from "../config/api.js";
import { get, post } from "./client.js";

export const getStatus = () => get(API_ENDPOINTS.STATUS);

export const login = (phone, password) =>
  post(API_ENDPOINTS.LOGIN, { phone, password });
