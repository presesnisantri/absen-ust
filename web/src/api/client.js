const TIMEOUT_MS = 15_000;

export const ERROR_TYPES = {
  NETWORK: "NETWORK",
  TIMEOUT: "TIMEOUT",
  UNAUTHORIZED: "UNAUTHORIZED",
  FORBIDDEN: "FORBIDDEN",
  NOT_FOUND: "NOT_FOUND",
  VALIDATION: "VALIDATION",
  SERVER: "SERVER",
  UNKNOWN: "UNKNOWN",
};

export class ApiError extends Error {
  constructor(message, status, type) {
    super(message);
    this.name = "ApiError";
    this.status = status ?? null;
    this.type = type ?? ERROR_TYPES.UNKNOWN;
  }
}

async function handleResponse(res) {
  const isLoginRequest = res.url && res.url.includes('/api/login');

  if (res.status === 401 && !isLoginRequest) {
    localStorage.removeItem("token");
    localStorage.removeItem("user");
    window.dispatchEvent(new CustomEvent("api:unauthorized"));
    throw new ApiError("Sesi login berakhir.", 401, ERROR_TYPES.UNAUTHORIZED);
  }

  if (res.status === 403) {
    throw new ApiError("Akses ditolak.", 403, ERROR_TYPES.FORBIDDEN);
  }

  if (res.status === 404) {
    throw new ApiError("Data tidak ditemukan.", 404, ERROR_TYPES.NOT_FOUND);
  }

  if (!res.ok) {
    let message =
      res.status >= 500
        ? "Terjadi kesalahan server."
        : "Terjadi kesalahan pada permintaan.";
    try {
      const body = await res.json();
      message = body.error || body.message || message;
    } catch {
      /* ignore */
    }
    const type =
      res.status >= 500 ? ERROR_TYPES.SERVER : ERROR_TYPES.VALIDATION;
    throw new ApiError(message, res.status, type);
  }

  return res.json();
}

async function request(url, options = {}) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), TIMEOUT_MS);

  const token = localStorage.getItem("token");
  const headers = { ...options.headers };

  if (options.body !== undefined) {
    headers["Content-Type"] = "application/json";
  }
  if (token) {
    headers["Authorization"] = `Bearer ${token}`;
  }

  try {
    const res = await fetch(url, {
      ...options,
      headers,
      signal: controller.signal,
    });
    clearTimeout(timeoutId);
    return await handleResponse(res);
  } catch (err) {
    clearTimeout(timeoutId);

    if (err instanceof ApiError) throw err;

    if (err.name === "AbortError") {
      throw new ApiError(
        "Permintaan habis waktu. Periksa koneksi dan coba lagi.",
        null,
        ERROR_TYPES.TIMEOUT
      );
    }

    throw new ApiError(
      "Gagal terhubung ke server.",
      null,
      ERROR_TYPES.NETWORK
    );
  }
}

export const get = (url) => request(url, { method: "GET" });

export const post = (url, body) => {
  const options = { method: "POST" };
  if (body !== undefined) {
    options.body = JSON.stringify(body);
  }
  return request(url, options);
};

export const put = (url, body) => {
  const options = { method: "PUT" };
  if (body !== undefined) {
    options.body = JSON.stringify(body);
  }
  return request(url, options);
};

export const del = (url, body) => {
  const options = { method: "DELETE" };
  if (body !== undefined) {
    options.body = JSON.stringify(body);
  }
  return request(url, options);
};

export { request };

