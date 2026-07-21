import { SignJWT, importPKCS8 } from "jose";

/**
 * Mendapatkan Access Token dari Google API menggunakan Service Account
 */
async function getAccessToken(env) {
  const privateKey = env.GOOGLE_PRIVATE_KEY;
  const clientEmail = env.GOOGLE_CLIENT_EMAIL;

  if (!privateKey || !clientEmail) {
    throw new Error("Missing Google Service Account credentials in environment.");
  }

  // Parse private key
  const alg = "RS256";
  let formattedKey = String(privateKey || "").trim();
  if ((formattedKey.startsWith('"') && formattedKey.endsWith('"')) || (formattedKey.startsWith("'") && formattedKey.endsWith("'"))) {
    formattedKey = formattedKey.slice(1, -1);
  }
  formattedKey = formattedKey.replace(/\\n/g, "\n").replace(/\r/g, "").trim();
  
  const key = await importPKCS8(formattedKey, alg);

  const iat = Math.floor(Date.now() / 1000);
  const exp = iat + 3600;

  const jwt = await new SignJWT({
    iss: clientEmail,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: "https://oauth2.googleapis.com/token",
    exp,
    iat,
  })
    .setProtectedHeader({ alg, typ: "JWT" })
    .sign(key);

  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to get Google Access Token: ${errorText}`);
  }

  const data = await response.json();
  return data.access_token;
}

/**
 * Base class untuk Google Sheets Helper
 */
export class GoogleSheetsHelper {
  constructor(env) {
    this.env = env;
    this.sheetId = env.GOOGLE_SHEET_ID;
  }

  async getHeaders() {
    const token = await getAccessToken(this.env);
    return {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
  }

  /**
   * Membaca data dari sheet
   */
  async readData(range) {
    const headers = await this.getHeaders();
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${this.sheetId}/values/${encodeURIComponent(range)}`;
    
    const response = await fetch(url, { method: "GET", headers });
    
    if (!response.ok) {
      throw new Error(`Error reading from Google Sheets: ${await response.text()}`);
    }

    const data = await response.json();
    return data.values || [];
  }

  /**
   * Membaca data jika sheet ada, kembalikan array kosong jika sheet belum dibuat
   */
  async readDataIfExists(range) {
    try {
      const sheetName = range.split('!')[0];
      const sheetList = await this.getSheetsList();
      const hasSheet = sheetList.some(s => s.properties?.title?.toLowerCase() === sheetName.toLowerCase());
      if (!hasSheet) return [];
      return await this.readData(range);
    } catch {
      return [];
    }
  }

  /**
   * Menambahkan baris data (Append)
   */
  async appendData(range, values) {
    try {
      const headers = await this.getHeaders();
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${this.sheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED`;
      
      const response = await fetch(url, {
        method: "POST",
        headers,
        body: JSON.stringify({ values }),
      });
      
      if (!response.ok) {
        throw new Error(`Error appending to Google Sheets: ${await response.text()}`);
      }

      return await response.json();
    } catch (error) {
      console.error("[GoogleSheetsHelper] appendData Error:", error.message);
      throw error;
    }
  }

  /**
   * Memperbarui data pada range tertentu (Update)
   */
  async updateData(range, values) {
    try {
      const headers = await this.getHeaders();
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${this.sheetId}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`;
      
      const response = await fetch(url, {
        method: "PUT",
        headers,
        body: JSON.stringify({ values }),
      });
      
      if (!response.ok) {
        throw new Error(`Error updating Google Sheets: ${await response.text()}`);
      }

      return await response.json();
    } catch (error) {
      console.error("[GoogleSheetsHelper] updateData Error:", error.message);
      throw error;
    }
  }

  /**
   * Menghapus isi cell
   */
  async clearData(range) {
    try {
      const headers = await this.getHeaders();
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${this.sheetId}/values/${encodeURIComponent(range)}:clear`;
      
      const response = await fetch(url, {
        method: "POST",
        headers,
      });
      
      if (!response.ok) {
        throw new Error(`Error clearing Google Sheets: ${await response.text()}`);
      }

      return await response.json();
    } catch (error) {
      console.error("[GoogleSheetsHelper] clearData Error:", error.message);
      throw error;
    }
  }

  /**
   * Membuat sheet baru jika belum ada
   */
  async createSheet(title) {
    try {
      const headers = await this.getHeaders();
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${this.sheetId}:batchUpdate`;
      
      const response = await fetch(url, {
        method: "POST",
        headers,
        body: JSON.stringify({
          requests: [
            {
              addSheet: {
                properties: {
                  title,
                },
              },
            },
          ],
        }),
      });
      
      if (!response.ok) {
        throw new Error(`Error creating sheet: ${await response.text()}`);
      }

      return await response.json();
    } catch (error) {
      console.error("[GoogleSheetsHelper] createSheet Error:", error.message);
      throw error;
    }
  }

  /**
   * Mengambil daftar sheet dalam spreadsheet
   */
  async getSheetsList() {
    try {
      const headers = await this.getHeaders();
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${this.sheetId}?fields=sheets.properties`;
      
      const response = await fetch(url, { method: "GET", headers });
      if (!response.ok) {
        throw new Error(`Error fetching spreadsheet metadata: ${await response.text()}`);
      }

      const data = await response.json();
      return data.sheets || [];
    } catch (error) {
      console.error("[GoogleSheetsHelper] getSheetsList Error:", error.message);
      throw error;
    }
  }
}
