// Centralized API configuration.
// In development, VITE_API_URL defaults to http://localhost:3000.
// In production (Render), set VITE_API_URL to your deployed backend URL.
const BASE_URL = (import.meta.env.VITE_API_URL || "http://localhost:3000").replace(/\/+$/, "");

export const API = {
  dossierGenerate: `${BASE_URL}/dossier/generate`,
  merge: `${BASE_URL}/merge`,
  convert: `${BASE_URL}/convert`,
};
