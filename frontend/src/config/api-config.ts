/**
 * API Configuration
 * 
 * TO USE WITH NGROK (Recommended for Excel Add-in development):
 * 
 * 1. Install ngrok (if not already installed):
 *    - Mac: brew install ngrok
 *    - Windows: Download from https://ngrok.com/download
 * 
 * 2. Start your backend:
 *    cd backend
 *    uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
 * 
 * 3. In another terminal, start ngrok:
 *    ngrok http 8000
 * 
 * 4. Copy the HTTPS URL from ngrok output (e.g., https://abc123.ngrok.io)
 * 
 * 5. Replace the BASE_URL below with your ngrok URL:
 *    export const BASE_URL = 'https://abc123.ngrok.io';
 */

// CHANGE THIS TO YOUR NGROK URL:
export const BASE_URL = 'https://c9c26ea399f4.ngrok-free.app';

// Fallback to localhost if ngrok is down:
// export const BASE_URL = process.env.BACKEND_URL || 'http://127.0.0.1:8000';

export const API_ENDPOINTS = {
  health: '/health',
  test: '/api/excel/test',
  processData: '/api/excel/process-data',
  generateChart: '/api/excel/generate-chart',
  analyzeData: '/api/excel/analyze-data',
} as const;

export function getApiUrl(endpoint: keyof typeof API_ENDPOINTS): string {
  return `${BASE_URL}${API_ENDPOINTS[endpoint]}`;
}

export function getFullUrl(path: string): string {
  return `${BASE_URL}${path}`;
}