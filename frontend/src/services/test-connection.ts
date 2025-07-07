/**
 * Test connection to backend from Excel Add-in environment
 */

export async function testBackendConnection() {
  const baseUrls = [
    'http://127.0.0.1:8000',
    'http://localhost:8000'
  ];

  console.log('=== Testing Backend Connection ===');
  
  for (const baseUrl of baseUrls) {
    try {
      console.log(`\nTesting: ${baseUrl}/health`);
      
      // Shorter timeout for Excel Add-in environment
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 3000); // 3 second timeout
      
      const response = await fetch(`${baseUrl}/health`, {
        method: 'GET',
        mode: 'cors',
        headers: {
          'Content-Type': 'application/json',
        },
        signal: controller.signal
      });
      
      clearTimeout(timeoutId);
      console.log(`✅ ${baseUrl} - Status: ${response.status}, OK: ${response.ok}`);
      
      if (response.ok) {
        const data = await response.json();
        console.log(`Response data:`, data);
        return baseUrl; // Return working URL
      }
    } catch (error) {
      console.log(`❌ ${baseUrl} - Error:`, error instanceof Error ? error.message : error);
      console.log(`Error name:`, error instanceof Error ? error.name : 'Unknown');
    }
  }
  
  console.log('\n=== All connection attempts failed ===');
  return null;
}

// Test different fetch configurations
export async function testFetchMethods(baseUrl: string) {
  const configurations = [
    { name: 'Default', options: {} },
    { name: 'No CORS', options: { mode: 'no-cors' as RequestMode } },
    { name: 'Same-origin', options: { mode: 'same-origin' as RequestMode } },
    { name: 'Include credentials', options: { credentials: 'include' as RequestCredentials } },
  ];

  console.log(`\n=== Testing different fetch configurations for ${baseUrl} ===`);

  for (const config of configurations) {
    try {
      console.log(`\nTesting: ${config.name}`);
      const response = await fetch(`${baseUrl}/health`, {
        method: 'GET',
        headers: { 'Content-Type': 'application/json' },
        ...config.options
      });
      
      console.log(`✅ ${config.name} - Status: ${response.status}`);
      if (response.ok) {
        console.log(`Working configuration found: ${config.name}`);
        return config.options;
      }
    } catch (error) {
      console.log(`❌ ${config.name} - Error:`, error instanceof Error ? error.message : error);
    }
  }
  
  return null;
}