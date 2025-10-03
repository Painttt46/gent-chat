export const conversations = new Map();

// Model tracking with limits
export const models = {
    'gemini-2.5-flash': { name: 'Gemini 2.5 Flash', count: 0, limit: 500 },
    'gemini-2.5-pro': { name: 'Gemini 2.5 Pro', count: 0, limit: 100 }
};
export const userModels = new Map(); // Track current model per user
let lastResetDate = new Date().toDateString();
export let currentApiKeyIndex = 0; // 0 for primary, 1 for secondary

// Get current API key
export function getCurrentApiKey() {
    return currentApiKeyIndex === 0 ? process.env.GEMINI_API_KEY : process.env.GEMINI_API_KEY_2;
}

// Check if model hit limit and switch API key if needed
export function checkLimitsAndSwitchKey(modelKey) {
    const model = models[modelKey];
    if (model.count >= model.limit) {
        // Try to switch to other API key
        const newApiKeyIndex = currentApiKeyIndex === 0 ? 1 : 0;

        // Check if we've already tried both keys (both are maxed)
        const allModelsMaxed = Object.values(models).every(m => m.count >= m.limit);
        if (allModelsMaxed) {
            return 'MAXED_OUT';
        }

        // Switch to other API key and reset counters
        currentApiKeyIndex = newApiKeyIndex;
        Object.keys(models).forEach(key => models[key].count = 0);
        return true; // Switched
    }
    return false; // No switch
}

// Reset counters daily
export function checkDailyReset() {
    const today = new Date().toDateString();
    if (today !== lastResetDate) {
        Object.keys(models).forEach(key => models[key].count = 0);
        lastResetDate = today;
    }
}

// Get total requests across all models
export function getTotalRequests() {
    return Object.values(models).reduce((sum, model) => sum + model.count, 0);
}