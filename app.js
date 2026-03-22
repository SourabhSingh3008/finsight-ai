// FinSight AI - Main Logic

// DOM Elements
const views = {
    onboarding: document.getElementById('onboardingView'),
    input: document.getElementById('inputView'),
    loading: document.getElementById('loadingView'),
    dashboard: document.getElementById('dashboardView')
};

// Buttons & Inputs
const apiKeyInput = document.getElementById('apiKey');
const saveKeyBtn = document.getElementById('saveKeyBtn');
const keyError = document.getElementById('keyError');
const transactionInput = document.getElementById('transactionInput');
const excelUpload = document.getElementById('excelUpload');
const analyzeBtn = document.getElementById('analyzeBtn');
const loadDemoBtn = document.getElementById('loadDemoBtn');
const newAnalysisBtn = document.getElementById('newAnalysisBtn');
const resetBtn = document.getElementById('resetBtn');
const homeBtn = document.getElementById('homeBtn');

// Dashboard Elements
const dashboardElements = {
    totalSpend: document.getElementById('totalSpend'),
    subCount: document.getElementById('subCount'),
    topCategory: document.getElementById('topCategory'),
    aiInsightsList: document.getElementById('aiInsightsList'),
    categoryList: document.getElementById('categoryList'),
    subscriptionsList: document.getElementById('subscriptionsList'),
    transactionsList: document.getElementById('transactionsList'),
    subTotalBadge: document.getElementById('subTotalBadge')
};

// State
let GEMINI_API_KEY = localStorage.getItem('finsight_api_key') || '';
let analyzedData = null;

// Initialize
function init() {
    if (GEMINI_API_KEY) {
        apiKeyInput.value = GEMINI_API_KEY; // Pre-fill if exists
    }
    
    // Always start at onboarding so user can verify their key
    switchView('onboarding');

    // Event Listeners
    homeBtn.addEventListener('click', () => switchView('onboarding'));
    saveKeyBtn.addEventListener('click', handleSaveKey);
    excelUpload.addEventListener('change', handleExcelUpload);
    analyzeBtn.addEventListener('click', handleAnalyze);
    loadDemoBtn.addEventListener('click', handleLoadDemo);
    newAnalysisBtn.addEventListener('click', () => {
        transactionInput.value = '';
        switchView('input');
    });
    resetBtn.addEventListener('click', () => {
        localStorage.removeItem('finsight_api_key');
        GEMINI_API_KEY = '';
        apiKeyInput.value = '';
        switchView('onboarding');
    });
}

// UI Utilities
function switchView(viewName) {
    // Hide all
    Object.values(views).forEach(view => {
        view.classList.remove('section-active');
    });
    // Show target
    views[viewName].classList.add('section-active');
}

function handleSaveKey() {
    const key = apiKeyInput.value.trim();
    if (key.length < 20 && key !== 'demo') {
        keyError.textContent = 'Please enter a valid Gemini API Key.';
        keyError.classList.remove('hidden');
        return;
    }
    
    keyError.classList.add('hidden');
    GEMINI_API_KEY = key;
    localStorage.setItem('finsight_api_key', key);
    // Add checkmark animation briefly
    const btnSpan = saveKeyBtn.querySelector('span');
    btnSpan.textContent = 'Saved!';
    setTimeout(() => {
        btnSpan.textContent = 'Continue';
        switchView('input');
    }, 1000);
}

async function handleLoadDemo() {
    const fallbackText = `Oct 01    Netflix                 $15.99
Oct 02    Uber                    $24.50
Oct 04    Whole Foods             $112.30
Oct 05    Amazon Prime            $8.99
Oct 08    Spotify                 $10.99
Oct 15    Uber Eats               $35.50
Oct 18    Adobe Creative Cloud    $54.99`;
    try {
        const response = await fetch('sample_data.txt');
        if (response.ok) {
            const text = await response.text();
            transactionInput.value = text;
        } else {
            transactionInput.value = fallbackText;
        }
    } catch (e) {
        console.warn('Failed to load demo data via fetch, using fallback', e);
        transactionInput.value = fallbackText;
    }
}

async function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert to CSV
        const csv = XLSX.utils.sheet_to_csv(worksheet);
        transactionInput.value = csv;
        
        // Optional: clear file input so same file can trigger change event again if needed
        excelUpload.value = '';
    } catch (err) {
        console.error('Error reading excel:', err);
        alert('Failed to read the Excel file. Please ensure it is a valid .xlsx or .xls file.');
    }
}

// AI Analysis Logic
const SYSTEM_PROMPT = `You are FinSight AI, an expert financial analyst. 
You are given unstructured, messy bank transaction text. Your task is to process it and return a strict JSON response. 
DO NOT output any markdown blocks (like json codeblocks), just the raw JSON string.

The JSON MUST have the following structure:
{
  "total_spend": 123.45,
  "top_category": "Groceries",
  "insights": [
    "Insight 1 (Short paragraph giving personalized financial advice or noticing a trend)",
    "Insight 2 (Warning about high spending or subscription redundancy)",
    "Insight 3 (Actionable tip to save money based on the data)"
  ],
  "categories": [
    { "name": "Groceries", "amount": 112.30 },
    { "name": "Entertainment", "amount": 57.97 }
  ],
  "subscriptions": [
    { "name": "Netflix", "amount": 15.99, "frequency": "Monthly" }
  ],
  "extracted_transactions": [
    { "date": "Oct 01", "description": "Netflix", "amount": 15.99, "category": "Entertainment" }
  ]
}

Make sure to group similar items into categories. Identify recurring elements as subscriptions. Provide EXACTLY 3 insights.`;

async function handleAnalyze() {
    const rawData = transactionInput.value.trim();
    
    if (!rawData) {
        alert('Please paste some transactions to analyze.');
        return;
    }
    if (!GEMINI_API_KEY) {
        alert('API Key missing. Please reset session.');
        return;
    }

    switchView('loading');

    try {
        let result;
        if (GEMINI_API_KEY === 'demo') {
            await new Promise(resolve => setTimeout(resolve, 2000));
            result = {
                candidates: [{
                    content: {
                        parts: [{
                            text: JSON.stringify({
                                total_spend: 541.26,
                                top_category: "Groceries",
                                insights: [
                                    "Your grocery spending is quite high this month. Consider meal planning.",
                                    "You have multiple streaming subscriptions (Netflix, Prime, Spotify).",
                                    "Cutting down on eating out (Uber Eats) could save you over $50/month."
                                ],
                                categories: [
                                    { name: "Groceries", amount: 265.70 },
                                    { name: "Entertainment", amount: 67.97 },
                                    { name: "Transportation", amount: 84.70 },
                                    { name: "Dining", amount: 46.95 },
                                    { name: "Utilities", amount: 75.94 }
                                ],
                                subscriptions: [
                                    { name: "Netflix", amount: 15.99, frequency: "Monthly" },
                                    { name: "Amazon Prime", amount: 8.99, frequency: "Monthly" },
                                    { name: "Spotify", amount: 10.99, frequency: "Monthly" },
                                    { name: "Planet Fitness", amount: 29.99, frequency: "Monthly" }
                                ],
                                extracted_transactions: [
                                    { date: "Oct 01", description: "Netflix", amount: 15.99, category: "Entertainment" },
                                    { date: "Oct 04", description: "Whole Foods", amount: 112.30, category: "Groceries" }
                                ]
                            })
                        }]
                    }
                }]
            };
        } else {
            // Auto-discover permitted models for this specific API key
            let selectedModel = 'gemini-1.5-flash'; // safety default
            try {
                const listRes = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${GEMINI_API_KEY}`);
                const listData = await listRes.json();
                
                if (listData.models) {
                    const availableModels = listData.models
                        .filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent"))
                        .map(m => m.name.split('/')[1]); 

                    const preferred = ['gemini-1.5-flash', 'gemini-1.5-flash-latest', 'gemini-1.5-pro', 'gemini-pro'];
                    selectedModel = preferred.find(m => availableModels.includes(m)) || availableModels[0] || selectedModel;
                }
            } catch (e) {
                console.warn("Auto-discovery failed. Proceeding with default.", e);
            }

            // Perform generation with the confirmed model
            const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent?key=${GEMINI_API_KEY}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: SYSTEM_PROMPT + "\n\nAnalyze the following data:\n\n" + rawData }] }],
                    generationConfig: { temperature: 0.1 }
                })
            });
            
            result = await response.json();
            
            if (result.error) {
                throw new Error(`[Model: ${selectedModel}] ` + result.error.message);
            }
        }

        // Parse JSON from response
        const aiText = result.candidates[0].content.parts[0].text;
        
        // Clean out possible markdown code blocks
        let cleanText = aiText.trim();
        if (cleanText.startsWith('```json')) cleanText = cleanText.substring(7);
        if (cleanText.startsWith('```')) cleanText = cleanText.substring(3);
        if (cleanText.endsWith('```')) cleanText = cleanText.substring(0, cleanText.length - 3);
        
        analyzedData = JSON.parse(cleanText.trim());
        
        populateDashboard(analyzedData);
        switchView('dashboard');

    } catch (error) {
        console.error('Analysis failed', error);
        alert('Failed to analyze data: ' + error.message + '\nPlease check your API Key and formatting.');
        switchView('input');
    }
}

// Dashboard Hydration
function populateDashboard(data) {
    // Top Stats
    dashboardElements.totalSpend.textContent = `$${data.total_spend.toFixed(2)}`;
    dashboardElements.subCount.textContent = data.subscriptions.length;
    dashboardElements.topCategory.textContent = data.top_category || 'N/A';

    // Insights
    dashboardElements.aiInsightsList.innerHTML = '';
    data.insights.forEach((insight, i) => {
        const icons = ['fa-lightbulb', 'fa-triangle-exclamation', 'fa-piggy-bank'];
        const item = document.createElement('div');
        item.className = 'insight-item';
        item.style.animationDelay = `${i * 0.1}s`;
        item.innerHTML = `
            <h4><i class="fa-solid ${icons[i]} highlight"></i> Insight ${i+1}</h4>
            <p>${insight}</p>
        `;
        dashboardElements.aiInsightsList.appendChild(item);
    });

    // Categories
    dashboardElements.categoryList.innerHTML = '';
    const maxCat = Math.max(...data.categories.map(c => c.amount));
    data.categories.forEach(cat => {
        const percent = Math.min((cat.amount / maxCat) * 100, 100);
        const item = document.createElement('div');
        item.className = 'category-item';
        item.innerHTML = `
            <div class="cat-info">
                <span>${cat.name}</span>
                <strong>$${cat.amount.toFixed(2)}</strong>
            </div>
            <div class="cat-bar-bg">
                <div class="cat-bar-fill" style="width: 0%" data-target="${percent}%"></div>
            </div>
        `;
        dashboardElements.categoryList.appendChild(item);
    });
    
    // Animate category bars
    setTimeout(() => {
        const bars = document.querySelectorAll('.cat-bar-fill');
        bars.forEach(bar => {
            bar.style.width = bar.getAttribute('data-target');
        });
    }, 100);

    // Subscriptions
    dashboardElements.subscriptionsList.innerHTML = '';
    let subTotal = 0;
    if (data.subscriptions.length === 0) {
        dashboardElements.subscriptionsList.innerHTML = '<p class="text-muted">No subscriptions detected.</p>';
    } else {
        data.subscriptions.forEach(sub => {
            subTotal += sub.amount;
            const item = document.createElement('div');
            item.className = 'sub-item';
            
            // Generate initials or icon
            const initial = sub.name.substring(0, 2).toUpperCase();
            
            item.innerHTML = `
                <div class="sub-details">
                    <div class="sub-icon">${initial}</div>
                    <div>
                        <div class="sub-name">${sub.name}</div>
                        <div class="sub-freq">${sub.frequency}</div>
                    </div>
                </div>
                <div class="sub-amount">$${sub.amount.toFixed(2)}</div>
            `;
            dashboardElements.subscriptionsList.appendChild(item);
        });
    }
    dashboardElements.subTotalBadge.textContent = `$${subTotal.toFixed(2)} / mo`;

    // Transactions
    dashboardElements.transactionsList.innerHTML = '<div class="transaction-list-container"></div>';
    const txContainer = dashboardElements.transactionsList.querySelector('.transaction-list-container');
    
    data.extracted_transactions.forEach(tx => {
        const item = document.createElement('div');
        item.className = 'tx-item';
        item.innerHTML = `
            <div class="tx-left">
                <span class="tx-desc">${tx.description}</span>
                <span class="tx-meta"><span>${tx.date}</span> • <span>${tx.category}</span></span>
            </div>
            <div class="tx-amount">$${tx.amount.toFixed(2)}</div>
        `;
        txContainer.appendChild(item);
    });
}

// Start app
document.addEventListener('DOMContentLoaded', init);
