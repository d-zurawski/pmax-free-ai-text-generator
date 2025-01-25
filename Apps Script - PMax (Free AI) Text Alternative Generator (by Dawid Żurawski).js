/**
 * PMax (Free AI) Text Alternative Generator - Apps Script Project
 * Developed by Dawid Żurawski
 * LinkedIn: https://www.linkedin.com/in/dawid-%C5%BCurawski-61a1341b3/
 *
 * Description: This script generates text alternatives using various LLM models (Groq, Google AI Studio, and OpenRouter)
 * directly within Google Sheets. It handles API calls, rate limits, and provides a user-friendly interface for PPC specialists.
 *
 * Important: Modify the code cautiously. Incorrect changes may disrupt script functionality.
 *
 * The script includes detailed comments for better understanding.
 */

/**
 * Creates custom menu when spreadsheet opens.
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();

    // Retrieve API keys from script properties
    const apiKeys = {
        'API_GROQ': PropertiesService.getScriptProperties().getProperty('API_GROQ'),
        'API_GOOGLE': PropertiesService.getScriptProperties().getProperty('API_GOOGLE'),
        'API_OPENROUTER': PropertiesService.getScriptProperties().getProperty('API_OPENROUTER')
    };

    // Helper function to check if an API key is valid
    const isValidApiKey = (key, prefix) => key && key.startsWith(prefix);

    // Create the main menu
    const menu = ui.createMenu('PMax (Free AI) Text Alternative Generator (by Dawid Żurawski)');

    // Function to create model submenus
    function createModelSubMenu(providerName, models, rpm) {
        const modelMenu = ui.createMenu(`Select Model (${providerName}) ${rpm ? `(RPM: ${rpm})` : ''}`);
        for (const model in models) {
            modelMenu.addItem(`Generate with ${model}`, models[model]);
        }
        menu.addSubMenu(modelMenu);
    }

    // Groq Menu
     if (isValidApiKey(apiKeys.API_GROQ, 'gsk')) {
      const groqModels = {
        'gemma2-9b-it': 'generateWithGemma',
        'llama-3-70b-8192': 'generateWithLlama3_70b',
        'llama3-8b-8192': 'generateWithLlama3_8b',
        'llama-guard-3-8b': 'generateWithLlamaGuard',
        'mixtral-8x7b-32768': 'generateWithMixtral'
      };
      createModelSubMenu('Groq (Favourite Model)', groqModels, 60); // 60 to default value - można to zmienić
    }


    // Google AI Studio Menu
    if (isValidApiKey(apiKeys.API_GOOGLE, 'AI')) {
        const googleAiModels = {
            'gemini-1.5-flash': 'generateWithGemini15Flash',
            'gemini-1.5-flash-8b': 'generateWithGemini15Flash8B',
            'gemini-1.5-pro': 'generateWithGemini15Pro'
        };
        createModelSubMenu('Google AI Studio', googleAiModels, 15); // 15 to domyślne RPM Gemini
    }

    // OpenRouter Menu
    if (apiKeys.API_OPENROUTER && apiKeys.API_OPENROUTER.trim() !== '') {
        const openRouterModels = {
            'gemini-2.0-flash-exp:free': 'generateWithOpenRouter_Gemini2Flash',
            'google/gemini-2.0-flash-thinking-exp-1219:free': 'generateWithOpenRouter_Gemini2FlashThinking',
            'meta-llama/llama-3.1-405b-instruct:free': 'generateWithOpenRouter_llama_3_1_405b_instruct',
            'meta-llama/llama-3.2-1b-instruct:free': 'generateWithOpenRouter_llama_3_2_1b_instruct',
            'qwen/qwen-2-7b-instruct:free': 'generateWithOpenRouter_qwen_2_7b_instruct',
            'mistralai/mistral-7b-instruct:free': 'generateWithOpenRouter_mistral_7b_instruct',
            'microsoft/phi-3-mini-128k-instruct:free': 'generateWithOpenRouter_phi_3_mini_128k_instruct',
            'sophosympatheia/rogue-rose-103b-v0.2:free': 'generateWithOpenRouter_RogueRose103b_v02_free'
        };
        createModelSubMenu('OpenRouter', openRouterModels, 20); // 20 to domyślne RPM OpenRouter Free
    }

    // Add an item to stop the script to the main menu.
    menu.addItem('Stop Script', 'setScriptCancelled');
    // Add the complete menu to the UI.
    menu.addToUi();
}

/**
 * Sets the script stop flag to true.
 */
function setScriptCancelled() {
    PropertiesService.getScriptProperties().setProperty('SCRIPT_STATUS', 'STOP');
    Logger.log('Script was stopped by user.');
    SpreadsheetApp.getUi().alert('Script was stopped.');
}


// Functions to generate with different Groq models
function generateWithGemma() {
    promptForStartRow('gemma2-9b-it', 'groq');
}

function generateWithLlama3_70b() {
    promptForStartRow('llama-3-70b-8192', 'groq');
}

function generateWithLlama3_8b() {
    promptForStartRow('llama3-8b-8192', 'groq');
}

function generateWithLlamaGuard() {
    promptForStartRow('llama-guard-3-8b', 'groq');
}

function generateWithMixtral() {
    promptForStartRow('mixtral-8x7b-32768', 'groq');
}

// Functions to generate with different Google AI Studio models
function generateWithGemini15Flash() {
    promptForStartRow('gemini-1.5-flash', 'googleai');
}

function generateWithGemini15Flash8B() {
    promptForStartRow('gemini-1.5-flash-8b', 'googleai');
}

function generateWithGemini15Pro() {
    promptForStartRow('gemini-1.5-pro', 'googleai');
}

// Functions to generate with OpenRouter
function generateWithOpenRouter_Gemini2Flash() {
    promptForStartRow('google/gemini-2.0-flash-exp:free', 'openrouter');
}

function generateWithOpenRouter_Gemini2FlashThinking() {
    promptForStartRow('google/gemini-2.0-flash-thinking-exp-1219:free', 'openrouter');
}


function generateWithOpenRouter_llama_3_1_405b_instruct() {
    promptForStartRow('meta-llama/llama-3.1-405b-instruct:free', 'openrouter');
}
function generateWithOpenRouter_llama_3_2_1b_instruct() {
    promptForStartRow('meta-llama/llama-3.2-1b-instruct:free', 'openrouter');
}
function generateWithOpenRouter_qwen_2_7b_instruct() {
    promptForStartRow('qwen/qwen-2-7b-instruct:free', 'openrouter');
}
function generateWithOpenRouter_mistral_7b_instruct() {
    promptForStartRow('mistralai/mistral-7b-instruct:free', 'openrouter');
}
function generateWithOpenRouter_phi_3_mini_128k_instruct() {
    promptForStartRow('microsoft/phi-3-mini-128k-instruct:free', 'openrouter');
}

function generateWithOpenRouter_RogueRose103b_v02_free() {
    promptForStartRow('sophosympatheia/rogue-rose-103b-v0.2:free', 'openrouter');
}

/**
 * Prompts the user to choose if to start from scratch or from the last non-empty row.
 * @param {string} selectedModel - The selected LLM model.
 * @param {string} provider - The provider of the LLM model (e.g., 'groq', 'xai', 'googleai', 'openrouter').
 */
function promptForStartRow(selectedModel, provider) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
        'Start Generation',
        'Do you want to reset existing alternatives?',
        ui.ButtonSet.YES_NO_CANCEL
    );

    let startRow = 2;
    let clearAlternativesFlag = true;

    if (response === ui.Button.YES) {
        startRow = 2;
        clearAlternativesFlag = true;
    } else if (response === ui.Button.NO) {
        startRow = getFirstEmptyRowInAlternatives(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PMax Data & Alternatives'));
        clearAlternativesFlag = false;
    } else if (response === ui.Button.CANCEL) {
        return;
    }

    resetAndSetupGeneration(selectedModel, provider, startRow, clearAlternativesFlag);
}

/**
 * Resets script properties and prepares alternative generation.
 * @param {string} selectedModel - Selected model for alternative generation.
 * @param {string} provider - The provider of the selected model.
 * @param {Number} startRow - Starting row.
 * @param {boolean} clearAlternativesFlag - Flag indicating if the alternatives should be cleared.
 */
function resetAndSetupGeneration(selectedModel, provider, startRow = null, clearAlternativesFlag = true) {
    PropertiesService.getScriptProperties().setProperty('SCRIPT_STATUS', 'RUNNING');
    PropertiesService.getScriptProperties().setProperty('SELECTED_MODEL', selectedModel);
    PropertiesService.getScriptProperties().setProperty('PROVIDER', provider);

    if (clearAlternativesFlag) {
        clearAlternatives();
    }

    continueGeneration(startRow);
}

/**
 * Clears all alternative columns while keeping headers intact (including the second row).
 */
function clearAlternatives() {
    const SHEET_NAME = 'PMax Data & Alternatives';
    const ALTERNATIVE_COL_START = 14;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
        Logger.log(`Sheet "${SHEET_NAME}" not found.`);
        return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const alternativeColumns = lastCol >= ALTERNATIVE_COL_START ? lastCol - ALTERNATIVE_COL_START + 1 : 0;

    if (alternativeColumns > 0 && lastRow > 2) {
        sheet.getRange(3, ALTERNATIVE_COL_START, lastRow - 2, alternativeColumns).clearContent();
        Logger.log(`Cleared alternatives from row 3 downwards in columns ${ALTERNATIVE_COL_START} to ${ALTERNATIVE_COL_START + alternativeColumns - 1}.`);
    } else {
        Logger.log('No alternative columns to clear or no rows below the headers.');
    }
}

/**
 * Continues generating alternatives from the first empty row in the alternative columns
 * to the last row with resource text.
 */
function continueGeneration(startRow = null) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PMax Data & Alternatives');
    const selModel = PropertiesService.getScriptProperties().getProperty('SELECTED_MODEL');
    const provider = PropertiesService.getScriptProperties().getProperty('PROVIDER');

    if (!selModel || !provider) {
        SpreadsheetApp.getUi().alert('Error: No LLM model or provider selected.');
        Logger.log('Error: No LLM model or provider selected.');
        return;
    }

    let start = startRow || getFirstEmptyRowInAlternatives(sheet);
    const data = sheet.getDataRange().getValues();
    let lastRowWithText = 0;
    for (let i = data.length - 1; i >= 0; i--) {
        if (data[i][12] !== null && data[i][12] !== "") {
            lastRowWithText = i + 1;
            break;
        }
    }

    if (start > lastRowWithText) {
        SpreadsheetApp.getUi().alert('No rows to process.');
        Logger.log('No rows to process.');
        return;
    }

    let rowsToProcess = lastRowWithText - start + 1;
    Logger.log(`Starting generation from row ${start} to ${lastRowWithText}. Number of rows to process: ${rowsToProcess}`);
    SpreadsheetApp.getUi().alert(`Starting generation from row ${start}. ${rowsToProcess} rows remaining.`);

    for (let index = start - 1; index < lastRowWithText; index++) {
        checkScriptStatus();
        const row = data[index];
        const performanceLabel = row[9];
        const assetType = row[11];
        const assetText = row[12];
        const rowIndex = index + 1;

        if (rowIndex < 3) continue;

        if (assetText && performanceLabel === 'LOW') {
            Logger.log(`Processing row ${rowIndex}: Performance=${performanceLabel}, Type=${assetType}, Text="${assetText}"`);
            generateAlternatives(selModel, rowIndex, provider);
        }
    }

    Logger.log('Processing complete.');
    PropertiesService.getScriptProperties().deleteProperty('SELECTED_MODEL');
    PropertiesService.getScriptProperties().deleteProperty('PROVIDER');
    SpreadsheetApp.getUi().alert('Generation complete!');
}

/**
 * Normalizes text by removing extra spaces, and trailing dots.
 */
function normalizeText(text) {
  if (!text) return "";

  return text.trim().replace(/\s+/g, ' ').replace(/\.$/, '');
}

/**
 * Checks if the alternative contains invalid characters.
 * @param {string} alternative - The text to check.
 * @returns {boolean} - True if the text is valid, false otherwise.
 */
function isValidAlternative(alternative) {
    if (!alternative) {
        return false;
    }
  const invalidCharsRegex = /[?!"]/g;
    return !invalidCharsRegex.test(alternative);
}


/**
 * Main function to generate alternatives with detailed logging using all models.
 */
function generateAlternatives(selectedModel, startRow = null, provider) {
    const PERFORMANCE_LABEL_COL = 10;
    const ASSET_TYPE_COL = 12;
    const ASSET_TEXT_COL = 13;
    const ALTERNATIVE_COL_START = 14;
    const SHEET_NAME = 'PMax Data & Alternatives';

    // Rate limits for different models
    const RATE_LIMIT_MAP = {
        'gemma2-9b-it': { rpm: 60, rps: 1 },
        'llama-3-70b-8192': { rpm: 30, rps: 1 },
        'llama3-8b-8192': { rpm: 60, rps: 1 },
        'llama-guard-3-8b': { rpm: 60, rps: 1 },
        'mixtral-8x7b-32768': { rpm: 60, rps: 1 },
        'gemini-1.5-flash': { rpm: 15, rps: 1 },
        'gemini-1.5-flash-8b': { rpm: 15, rps: 1 },
        'gemini-1.5-pro': { rpm: 15, rps: 1 },
        'google/gemini-2.0-flash-exp:free': { rpm: 60, rps: 1 },
         'google/gemini-2.0-flash-thinking-exp-1219:free': { rpm: 20, rps: 1 },
        'meta-llama/llama-3.1-405b-instruct:free': { rpm: 20, rps: 1 },
         'meta-llama/llama-3.2-1b-instruct:free': { rpm: 20, rps: 1 },
         'qwen/qwen-2-7b-instruct:free': { rpm: 20, rps: 1 },
         'mistralai/mistral-7b-instruct:free': { rpm: 20, rps: 1 },
         'microsoft/phi-3-mini-128k-instruct:free': { rpm: 20, rps: 1 },
        'sophosympatheia/rogue-rose-103b-v0.2:free': { rpm: 20, rps: 1 },
        'default': { rpm: 30, rps: 1 }
    };
    const TIME_WINDOW_MS = 60000;

    // Maximum lengths for different asset types
    const maxLengthMap = {
        'headline': 30,
        'description': 90,
        'long_headline': 60,
    };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
        Logger.log(`Sheet "${SHEET_NAME}" not found.`);
        return;
    }

  const generatedAlternatives = getGeneratedAlternatives(sheet);
  let allGeneratedAlternatives = PropertiesService.getScriptProperties().getProperty('ALL_GENERATED_ALTERNATIVES');
  allGeneratedAlternatives = allGeneratedAlternatives ? JSON.parse(allGeneratedAlternatives) : {};


    const data = sheet.getDataRange().getValues();
    Logger.log(`Data loaded: ${data.length} rows`);
    const start = startRow ? startRow - 1 : 2;

    Logger.log(`Starting ${provider} generation from row: ${start + 1}, data length: ${data.length}`);
    let errorCount = 0;

    for (let index = start; index < data.length; index++) {
        checkScriptStatus();
        const rowIndex = index + 1;
        if (rowIndex < 3) continue;

        const row = data[index];
        const performanceLabel = row[PERFORMANCE_LABEL_COL - 1];
        const assetType = row[ASSET_TYPE_COL - 1];
        const assetText = row[ASSET_TEXT_COL - 1];

        if (performanceLabel === 'LOW' && assetType && assetText) {
            Logger.log(`Processing row ${rowIndex}: Performance=${performanceLabel}, Type=${assetType}, Text="${assetText}"`);
            let alternatives = generatedAlternatives[assetText.toLowerCase()] || new Set();


            for (let i = 0; i < 3; i++) {
                checkScriptStatus();
                const colIndex = ALTERNATIVE_COL_START + i;
                try {
                    Logger.log(`Generating ${provider} alternative for row: ${rowIndex}, column: ${colIndex}`);

                    const rateLimit = RATE_LIMIT_MAP[selectedModel] || RATE_LIMIT_MAP['default'];
                    let selectedApiKey = null;
                    if (provider === 'groq') {
                        selectedApiKey = PropertiesService.getScriptProperties().getProperty('API_GROQ');
                    }  else if (provider === 'googleai') {
                        selectedApiKey = PropertiesService.getScriptProperties().getProperty('API_GOOGLE');
                    } else if (provider === 'openrouter') {
                        selectedApiKey = PropertiesService.getScriptProperties().getProperty('API_OPENROUTER');
                    }

                    const alternative = generateAlternative(
                        provider,
                        selectedApiKey,
                        selectedModel,
                        assetType,
                        assetText,
                        maxLengthMap[assetType.toLowerCase()] || 90,
                        rateLimit,
                        TIME_WINDOW_MS,
                        rowIndex,
                        colIndex,
                        sheet,
                        alternatives,
                        allGeneratedAlternatives
                    );

                    if (alternative) {
                        const normalizedAlternative = normalizeText(alternative);
                         const cleanedAlternative = normalizedAlternative.replace(/[?!"]/g, '');
                        if (isValidAlternative(cleanedAlternative)) {
                          alternatives.add(normalizedAlternative);
                          generatedAlternatives[assetText.toLowerCase()] = alternatives;
                            if (!allGeneratedAlternatives[assetText.toLowerCase()]) {
                                allGeneratedAlternatives[assetText.toLowerCase()] = new Set();
                            }
                             allGeneratedAlternatives[assetText.toLowerCase()].add(normalizedAlternative);

                            const actualLength = cleanedAlternative.length;

                           if (actualLength <=  maxLengthMap[assetType.toLowerCase()] ) {
                                  sheet.getRange(rowIndex, colIndex).setValue(cleanedAlternative);
                                     Logger.log(`${provider} alternative generated successfully for row: ${rowIndex}, column: ${colIndex}. Alternative: ${cleanedAlternative}`);
                                        errorCount = 0;
                                }else {
                                  sheet.getRange(rowIndex, colIndex).setValue("");
                                     Logger.log(`REJECTED: "${cleanedAlternative}" exceeds max length (${actualLength} characters) for row ${rowIndex}, column: ${colIndex}`);
                                      errorCount++;
                                }

                            SpreadsheetApp.flush(); // Dodane flush()
                        } else {
                            sheet.getRange(rowIndex, colIndex).setValue("");
                            Logger.log(`Alternative "${alternative}" contains invalid character(s) (?!") and was not added to row: ${rowIndex}, column: ${colIndex}`);
                            errorCount++;
                            SpreadsheetApp.flush(); // Dodane flush()
                        }


                    } else {
                        sheet.getRange(rowIndex, colIndex).setValue("");
                        Logger.log(`${provider} alternative was empty for row: ${rowIndex}, column: ${colIndex}.`);
                        errorCount++;
                          SpreadsheetApp.flush(); // Dodane flush()
                    }


                    if (errorCount >= 3) {
                        const errorMessage = `Too many consecutive errors (${errorCount}) during API calls. Stopping script.`;
                        Logger.log(errorMessage);
                        SpreadsheetApp.getUi().alert(errorMessage);
                        PropertiesService.getScriptProperties().setProperty('SCRIPT_STATUS', 'STOP');
                        return;
                    }
                } catch (error) {
                    errorCount++;
                    Logger.log(`Error generating alternative for row ${rowIndex}, column ${colIndex}: ${error.message}`);
                    sheet.getRange(rowIndex, colIndex).setValue("");
                     SpreadsheetApp.flush(); // Dodane flush()
                    if (error.message.includes("Rate limit exceeded")) {
                        const errorMessage = `Rate limit exceeded. Stopping script. Error: ${error.message}`;
                        Logger.log(errorMessage);
                        SpreadsheetApp.getUi().alert(errorMessage);
                        PropertiesService.getScriptProperties().setProperty('SCRIPT_STATUS', 'STOP');
                        return;
                    }

                    if (errorCount >= 3) {
                        const errorMessage = `Too many consecutive errors (${errorCount}) during API calls. Stopping script. Error: ${error.message}`;
                        Logger.log(errorMessage);
                        SpreadsheetApp.getUi().alert(errorMessage);
                        PropertiesService.getScriptProperties().setProperty('SCRIPT_STATUS', 'STOP');
                        return;
                    }
                }
            }
        }
    }

    PropertiesService.getScriptProperties().setProperty('ALL_GENERATED_ALTERNATIVES', JSON.stringify(allGeneratedAlternatives));
    sheet.getRange(1, 1).setValue('Alternatives updated. Refresh to see changes.');
    SpreadsheetApp.flush();
    Logger.log("Processing complete. Alternatives saved to the spreadsheet.");
}

/**
 * Generates a single alternative with dynamic rate limit handling and detailed logging for any model.
 */
function generateAlternative(provider, apiKey, model, assetType, assetText, maxLength, rateLimit, timeWindowMs, rowIndex, colIndex, sheet, alternatives, allGeneratedAlternatives) {
    const MAX_RETRIES = 5;
    let retryCount = 0;
    const lastRequestTimes = [];
    let endpoint;
    let requestBody;
    let options;

    while (retryCount < MAX_RETRIES) {
        checkScriptStatus();

        const now = Date.now();
        lastRequestTimes.push(now);
        while (lastRequestTimes.length > 0 && now - lastRequestTimes[0] > timeWindowMs) {
            lastRequestTimes.shift();
        }

        if (lastRequestTimes.length > rateLimit.rpm) {
            const timeToWait = timeWindowMs - (now - lastRequestTimes[0]);
            Logger.log(`Rate limit exceeded (RPM). Waiting for ${timeToWait}ms...`);
            Utilities.sleep(timeToWait);
            continue;
        }

        const requestsInLastSecond = lastRequestTimes.filter(time => now - time < 1000).length;
        if (requestsInLastSecond > rateLimit.rps) {
            const timeToWait = 1000 - (now - lastRequestTimes[lastRequestTimes.length - requestsInLastSecond]);
            Logger.log(`Rate limit exceeded (RPS). Waiting for ${timeToWait}ms...`);
            Utilities.sleep(timeToWait);
            continue;
        }

        const prompt = generatePrompt(assetType, assetText, maxLength, alternatives);

        if (provider === 'groq') {
            endpoint = "https://api.groq.com/openai/v1/chat/completions";
            requestBody = {
                model: model,
                messages: [{ role: "user", content: prompt }],
                max_tokens: Math.min(100, maxLength + 20),
                temperature: 0.7,
            };
            options = {
                method: 'post',
                headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
                payload: JSON.stringify(requestBody),
                muteHttpExceptions: true,
            };
        }  else if (provider === 'googleai') {
            endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
            requestBody = { contents: [{ parts: [{ text: prompt }] }] };
            options = {
                method: 'post',
                headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
                payload: JSON.stringify(requestBody),
                muteHttpExceptions: true,
            };
        } else if (provider === 'openrouter') {
            endpoint = "https://openrouter.ai/api/v1/chat/completions";
            requestBody = {
                model: model,
                messages: [{ role: "user", content: prompt }],
            };
            options = {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${apiKey}`,
                    "HTTP-Referer": SpreadsheetApp.getActiveSpreadsheet().getUrl(),
                    "X-Title": SpreadsheetApp.getActiveSpreadsheet().getName(),
                    "Content-Type": "application/json"
                },
                payload: JSON.stringify(requestBody),
                muteHttpExceptions: true
            };
        }

        try {
            const response = UrlFetchApp.fetch(endpoint, options);
            const jsonResponse = JSON.parse(response.getContentText());
            Logger.log(`Full response from ${provider} with model ${model}: ${response.getContentText()}`);

            if (jsonResponse.error) {
                if (provider === 'groq' && jsonResponse.error.code === "rate_limit_exceeded") {
                    const retryAfterMatch = jsonResponse.error.message.match(/Please try again in (\d+(\.\d+)?)s/);
                    const delay = retryAfterMatch ? parseFloat(retryAfterMatch[1]) * 1000 : 2000;
                    Logger.log(`Rate limit exceeded. Retrying after ${delay}ms...`);
                    Utilities.sleep(delay);
                    retryCount++;
                    continue;
                }
                Logger.log(`Error from ${provider} API with model ${model}: ${jsonResponse.error.message}`);
                retryCount++;
                continue;
            }

            let alternative;
            if (provider === 'groq' ) {
                alternative = jsonResponse.choices[0]?.message?.content?.trim();
            } else if (provider === 'googleai') {
                alternative = jsonResponse.candidates[0]?.content?.parts[0]?.text?.trim();
            } else if (provider === 'openrouter') {
                alternative = jsonResponse.choices[0]?.message?.content?.trim();
            }

            if (!alternative) {
                throw new Error(`Empty or undefined content in response: ${response.getContentText()}`);
            }
             Logger.log(`Extracted alternative: "${alternative}"`);

            if (!alternative.startsWith("Alternative:")) {
                alternative = "Alternative: " + alternative;
            }

             alternative = alternative.replace(/^Alternative:\s*/i, '').trim().replace(/[\r\n]+/g," ").replace(/\.$/, '');

            return alternative;

        } catch (error) {
            retryCount++;
            Logger.log(`ERROR: Retry ${retryCount} for row ${rowIndex}, column ${colIndex}: ${error.message}`);
            Utilities.sleep(2000);
        }
    }
    return "";
}


/**
 * Generates prompt for LLM model.
 */
function generatePrompt(assetType, assetText, maxLength, alternatives) {
    const prompt = `Generate a distinct and creative ${assetType} for the following text: ${assetText}. Do not use these alternatives: ${Array.from(alternatives).join(', ')}.
  Requirements:
  - The text must be strictly under ${maxLength} characters.
  - Ensure the response is concise, precise, and impactful.
  - Use direct and simple language.
  - Do not include unnecessary punctuation or special characters.
  - Only the alternative text is allowed in the response.
  - Alternatives exceeding ${maxLength} characters are invalid and must not be returned.
  - The output must always follow this format:
    Alternative: <generated alternative here>.
  - Do not include any explanations, descriptions, or metadata.
  - Respond in the same language as the input text.
  - Aim for the maximum length allowed, make the generated text as long as possible without being verbose, while still being impactful, also make sure it is relevant to the input text.`;
    return prompt;
}

/**
 * Gets existing alternatives from the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet from which the alternatives are to be retrieved.
 * @returns {Object} - Object where the key is the resource text, and the value is a set of alternatives.
 */
function getGeneratedAlternatives(sheet) {
  const ALTERNATIVE_COL_START = 14;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const generatedAlternatives = {};

  for (let i = 0; i < data.length; i++) {
    const assetText = data[i][12]?.toLowerCase() || '';
    if (!generatedAlternatives[assetText]) {
      generatedAlternatives[assetText] = new Set();
    }
    for (let j = ALTERNATIVE_COL_START - 1; j < lastCol; j++) {
      const alternative = (data[i][j] || '').trim().toLowerCase();
      if (alternative) {
        generatedAlternatives[assetText].add(alternative);
      }
    }
  }
  Logger.log(`Generated alternatives loaded: ${Object.keys(generatedAlternatives).length} assets`);
  return generatedAlternatives;
}


/**
 * Gets the first empty row in the alternative columns, starting from a given row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to check.
 * @param {number} startRow The row to start checking from (1-indexed, defaults to 3 to skip headers).
 * @param {number} alternativeColStart The starting column number for alternatives (1-indexed, e.g., 14 for column N).
 * @returns {number} The first empty row number in the alternative columns, or last row + 1 if all are filled.
 */
function getFirstEmptyRowInAlternatives(sheet, startRow = 3, alternativeColStart = 14) {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (startRow > lastRow) {
        return startRow;
    }

    const numAlternativeCols = lastColumn >= alternativeColStart ? lastColumn - alternativeColStart + 1 : 0;

    if (numAlternativeCols <= 0) {
        return startRow;
    }

    for (let i = startRow; i <= lastRow; i++) {
        let isRowEmpty = true;
        for (let j = 0; j < numAlternativeCols; j++) {
            const cellValue = sheet.getRange(i, alternativeColStart + j).getValue();
            if (cellValue !== null && cellValue !== "") {
                isRowEmpty = false;
                break;
            }
        }
        if (isRowEmpty) {
            Logger.log(`First empty row in alternatives found at row: ${i}`);
            return i;
        }
    }
    Logger.log(`No empty row in alternatives found, returning next row after last: ${lastRow + 1}`);
    return lastRow + 1;
}

/**
 * Checks if the script should be stopped and throws an error if it is.
 */
function checkScriptStatus() {
    const scriptStatus = PropertiesService.getScriptProperties().getProperty('SCRIPT_STATUS');
    if (scriptStatus === 'STOP') {
        throw new Error('Script was stopped by user.');
    }
}
