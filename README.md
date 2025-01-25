# PMax (Free AI) Text Alternative Generator (by Dawid Żurawski)

This project provides **a free solution** for generating alternative texts for asset groups in Google Ads Performance Max campaigns using AI. It leverages various language models, many of which offer free tiers or usage, to create optimized and creative ad text variations. The aim is to help you improve campaign performance without incurring significant costs for AI processing.

## Key Features

*   **Free AI-Powered Text Generation:** Utilizes free or freemium language models to generate alternative texts for headlines, long headlines, and descriptions.
*   **Google Ads Data Retrieval:** Automatically fetches asset group performance metrics directly from Google Ads (conversions, cost per conversion, clicks, CTR, impressions).
*   **Google Sheet Data Storage:** Stores the retrieved data in a clear, organized format within a Google Sheet for easy analysis.
*   **Diverse LLM Options:** Supports a variety of AI language models via Groq, Google AI Studio, and OpenRouter, with a focus on their free usage tiers.
*  **Rate Limit Management:** Handles API calls and rate limits to ensure smooth, uninterrupted operation.
*   **Targeted Text Generation:** Generates alternative texts only for low-performing assets ("LOW" performance labels), optimizing resource use.
*   **Comprehensive Logging:** Tracks all actions and errors, with a stop feature for user control.
*  **Data Sorting:** The data in the sheet is sorted by asset type (Headline, Long Headline, Description) and performance label (Low, Good, Best).
*  **Text Validation:** Ensures the generated text is within character limits and free from invalid characters.

## How to Use

1.  **Copy the Google Sheet:**
    *   Make a copy of [this template](https://docs.google.com/spreadsheets/d/1Lcef-sKrw0Bz02xDUMnl2Mhe10v9DcYbdqLg7PAyNzI/copy).
    *   Copy the URL of your new sheet.

2.  **Set Up the Scripts:**
   *   **The Google Apps Script should be copied automatically with the Google Sheet.**
      *  If, for some reason, the Apps Script did not copy over, or you are having trouble running it:
         *   Open the script editor in your Google Sheet (Extensions > Apps Script).
         *    Create a new script file and paste the content of `Apps Script - PMax (Free AI) Text Alternative Generator (by Dawid Żurawski).js` into it.
   *   **Set up the Google Ads Script:**
         *    Open your Google Ads account.
         *    Navigate to "Scripts" under "Bulk actions".
         *    Create a new script and paste the content of `Google Ads Script - PMax (Free AI) Text Alternative Generator (by Dawid Żurawski).js` into the editor. Modify the `SPREADSHEET_URL` variable (line 12) to point to your Google Sheet URL.

       In the **Google Apps Script Properties** (File > Script properties) add the following API keys (note that free usage tiers may have limits):
        *   `API_GROQ`: Your Groq API key (free tier available).
        *   `API_GOOGLE`: Your Google AI Studio API key (free tier available).
        *   `API_OPENROUTER`: Your OpenRouter API key (free tier available).

3.  **Run the Script:**
    *   A new menu item called "PMax (Free AI) Text Alternative Generator (by Dawid Żurawski)" will appear in the Google Sheet menu.
    *   Select your preferred AI model (Groq, Google AI Studio, or OpenRouter) from the menu.
    *   Choose whether to start the generation from scratch or from the last empty row with alternatives.
    *   The script will start generating alternatives automatically and update the Google Sheet.

4.  **Monitor the Results:**
    *   The script will populate the 'Alternative 1', 'Alternative 2', and 'Alternative 3' columns with the generated text options.
    *   The first cell in Column A will be updated with the last refresh time.

## Available AI Models

The script supports the following AI models, focusing on those with free tiers or free usage:

*   **Groq (Free Tier Available):**
    *   `gemma2-9b-it`
    *   `llama-3-70b-8192`
    *   `llama3-8b-8192`
    *   `llama-guard-3-8b`
    *   `mixtral-8x7b-32768`
*   **Google AI Studio (Free Tier Available):**
    *   `gemini-1.5-flash`
    *   `gemini-1.5-flash-8b`
    *   `gemini-1.5-pro`
*   **OpenRouter (Free Tier Available, check model details):**
    *   `gemini-2.0-flash-exp:free`
    *   `google/gemini-2.0-flash-thinking-exp-1219:free`
    *   `meta-llama/llama-3.1-405b-instruct:free`
    *   `meta-llama/llama-3.2-1b-instruct:free`
    *   `qwen/qwen-2-7b-instruct:free`
    *   `mistralai/mistral-7b-instruct:free`
    *   `microsoft/phi-3-mini-128k-instruct:free`
    *   `sophosympatheia/rogue-rose-103b-v0.2:free`

## Limitations and Important Notes

*   **Google Apps Script Execution Limit:** Google Apps Script has a 6-minute execution limit. This means the script may stop running after 6 minutes of processing. To continue, you may need to select a model from the menu again to restart processing where it left off.
*   **Free Tier Limits:** Be aware that free tiers for the AI models used in this script may have usage limits.
*   **API Rate Limits:** This script respects API rate limits for all models (RPM - Requests Per Minute) and (RPS - Requests Per Second) to ensure smooth and reliable operation.
*   **Manual Stop:** You can stop the script manually using the "Stop Script" option in the menu.
*   **Error Logging:** The script will log all errors, including rate limit issues and invalid API responses. It will stop automatically after multiple consecutive errors to prevent issues.
*   **Check Model Costs:** While many models offer free tiers, you are responsible for checking the pricing of any APIs used, as this project is designed to leverage freely available resources.

## Author

*   **Dawid Żurawski**
    *   LinkedIn: [https://www.linkedin.com/in/dawid-%C5%BCurawski-61a1341b3/](https://www.linkedin.com/in/dawid-%C5%BCurawski-61a1341b3/)

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).
