/**
 * This script retrieves asset group performance data from Google Ads and writes it to a Google Sheet.
 * @author Dawid Żurawski
 * LinkedIn: https://www.linkedin.com/in/dawid-%C5%BCurawski-61a1341b3/
 *
 *  // First, make a copy of this template sheet: 
 *  // https://docs.google.com/spreadsheets/d/1Lcef-sKrw0Bz02xDUMnl2Mhe10v9DcYbdqLg7PAyNzI/copy
 *  // Then, paste the URL of YOUR copied sheet into the SPREADSHEET_URL variable below.
 * 
 *  // WARNING: Do not modify anything below the DAYS_BACK constant unless you know what you are doing.
 */
function main() {
  // Paste your Google Sheet URL here, between the single quotes.
  const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/...';
  // Specify the name of the sheet in your Google Sheet where the data will be written.
  const SHEET_NAME = 'PMax Data & Alternatives';
  // Set the number of days back from yesterday for which you want to fetch data.
  const DAYS_BACK = 7;
  
  // Do not modify below this line unless you know what you are doing

  // Get today's date and calculate the start and end dates for the report
  const today = new Date();
  const yesterday = new Date(today.getTime() - (1 * 24 * 60 * 60 * 1000));
  const startDate = new Date(yesterday.getTime() - (DAYS_BACK * 24 * 60 * 60 * 1000));
  const formattedStartDate = Utilities.formatDate(startDate, "GMT", "yyyy-MM-dd");
  const formattedEndDate = Utilities.formatDate(yesterday, "GMT", "yyyy-MM-dd");

  Logger.log(`Date range: ${formattedStartDate} to ${formattedEndDate}`);

  // Query to fetch asset group performance data
  const conversionQuery = `
    SELECT
      asset_group.id,
      asset_group.name,
      asset_group.status,
      metrics.conversions,
      metrics.cost_per_conversion,
      metrics.clicks,
      metrics.ctr,
      metrics.impressions,
      campaign.status
    FROM asset_group
    WHERE segments.date BETWEEN '${formattedStartDate}' AND '${formattedEndDate}'
      AND campaign.status = 'ENABLED'
  `;

  Logger.log("Fetching asset group data...");
  const conversionReport = AdsApp.report(conversionQuery);
  const assetGroupConversions = {};
  const conversionRows = conversionReport.rows();

  let processedGroups = 0;
  // Loop through each row of the asset group performance data
  while (conversionRows.hasNext()) {
    const row = conversionRows.next();
    const assetGroupId = row['asset_group.id'];
    const assetGroupName = row['asset_group.name'];
    const status = row['asset_group.status'];
    const campaignStatus = row['campaign.status'];
    const conversions = parseFloat(row['metrics.conversions']) || 0;
    const costPerConversion = parseFloat(row['metrics.cost_per_conversion']) / 1000000 || 0;
    const clicks = parseFloat(row['metrics.clicks']) || 0;
    const ctr = parseFloat(row['metrics.ctr']) || 0;
    const impressions = parseFloat(row['metrics.impressions']) || 0;

    Logger.log(`Group ID: ${assetGroupId} | Name: ${assetGroupName} | Status: ${status} | Campaign Status: ${campaignStatus} | Conversions: ${conversions} | Cost per Conversion: ${costPerConversion} | Clicks: ${clicks} | CTR: ${ctr}% | Impressions: ${impressions}`);

    // Store the asset group data in an object
    assetGroupConversions[assetGroupId] = {
      assetGroupName: assetGroupName,
      status: status,
      conversions: conversions,
      costPerConversion: costPerConversion,
      clicks: clicks,
      ctr: ctr,
      impressions: impressions
    };
    processedGroups++;
  }

    // Check if any active asset groups were found
  if (Object.keys(assetGroupConversions).length === 0) {
    Logger.log('No active asset groups found for the selected date range.');
    return;
  }

  Logger.log(`Processed ${processedGroups} asset groups.`);

    // Query to fetch asset details for the active asset groups and campaigns
  const assetGroupQuery = `
    SELECT
      asset_group.id,
      asset_group.name,
      campaign.name,
      campaign.id,
      campaign.status,
      asset_group_asset.field_type,
      asset_group_asset.performance_label,
      asset_group_asset.status,
      asset.text_asset.text
    FROM
      asset_group_asset
    WHERE
      asset_group.id IN (${Object.keys(assetGroupConversions).join(',')})
      AND asset_group.status = 'ENABLED'
      AND campaign.status = 'ENABLED'
      AND asset_group_asset.status != 'REMOVED'
  `;

  Logger.log("Fetching asset details...");
  const report = AdsApp.report(assetGroupQuery);
  const rows = report.rows();

  // Get the Google Sheet object
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    // If the sheet does not exist, create it. If it exists, clear all contents
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  } else {
    sheet.clearContents();
  }

  // Clear formatting from the range before merging
  sheet.getRange('A1:P1').clearFormat();

    // Merge cells in row 1 and set the headers for the sheet
  sheet.getRange('A1:B1').merge().setValue('Campaign');
  sheet.getRange('C1:G1').merge().setValue('Performance');
  sheet.getRange('H1:I1').merge().setValue('Report Date');
  sheet.getRange('J1:L1').merge().setValue('Asset Data');
  sheet.getRange('M1').setValue('');
  sheet.getRange('N1:P1').merge().setValue('AI Proposals');

  // Set headers for the columns in row 2
  const headers = [
    'Campaign Name',
    'Asset Group Name',
    'Asset Group Conversions',
    'Cost per Conversion',
    'CTR',
    'Clicks',
    'Impressions',
    'Report Start Date',
    'Report End Date',
    'Performance Label',
    'Asset Status',
    'Field Type',
    'Asset Text',
    'Alternative 1',
    'Alternative 2',
    'Alternative 3'
  ];

  sheet.appendRow(headers);

    // Initialize the data array and the set for tracking processed combinations
  const data = [];
  const processedCombinations = new Set();

    // Define the order of field types and performance labels for sorting
  const fieldTypeOrder = {
    'HEADLINE': 1,
    'LONG_HEADLINE': 2,
    'DESCRIPTION': 3
  };

  const performanceLabelOrder = {
    'LOW': 1,
    'GOOD': 2,
    'BEST': 3,
    'UNKNOWN': 4
  };

  let processedAssets = 0;
   // Loop through each row of the asset details data
  while (rows.hasNext()) {
    const row = rows.next();
    const fieldType = row['asset_group_asset.field_type'];

        // Check if the field type is one of the supported types
    if (['DESCRIPTION', 'HEADLINE', 'LONG_HEADLINE'].includes(fieldType)) {
      const assetGroupId = row['asset_group.id'];
      const campaignName = row['campaign.name'];
      const assetGroupName = row['asset_group.name'];
      const conversions = assetGroupConversions[assetGroupId].conversions || 0;
      const costPerConversion = assetGroupConversions[assetGroupId].costPerConversion || 0;
      const clicks = assetGroupConversions[assetGroupId].clicks || 0;
      const ctr = assetGroupConversions[assetGroupId].ctr || 0;
      const impressions = assetGroupConversions[assetGroupId].impressions || 0;
      const performanceLabel = row['asset_group_asset.performance_label'] || 'UNKNOWN';
      const assetStatus = row['asset_group_asset.status'];
      const text = row['asset.text_asset.text'] || 'N/A';

            // Create a unique key for each combination of campaign, asset group, field type, performance label, status, and text
      const uniqueKey = `${campaignName}_${assetGroupName}_${fieldType}_${performanceLabel}_${assetStatus}_${text}_${conversions}`;

          // Check if this combination has already been processed
      if (!processedCombinations.has(uniqueKey)) {
        data.push([
          campaignName,                // Campaign Name
          assetGroupName,              // Asset Group Name
          conversions,                 // Asset Group Conversions
          costPerConversion,           // Cost per Conversion
          ctr,                         // CTR
          clicks,                      // Clicks
          impressions,                 // Impressions
          formattedStartDate,          // Report Start Date
          formattedEndDate,            // Report End Date
          performanceLabel,            // Performance Label
          assetStatus,                 // Asset Status
          fieldType,                   // Field Type
          text                          // Asset Text
        ]);
        processedCombinations.add(uniqueKey);
        processedAssets++;
      }
    }
  }

  Logger.log(`Processed ${processedAssets} assets.`);

  // Sort the data based on the field type order and performance label order
  data.sort((a, b) => {
    const fieldTypeOrderComparison = fieldTypeOrder[a[11]] - fieldTypeOrder[b[11]];
    if (fieldTypeOrderComparison !== 0) {
      return fieldTypeOrderComparison;
    }

    const performanceComparison = performanceLabelOrder[a[9]] - performanceLabelOrder[b[9]];
    if (performanceComparison !== 0) {
      return performanceComparison;
    }

    return 0; // If no difference, no need to sort further
  });

    // Write the data to the Google Sheet
  data.forEach(row => sheet.appendRow(row));

  Logger.log('Asset group data has been written to the sheet, sorted by Field Type and Performance Label.');
}
