/**
 * @OnlyCurrentDoc
 * 
 * LDPT - Logistics Distribution Planning Tool
 * 
 * RECENT OPTIMIZATIONS (for improved efficiency):
 * - Reduced MAX_ROUTE_MILEAGE from 500 to 425 miles to keep routes well below 450 miles
 * - Added exponential mileage penalty in route scoring to discourage high-mileage routes
 * - Set mileage target at 85% of max (361 miles) with penalties for exceeding
 * - Reduced MAX_DISTANCE_BETWEEN_STOPS from 120 to 60 miles (hard limit 75 miles)
 *   to reduce miles between stops and create tighter routes
 * - Increased MIN_PALLETS_PER_ROUTE from 18 to 20 pallets to improve truck utilization
 * - Strengthened utilization penalty in scoring (2000 -> 5000) to discourage underutilized trucks
 * - Increased cluster penalty (500 -> 800) to keep routes more geographically focused
 * - Reduced CLUSTER_FAILURE_THRESHOLD from 5 to 3 for tighter geographic clustering
 * - Increased MAX_PLANNING_ATTEMPTS from 3 to 5 to try harder before marking shipments as unplannable
 * - Enhanced rebalanceRoutes with route consolidation to reduce total truck count
 * - Added proactive insertion logic in runPlanningLoop to fill existing routes before creating new ones
 * - Increased pull-forward minimum threshold from 10 to 15 pallets to minimize unplanned trucks
 * 
 * These changes aim to:
 * - Reduce unplanned trucks from 12-17 average to lower numbers
 * - Increase overall truck utilization
 * - Keep route distances well below 450 miles (target: ~360 miles)
 * - Reduce leg distances to 50-75 miles maximum
 * - Fully utilize available carrier resources
 */

/**
 * Creates the custom menu in the spreadsheet when the workbook is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Planning Tool')
    .addItem('Open Sidebar', 'showSidebar')
    .addItem('Generate Plan', 'generatePlan')
    .addToUi();
}

/**
 * Displays the HTML sidebar for user interaction.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(350)
    .setTitle('Team Logistics Planner');
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- SIDEBAR DATA MANAGEMENT & SETTINGS SERVICE ---

const SettingsService = {
  get: function() {
    const properties = PropertiesService.getScriptProperties();
    const storedSettings = properties.getProperty('routingSettings');
    const defaultSettings = {
      MAX_ROUTE_MILEAGE: 425, // Reduced from 500 to 425 miles to ensure routes stay well below 450
      MAX_STOPS_GENERAL: 5,
      MAX_STOPS_NY: 3,
      MAX_DISTANCE_BETWEEN_STOPS: 60, // Reduced from 120 to 60 miles for tighter routing
      ABSOLUTE_MAX_DISTANCE_BETWEEN_STOPS: 75, // Reduced from 120 to 75 miles max leg
      MAX_ROUTE_DURATION_MINS: 720, // Pre-HOS duration limit
      ALLOW_ALL_ALL: true, // Enable ALL-ALL rule for ambient/chiller/produce
      SMALL_AMBIENT_THRESHOLD: 6 // Allow up to 6 pallets when mixing ambient/produce with chiller
    };
    if (storedSettings) {
      try {
        const parsed = JSON.parse(storedSettings);
        return { ...defaultSettings, ...parsed };
      } catch (e) {
        Logger.log(`Warning: Failed to parse stored routing settings: ${e}`);
      }
    }
    return defaultSettings;
  },
  save: function(settings) {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('routingSettings', JSON.stringify(settings));
  }
};

/**
 * Gets the current routes that can be replanned.
 * @param {string} filter - Optional filter for routes (all, low-utilization, high-mileage, mixed-clusters)
 * @returns {Array} Array of route objects with relevant details for the UI
 */
function getRoutesForReplanning(filter = 'all') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Generated Plan');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const routeIdIdx = headers.indexOf('Route ID');
  const carrierIdx = headers.indexOf('Carrier');
  const utilizationIdx = headers.indexOf('Truck Utilization');
  const mileageIdx = headers.indexOf('Total Route Mileage');
  const costIdx = headers.indexOf('Estimated Cost');
  const clusterIdx = headers.indexOf('Cluster');
  const stopSeqIdx = headers.indexOf('Stop Sequence');

  const routes = new Map(); // Use Map to aggregate stops by route
  
  // Process all rows and aggregate by route
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const routeId = row[routeIdIdx];
    const stopSeq = row[stopSeqIdx];
    
    // Skip unused capacity and unplannable rows
    if (stopSeq === 'UNUSED CAPACITY' || stopSeq === 'NOT PLANNED' || 
        String(row[carrierIdx]).includes('Unplannable') || 
        String(row[carrierIdx]).includes('Overspill')) {
      continue;
    }
    
    if (!routes.has(routeId)) {
      routes.set(routeId, {
        routeId: routeId,
        carrier: row[carrierIdx],
        stops: [],
        clusters: new Set(),
        utilization: parseFloat(String(row[utilizationIdx]).replace('%', '')),
        totalMiles: parseFloat(row[mileageIdx]) || 0,
        cost: parseFloat(row[costIdx]) || 0
      });
    }
    
    const route = routes.get(routeId);
    route.stops.push(stopSeq);
    if (row[clusterIdx] && row[clusterIdx] !== 'N/A') {
      route.clusters.add(row[clusterIdx]);
    }
  }

  // Convert routes Map to array and apply filters
  let routeArray = Array.from(routes.values())
    .map(r => ({
      ...r,
      numStops: new Set(r.stops).size,
      hasMixedClusters: r.clusters.size > 1
    }));

  // Apply filters
  switch (filter) {
    case 'low-utilization':
      routeArray = routeArray.filter(r => r.utilization < 85);
      break;
    case 'high-mileage':
      routeArray = routeArray.filter(r => r.totalMiles > 350);
      break;
    case 'mixed-clusters':
      routeArray = routeArray.filter(r => r.hasMixedClusters);
      break;
  }
  return routeArray;
}

/**
 * Replans specified routes by removing their shipments and running the planning algorithm again.
 * @param {Array<string>} routeIds - Array of route IDs to replan
 */
function replanRoutes(routeIds) {
  if (!routeIds || routeIds.length === 0) {
    throw new Error('No routes selected for replanning');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Generated Plan');
  if (!sheet) throw new Error('Generated Plan sheet not found');

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) throw new Error('No plan data found');

  // Get indices for required columns
  const headers = data[0];
  const routeIdIdx = headers.indexOf('Route ID');
  const stopSeqIdx = headers.indexOf('Stop Sequence');
  
  // Collect all shipments from selected routes
  const shipmentsToReplan = [];
  const routeSet = new Set(routeIds);
  
  // Add logging to debug the collection process
  Logger.log(`Looking for shipments in routes: ${Array.from(routeSet).join(', ')}`);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const routeId = row[routeIdIdx];
    const stopSeq = row[stopSeqIdx];
    
    // Log what we find for each relevant row
    if (routeSet.has(routeId)) {
      Logger.log(`Found row for route ${routeId}: Stop=${stopSeq}, Pallets=${row[headers.indexOf('Total Pallets')]}`);
    }
    
    if (routeSet.has(routeId) && 
        stopSeq && stopSeq !== 'UNUSED CAPACITY' && 
        stopSeq !== 'NOT PLANNED' &&
        stopSeq !== 'N/A') {
      const shipment = {
        store: stopSeq,
        pallets: parseFloat(row[headers.indexOf('Total Pallets')]) || 0,
        category: row[headers.indexOf('Category')] || 'Ambient',
        palletType1: row[headers.indexOf('Pallet Type 1')] || '',
        palletType2: row[headers.indexOf('Pallet Type 2')] || '',
        attempts: 0,
        insertionFailures: 0
      };
      
      // Only add valid shipments
      if (shipment.store && shipment.pallets > 0) {
        shipmentsToReplan.push(shipment);
        Logger.log(`Added shipment: ${JSON.stringify(shipment)}`);
      } else {
        Logger.log(`Skipped invalid shipment: ${JSON.stringify(shipment)}`);
      }
    }
  }
  
  Logger.log(`Found ${shipmentsToReplan.length} shipments to replan`);

  if (shipmentsToReplan.length === 0) {
    throw new Error('No valid shipments found in selected routes');
  }

  // Remove the original routes
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (routeSet.has(data[i][routeIdIdx])) {
      rowsToDelete.push(i + 1); // +1 because sheet rows are 1-based
    }
  }
  
  // Delete rows in reverse order to maintain correct indices
  rowsToDelete.sort((a, b) => b - a).forEach(row => {
    sheet.deleteRow(row);
  });

  // Update carrier availability for replanning
  const originalCarriers = getCarriers(); // This gets fresh carrier data
  const timeSlotUsage = new Map(); // Track current usage
  
  // Count existing usage from remaining routes
  const remainingData = sheet.getDataRange().getValues();
  for (let i = 1; i < remainingData.length; i++) {
    const row = remainingData[i];
    const carrier = row[carrierIdx];
    const timeSlot = row[headers.indexOf('Time Slot')];
    const key = `${carrier}-${timeSlot}`;
    timeSlotUsage.set(key, (timeSlotUsage.get(key) || 0) + 1);
  }

  // Update carrier time slot availability
  originalCarriers.forEach(carrier => {
    carrier.timeSlots.forEach(slot => {
      const key = `${carrier.name}-${slot.time}`;
      slot.used = timeSlotUsage.get(key) || 0;
    });
  });

  // Get all the context data needed for replanning
  const settings = SettingsService.get();
  const warehouse = 'US0007';
  
  const productTypeMap = getProductTypeMap();
  const allCarriers = getCarriers();
  const restrictions = getRestrictions();
  const detailedDurations = getDetailedDurations();
  const addressData = getAddressData();
  const distanceMatrix = getDistanceMatrix(shipmentsToReplan, addressData, warehouse, "YOUR_API_KEY_HERE");

  // Create planning context
  const context = {
    ...settings,
    restrictions,
    detailedDurations,
    distanceMatrix,
    addressData,
    warehouse,
    OVERPLAN_FACTOR: 1.0,
    MIN_PALLETS_PER_ROUTE: 20, // Increased from 15 to 20 for better utilization
    MAX_PLANNING_ATTEMPTS: 5, // Increased attempts to try harder
    RELAX_MIN_ATTEMPTS: 4, // Require more attempts before relaxing
    RELAX_MIN_FACTOR: 0.8, // Less relaxation (was 0.7)
    CLUSTER_FAILURE_THRESHOLD: 3 // Reduced to keep routes more clustered
  };

  // Ensure shipments have cluster information
  shipmentsToReplan.forEach(s => {
    s.cluster = (addressData[s.store] && addressData[s.store].cluster) || 'Others';
    if (s.insertionFailures === undefined) s.insertionFailures = 0;
  });

  // Replan the shipments
  const result = runPlanningLoop(shipmentsToReplan, allCarriers, context, sheet, () => {}, 1);
  
  // Handle any remaining unplanned shipments
  if (result.remainingShipments.length > 0) {
    planRemainingShipments(result.remainingShipments, context, sheet);
  }

  // Diagnostics after replanning
  try { generateDiagnosticReport(sheet, settings); } catch(e) { Logger.log('Diagnostics error: '+e); }
}

// --- CORE DATA FETCHING FUNCTIONS ---

function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Error: '${sheetName}' sheet not found.`);
    return null;
  }
  const values = sheet.getDataRange().getValues();
  if (values.length < 1) {
    Logger.log(`Info: '${sheetName}' sheet is empty.`);
    return null;
  }
  return values;
}

/**
 * [NEW] Fetches carrier cost data.
 * Assumes a sheet named 'Carrier Costs' with 'Carrier', 'Cost Per Mile', 'Cost Per Route'.
 * @returns {object} A map of carrier names to their cost structures.
 */
function getCarrierCosts() {
  const values = getSheetData('Carrier Costs');
  if (!values || values.length <= 1) {
    Logger.log("Warning: 'Carrier Costs' sheet not found or empty. Costs will be calculated as 0.");
    return {};
  }
  const costs = {};
  const headers = values[0];
  const carrierIdx = 0; // Column A
  const costPerMileIdx = 1; // Column B
  const costPerRouteIdx = 2; // Column C
  const costNotToUseIdx = 3; // Column D

  values.slice(1).forEach(row => {
    const carrierName = String(row[carrierIdx] || '').trim();
    if (carrierName) {
      costs[carrierName] = {
        costPerMile: parseFloat(row[costPerMileIdx]) || 0,
        costPerRoute: parseFloat(row[costPerRouteIdx]) || 0,
        costNotToUse: parseFloat(row[costNotToUseIdx]) || 0
      };
    }
  });
  Logger.log(`Fetched cost data for ${Object.keys(costs).length} carriers.`);
  return costs;
}

/**
 * [UPDATED] Reads carrier data and merges it with cost data.
 * @returns {Array<object>} An array of carrier objects with capacities and costs.
 */
function getCarriers() {
  Logger.log('Starting to fetch carrier data...');
  const values = getSheetData('Carriers_2');
  if (!values || values.length <= 1) {
    Logger.log("Warning: 'Carriers_2' sheet is empty or has no data.");
    return [];
  }

  const carriers = {};
  const palletCapacities = getPalletCapacities();
  const carrierCosts = getCarrierCosts(); // Get cost data
  const headers = values[0];
  const timeSlotsHeaders = headers.slice(1);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const carrierName = String(row[0] || '').trim();
    if (!carrierName) continue;

    if (!carriers[carrierName]) {
      const costs = carrierCosts[carrierName] || { costPerMile: 0, costPerRoute: 0, costNotToUse: 0 };
      Logger.log(`Carrier ${carrierName} costs: ${JSON.stringify(costs)}`);
      carriers[carrierName] = {
        name: carrierName,
        timeSlots: [],
        pallets36: palletCapacities[carrierName]?.pallets36 || 18,
        pallets48: palletCapacities[carrierName]?.pallets48 || 22,
        pallets53: palletCapacities[carrierName]?.pallets53 || 26,
        costPerMile: costs.costPerMile,
        costPerRoute: costs.costPerRoute,
        costNotToUse: costs.costNotToUse
      };
    }

    for (let j = 0; j < timeSlotsHeaders.length; j++) {
      const time = timeSlotsHeaders[j];
      const capacity = parseInt(row[j + 1], 10) || 0;
      if (capacity > 0) {
        carriers[carrierName].timeSlots.push({
          time: String(time).trim(),
          capacity: capacity,
          used: 0
        });
      }
    }
  }

  Logger.log(`Fetched and aggregated ${Object.keys(carriers).length} carriers from 'Carriers_2' sheet.`);
  return Object.values(carriers);
}


/**
 * Helper function to get pallet capacities from Carrier Inventory sheet.
 * @returns {object}
 */
function getPalletCapacities() {
  const values = getSheetData('Carrier Inventory');
  if (!values || values.length <= 1) return {};
  const headers = values[0];
  const capacities = {};
  const carrierIndex = headers.indexOf('Carrier');
  const sizeIndex = headers.indexOf('Size');
  const maxPalletsIndex = headers.indexOf('Max Pallets');

  values.slice(1).forEach(row => {
      const carrier = String(row[carrierIndex] || '').trim();
      if (carrier) {
          if (!capacities[carrier]) {
              capacities[carrier] = { pallets36: 18, pallets48: 22, pallets53: 26 };
          }
          const size = String(row[sizeIndex] || '');
          const maxPallets = parseFloat(row[maxPalletsIndex]);
          if (size === '36') capacities[carrier].pallets36 = maxPallets;
          else if (size === '48') capacities[carrier].pallets48 = maxPallets;
          else if (size === '53') capacities[carrier].pallets53 = maxPallets;
      }
  });
  return capacities;
}

function getDistanceMatrix(shipments, addressData, warehouseLocation, apiKey) {
  const matrix = {};
  const sheetData = getSheetData('Distance Matrix');
  
  Logger.log('Starting to build distance matrix...');
  
  if (sheetData && sheetData.length > 1) {
    const headers = sheetData[0];
    const fromIndex = headers.indexOf('From');
    const toIndex = headers.indexOf('To');
    const distIndex = headers.indexOf('Distance (M)');
    const durIndex = headers.indexOf('Duration (Min)');

    if (fromIndex === -1 || toIndex === -1 || distIndex === -1 || durIndex === -1) {
      Logger.log('Error: Missing required columns in Distance Matrix sheet');
      throw new Error('Distance Matrix sheet missing required columns');
    }

    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      const from = row[fromIndex];
      const to = row[toIndex];
      if (!from || !to) continue;
      if (!matrix[from]) matrix[from] = {};
      matrix[from][to] = {
        distance: parseFloat(row[distIndex]) || 0,
        duration: parseFloat(row[durIndex]) || 0
      };
    }
    Logger.log(`Built initial distance matrix from sheet for ${Object.keys(matrix).length} origins.`);
  }

  const uniqueStoreNumbers = [...new Set(shipments.map(s => s.store)), warehouseLocation];
  const missingPairs = [];
  for (const from of uniqueStoreNumbers) {
    for (const to of uniqueStoreNumbers) {
      if ((!matrix[from] || !matrix[from][to]) && from !== to) {
        missingPairs.push({ from, to });
      }
    }
  }

  if (missingPairs.length > 0 && apiKey && apiKey !== "YOUR_API_KEY") {
    Logger.log(`Found ${missingPairs.length} missing distance pairs. Calling Google Maps API.`);
    missingPairs.forEach(pair => {
      const fromAdr = addressData[pair.from] ? `${addressData[pair.from].street}, ${addressData[pair.from].city}` : null;
      const toAdr = addressData[pair.to] ? `${addressData[pair.to].street}, ${addressData[pair.to].city}` : null;
      if (fromAdr && toAdr) {
        try {
          const directions = Maps.newDirectionFinder().setOrigin(fromAdr).setDestination(toAdr).getDirections();
          if (directions && directions.routes.length > 0) {
            const leg = directions.routes[0].legs[0];
            const distanceMiles = Math.round(leg.distance.value * 0.000621371);
            const durationMinutes = Math.round(leg.duration.value / 60);
            if (!matrix[pair.from]) matrix[pair.from] = {};
            matrix[pair.from][pair.to] = { distance: distanceMiles, duration: durationMinutes };
          }
        } catch (e) {
          Logger.log(`Could not get directions for ${pair.from} to ${pair.to}: ${e.toString()}`);
        }
      }
    });
  }

  return matrix;
}

function getRestrictions() {
  const values = getSheetData('Restrictions');
  if (!values || values.length <= 1) return {};
  const headers = values[0];
  const restrictions = {};
  values.slice(1).forEach(row => {
    const store = row[headers.indexOf('Store')];
    if (store) {
      restrictions[store] = {
        noise: row[headers.indexOf('Noise Restriction')],
        equipmentNight: row[headers.indexOf('Equipment Restriction Night')],
        equipmentDay: row[headers.indexOf('Equipment Restriciton Day')],
        details: row[headers.indexOf('Special Details')],
        deliveryWindow: row[headers.indexOf('Delivery Window')]
      };
    }
  });
  Logger.log('Fetched restrictions.');
  return restrictions;
}

function getProductTypeMap() {
  const values = getSheetData('ProductTypes');
  if (!values || values.length <= 1) return {};
  const productMap = {};
  values.slice(1).forEach(row => {
    if (row[0] && row[1]) productMap[row[1]] = row[0];
  });
  Logger.log("Created product type map.");
  return productMap;
}

/**
 * [UPDATED] Fetches and processes shipment data from Shipments sheet.
 * Aggregates shipments by store and category, and filters out Standby products.
 * @param {object} productTypeMap - Map of product types from ProductTypes sheet
 * @returns {Array<object>} Array of shipment objects
 */
function getShipments(productTypeMap) {
  const values = getSheetData('Shipments');
  if (!values || values.length <= 1) {
    Logger.log('Error: Shipments sheet is empty or has no data');
    return [];
  }
  
  const headers = values[0];
  Logger.log(`Found Shipments sheet headers: ${JSON.stringify(headers)}`);
  
  // Try different possible column name formats
  const storeIndex = headers.indexOf('Store');
  const palletType1Index = headers.indexOf('Pallet Type 1') !== -1 ? 
                          headers.indexOf('Pallet Type 1') : 
                          headers.indexOf('PalletType1');
  const palletType2Index = headers.indexOf('Pallet Type 2') !== -1 ? 
                          headers.indexOf('Pallet Type 2') : 
                          headers.indexOf('PalletType2');
  const palletsIndex = headers.indexOf('Pallets');
  const productStatusIndex = headers.indexOf('Status'); // For filtering Standby
  
  Logger.log(`Column indices - Store: ${storeIndex}, Pallet Type 1: ${palletType1Index}, Pallet Type 2: ${palletType2Index}, Pallets: ${palletsIndex}, Status: ${productStatusIndex}`);

  if ([storeIndex, palletType1Index, palletType2Index, palletsIndex].includes(-1)) {
    Logger.log('Error: Missing required columns in Shipments sheet. Make sure the following columns exist: Store, Pallet Type 1, Pallet Type 2, Pallets');
    return [];
  }

  // Use an object to aggregate shipments by store-category
  const aggregatedShipments = {};

  let validShipments = 0, invalidShipments = 0;

  values.slice(1).forEach((row, index) => {
    // Clean and validate store number
    const store = String(row[storeIndex] || '').trim();
    const palletType1 = String(row[palletType1Index] || '').trim();
    const palletType2 = String(row[palletType2Index] || '').trim();
    const pallets = parseFloat(row[palletsIndex]);
    const status = productStatusIndex !== -1 ? String(row[productStatusIndex] || '').toLowerCase() : '';
    
    // Log validation details for troubleshooting
    if (!store || !palletType1 || isNaN(pallets)) {
      invalidShipments++;
      Logger.log(`Invalid shipment at row ${index + 2}: Store='${store}', Type1='${palletType1}', Pallets=${pallets}`);
    } else {
      validShipments++;
    }

    // Skip invalid rows and Standby products
    if (!store || !palletType1 || isNaN(pallets)) return;
    if (status === 'standby') return;

    // Determine category
    let category = 'Ambient';
    const palletTypeLower = String(palletType1).toLowerCase();
    if (['chiller', 'meat'].includes(palletTypeLower)) category = 'Chiller';
    else if (['freezer', 'freezertkt'].includes(palletTypeLower)) category = 'Freezer';

    // Get product type and create key
    const productTypeId = productTypeMap[palletType1] || null;
    const key = `${store}-${category}`;

    // Aggregate by store-category
    if (aggregatedShipments[key]) {
      aggregatedShipments[key].pallets += pallets;
    } else {
      aggregatedShipments[key] = {
        store,
        category,
        pallets,
        palletType1,
        palletType2,
        productTypeId,
        attempts: 0,
        insertionFailures: 0,
        failureReason: null
      };
    }
  });

  const shipments = Object.values(aggregatedShipments);
  Logger.log(`Processed ${validShipments + invalidShipments} total rows: ${validShipments} valid, ${invalidShipments} invalid`);
  Logger.log(`Aggregated into ${shipments.length} unique store-category combinations`);
  
  if (shipments.length === 0) {
    Logger.log('Warning: No valid shipments found after processing');
  }
  
  return shipments;
}

function getDetailedDurations() {
  const values = getSheetData('Adress Product Type Duration');
  if (!values || values.length <= 1) return {};
  const durations = {};
  // ...existing code to populate durations...
  return durations;
}

/**
 * [NEW] Fetches address and cluster data for all store locations.
 * Reads from Addresses sheet with required columns for addresses and clustering.
 * @returns {object} Map of store numbers to their address and cluster info
 */
function getAddressData() {
  const values = getSheetData('Addresses');
  if (!values || values.length <= 1) {
    Logger.log('Error: Addresses sheet not found or empty');
    return {};
  }

  const headers = values[0];
  const requiredColumns = ['Store Number', 'Name', 'Street', 'City', 'ZipCode', 'Cluster'];
  const columnIndices = {};
  
  // Get indices for all required columns
  requiredColumns.forEach(col => {
    columnIndices[col] = headers.indexOf(col);
    if (columnIndices[col] === -1) {
      Logger.log(`Warning: Column '${col}' not found in Addresses sheet`);
    }
  });

  const addresses = {};
  values.slice(1).forEach(row => {
    const storeNumber = row[columnIndices['Store Number']];
    if (storeNumber) {
      addresses[storeNumber] = {
        name: row[columnIndices['Name']] || '',
        street: row[columnIndices['Street']] || '',
        city: row[columnIndices['City']] || '',
        zip: row[columnIndices['ZipCode']] || '',
        cluster: row[columnIndices['Cluster']] || 'Others'
      };
    }
  });

  Logger.log(`Loaded address data for ${Object.keys(addresses).length} stores`);
  return addresses;
}

// --- SHEET SETUP & UTILITY FUNCTIONS ---

function setupPlanSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Generated Plan');
  if (sheet) sheet.clear(); else sheet = ss.insertSheet('Generated Plan');
  const headers = [
    'Carrier', 'Route ID', 'Cluster', 'Stops', 'Stop Sequence', 'Detailed Route',
    'Total Route Mileage', 'Has Restrictions?', 'Trailer Size', 'Temp Zones', 'Total Pallets',
    'Truck Utilization', 'Total Travel (min)', 'Total Stop Time (min)', 'Total Duration (min)',
    'HOS Status', 'Estimated Cost', 'Mileage Status', 'Notes', 'Map Link'
  ];
  sheet.appendRow(headers).getRange("A1:T1").setFontWeight("bold");
  sheet.setFrozenRows(1);
  return sheet;
}

// --- DIAGNOSTICS ---

function generateDiagnosticReport(planSheet, settings) {
  try {
    const stats = { 
      totalRoutes: 0, 
      totalPallets: 0, 
      totalMileage: 0, 
      totalDuration: 0, 
      hosViolations: 0, 
      overMileage: 0, 
      utilization: [],
      mileageByRoute: [] // Track individual route mileages for better analysis
    };
    const planData = planSheet.getDataRange().getValues();
    if (!planData || planData.length <= 1) {
      Logger.log('Diagnostic: No planned routes to report.');
      return;
    }
    const headers = planData[0];
    const palletsIdx = headers.indexOf('Total Pallets');
    const milesIdx = headers.indexOf('Total Route Mileage');
    const durationIdx = headers.indexOf('Total Duration (min)');
    const hosIdx = headers.indexOf('HOS Status');
    const utilIdx = headers.indexOf('Truck Utilization');
    const mileageStatusIdx = headers.indexOf('Mileage Status');
    const routeIdIdx = headers.indexOf('Route ID');

    const parseNumber = (value) => {
      if (typeof value === 'number') return value;
      if (typeof value === 'string') {
        const cleaned = value.replace(/[^0-9.-]/g, '');
        const parsed = parseFloat(cleaned);
        return isNaN(parsed) ? 0 : parsed;
      }
      return 0;
    };

    for (let i = 1; i < planData.length; i++) {
      const row = planData[i];
      if (!row || String(row[0] || '').includes('Unplannable') || String(row[0] || '').includes('Overspill')) continue;
      stats.totalRoutes++;
      const pallets = parseNumber(row[palletsIdx]);
      stats.totalPallets += pallets;
      const routeMiles = parseNumber(row[milesIdx]);
      stats.totalMileage += routeMiles;
      stats.mileageByRoute.push({ routeId: row[routeIdIdx], miles: routeMiles });
      stats.totalDuration += parseNumber(row[durationIdx]);
      if (row[hosIdx] && row[hosIdx].toString().toUpperCase() !== 'OK') stats.hosViolations++;
      const mileageStatus = mileageStatusIdx !== -1 ? String(row[mileageStatusIdx] || '').toUpperCase() : '';
      if (mileageStatus === 'OVER' || routeMiles > (settings.MAX_ROUTE_MILEAGE || 425)) stats.overMileage++;
      const utilStr = String(row[utilIdx] || '');
      const utilPct = utilStr.endsWith('%') ? parseFloat(utilStr) / 100 : parseNumber(utilStr);
      if (utilPct > 0 && utilPct <= 1.5) stats.utilization.push(utilPct);
    }
    
    // Calculate mileage statistics
    const avgUtil = stats.utilization.length ? (stats.utilization.reduce((a,b)=>a+b,0)/stats.utilization.length) : 0;
    const avgMileage = stats.totalRoutes ? (stats.totalMileage / stats.totalRoutes) : 0;
    const sortedMileages = stats.mileageByRoute.map(r => r.miles).sort((a,b) => a-b);
    const maxMileage = sortedMileages.length > 0 ? sortedMileages[sortedMileages.length - 1] : 0;
    const minMileage = sortedMileages.length > 0 ? sortedMileages[0] : 0;
    
    Logger.log('--- PLAN DIAGNOSTIC REPORT ---');
    Logger.log('Total Routes: ' + stats.totalRoutes);
    Logger.log('Total Pallets: ' + stats.totalPallets);
    Logger.log('Total Mileage: ' + stats.totalMileage);
    Logger.log('Avg Pallets/Truck: ' + (stats.totalRoutes ? (stats.totalPallets/stats.totalRoutes).toFixed(2) : 0));
    Logger.log('Avg Utilization: ' + (avgUtil*100).toFixed(1) + '%');
    Logger.log('Avg Route Mileage: ' + avgMileage.toFixed(1) + ' miles');
    Logger.log('Min Route Mileage: ' + minMileage.toFixed(1) + ' miles');
    Logger.log('Max Route Mileage: ' + maxMileage.toFixed(1) + ' miles');
    Logger.log('Routes Over Mileage: ' + stats.overMileage);
    Logger.log('HOS Violations: ' + stats.hosViolations);
    
    // Log the top 5 longest routes for debugging
    if (stats.mileageByRoute.length > 0) {
      const topRoutes = stats.mileageByRoute.sort((a,b) => b.miles - a.miles).slice(0, 5);
      Logger.log('Top 5 Longest Routes:');
      topRoutes.forEach(r => Logger.log(`  ${r.routeId}: ${r.miles.toFixed(1)} miles`));
    }
    
    Logger.log('-----------------------------');
  } catch (e) {
    Logger.log('Error in diagnostic report: ' + e);
  }
}

function parseTimeToMinutes(timeStr) {
  if (!timeStr || typeof timeStr !== 'string') return null;
  let match = timeStr.match(/^(\d{1,2}):(\d{2})$/);
  if (match) return parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
  match = timeStr.match(/^(\d{1,2})(AM|PM)$/i);
  if (match) {
    let hours = parseInt(match[1], 10);
    if (hours === 12) hours = (match[2].toUpperCase() === 'AM') ? 0 : 12;
    else if (match[2].toUpperCase() === 'PM') hours += 12;
    return hours * 60;
  }
  return null;
}

function isTimeInWindow(routeTimeStr, windowStr) {
  if (!windowStr || typeof windowStr !== 'string' || windowStr.trim().toUpperCase() === 'N/A' || windowStr.trim() === '') return true;
  const parts = windowStr.split('-');
  if (parts.length !== 2) return true;
  const routeTime = parseTimeToMinutes(routeTimeStr);
  const startTime = parseTimeToMinutes(parts[0].trim());
  const endTime = parseTimeToMinutes(parts[1].trim());
  if (routeTime === null || startTime === null || endTime === null) return true;
  return startTime <= endTime ? (routeTime >= startTime && routeTime <= endTime) : (routeTime >= startTime || routeTime <= endTime);
}

/**
 * Determines if a set of product categories can co-load on a single trailer.
 * Default: disallow Chiller + Freezer together.
 * New behavior: allow ALL-ALL mixes when all stores/categories are identical or when
 * a small ambient/produce quantity can be carried with freezer (configurable threshold).
 * @param {Set<string>} requiredCategories - set of category strings present on the route/stops
 * @param {object} [options] - optional overrides: { allowAllAll: boolean, smallAmbientThreshold: number, details: {perStoreCounts} }
 */
function isTempCompatible(requiredCategories, options) {
  options = options || {};
  const allowAllAll = options.allowAllAll || false;
  const smallAmbientThreshold = options.smallAmbientThreshold || 6;
  const perStoreCounts = options.details && options.details.perStoreCounts ? options.details.perStoreCounts : null;

  const hasAmbient = requiredCategories.has('Ambient');
  const hasProduce = requiredCategories.has('Produce');
  const hasChiller = requiredCategories.has('Chiller');
  const hasFreezer = requiredCategories.has('Freezer');

  // Freezer rules: cannot mix with Chiller, and cannot mix with Ambient/Produce
  if (hasFreezer) {
    if (hasChiller) return false;
    if (hasAmbient || hasProduce) return false;
    return true;
  }

  // No freezer present: optionally allow ALL-ALL (Ambient+Chiller+Produce) with small ambient/produce override
  if (allowAllAll) {
    if (!perStoreCounts) return true;
    return Object.keys(perStoreCounts).every(storeId => {
      const counts = perStoreCounts[storeId] || {};
      const hasStoreChiller = (counts.Chiller || 0) > 0;
      const ambientTotal = (counts.Ambient || 0) + (counts.Produce || 0);
      // If store has chiller plus ambient/produce, enforce threshold
      if (hasStoreChiller && ambientTotal > smallAmbientThreshold) return false;
      return true;
    });
  }

  // Default fallback: disallow Chiller+Ambient mixes without override? Keep previous behaviour (only block Chiller+Freezer)
  return !(hasChiller && hasFreezer);
}

function formatStopSequence(stops) {
  const stopsByStore = {};
  stops.forEach(stop => {
    if (!stopsByStore[stop.store]) stopsByStore[stop.store] = new Set();
    stopsByStore[stop.store].add(stop.category);
  });
  const uniqueStoresInOrder = [...new Set(stops.map(s => s.store))];
  return uniqueStoresInOrder.map(store => {
    const catAbbr = [...stopsByStore[store]].map(c => c.substring(0, 3).toUpperCase()).join(',');
    return `${store} (${catAbbr})`;
  }).join(' / ');
}


// --- CORE ROUTING & PLANNING LOGIC ---

/**
 * Main function to generate the transportation plan.
 */
function generatePlan() {
  const API_KEY = "YOUR_API_KEY_HERE";
  const WAREHOUSE_LOCATION = 'US0007';
  
  const settings = SettingsService.get();
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(HtmlService.createHtmlOutput('<p>Planning in progress...</p>'), 'Running Planner');
  
  const planSheet = setupPlanSheet();
  
  let shipments, allCarriers, restrictions, detailedDurations, addressData, distanceMatrix;
  try {
    Logger.log('Starting data collection...');
    const productTypeMap = getProductTypeMap();
    Logger.log('Got product type map');
    
    shipments = getShipments(productTypeMap);
    Logger.log(`Got ${shipments.length} shipments`);
    
    allCarriers = getCarriers();
    Logger.log(`Got ${allCarriers.length} carriers`);
    
    restrictions = getRestrictions();
    Logger.log('Got restrictions');
    
    detailedDurations = getDetailedDurations();
    Logger.log('Got detailed durations');
    
    addressData = getAddressData();
    Logger.log('Got address data');
    
    // Verify we have shipments to process
    if (!shipments || shipments.length === 0) {
      ui.alert('Error', 'No valid shipments found in the Shipments sheet. Please check the data format.', ui.ButtonSet.OK);
      Logger.log('Error: No shipments to process');
      return;
    }

    Logger.log('Starting distance matrix calculation...');
    distanceMatrix = getDistanceMatrix(shipments, addressData, WAREHOUSE_LOCATION, API_KEY);
    
    // Assign clusters and initialize insertion failures
    Logger.log('Assigning clusters to shipments...');
    shipments.forEach(s => {
      if (!s || !s.store) {
        Logger.log(`Warning: Invalid shipment found: ${JSON.stringify(s)}`);
        return;
      }
      s.cluster = (addressData[s.store] && addressData[s.store].cluster) || 'Others';
      if (s.insertionFailures === undefined) s.insertionFailures = 0;
    });
  } catch (e) {
    ui.alert('Error loading data. Check script logs.');
    Logger.log("Data Loading Error: " + e.toString() + e.stack);
    return;
  }

  const context = {
    ...settings,
    restrictions,
    detailedDurations,
    distanceMatrix,
    addressData,
    warehouse: WAREHOUSE_LOCATION,
    OVERPLAN_FACTOR: 1.0,
    MIN_PALLETS_PER_ROUTE: 20, // Increased from 18 to 20 for better utilization
    MAX_PLANNING_ATTEMPTS: 5, // Increased attempts to try harder
    RELAX_MIN_ATTEMPTS: 4, // Require more attempts before relaxing
    RELAX_MIN_FACTOR: 0.8, // Less relaxation (was 0.7)
    CLUSTER_FAILURE_THRESHOLD: 3, // Reduced from 5 to 3 to keep routes more clustered
  };
  
  let carriersForPlanning = JSON.parse(JSON.stringify(allCarriers));

  let { remainingShipments } = runPlanningLoop(shipments, carriersForPlanning, context, planSheet, ()=>{}, 1);
  
  planRemainingShipments(remainingShipments, context, planSheet);

  const plannedCarrierNames = new Set();
  const generatedPlanValues = planSheet.getDataRange().getValues();
  const carrierColIndex = 0;
  for (let i = 1; i < generatedPlanValues.length; i++) {
    const carrierName = generatedPlanValues[i][carrierColIndex];
    if (carrierName && !carrierName.includes('Overspill') && !carrierName.includes('Unplannable')) {
      plannedCarrierNames.add(carrierName);
    }
  }

  allCarriers.forEach(originalCarrier => {
    const wasUsed = plannedCarrierNames.has(originalCarrier.name);
    const plannedCarrier = carriersForPlanning.find(c => c.name === originalCarrier.name);

    if (wasUsed) {
      originalCarrier.timeSlots.forEach(originalTimeSlot => {
        const plannedTimeSlot = plannedCarrier.timeSlots.find(ts => ts.time === originalTimeSlot.time);
        const unusedCapacity = originalTimeSlot.capacity - (plannedTimeSlot ? plannedTimeSlot.used : 0);

        if (unusedCapacity > 0) {
          const reason = `Unused Capacity at ${originalTimeSlot.time}`;
          for (let i = 0; i < unusedCapacity; i++) {
            planSheet.appendRow([
              originalCarrier.name, 'N/A', 'N/A', 0, 'UNUSED CAPACITY', 'N/A',
              'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A',
              reason
            ]).getRange(planSheet.getLastRow(), 1, 1, 20).setBackground('#fce5cd');
          }
        }
      });
    } else {
      const reason = diagnoseUnusedCarrier(originalCarrier, shipments, context);
      originalCarrier.timeSlots.forEach(timeSlot => {
        for (let i = 0; i < timeSlot.capacity; i++) {
          planSheet.appendRow([
            originalCarrier.name, 'N/A', 'N/A', 0, 'NOT PLANNED', 'N/A',
            'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A',
            reason
          ]).getRange(planSheet.getLastRow(), 1, 1, 20).setBackground('#fce5cd');
        }
      });
    }
  });

  updateDashboard();

  // Generate diagnostics after plan build
  try { generateDiagnosticReport(planSheet, settings); } catch(e) { Logger.log('Diagnostics error: '+e); }

  ui.alert('Success', 'Planning complete! Check the Dashboard sheet for analytics.', ui.ButtonSet.OK);
}

function runPlanningLoop(shipmentsToPlan, availableCarriers, context, planSheet, log, startIteration) {
    let unassignedShipments = [...shipmentsToPlan];
    let iteration = startIteration;
    const {
        MAX_PLANNING_ATTEMPTS,
        MIN_PALLETS_PER_ROUTE,
        restrictions,
        RELAX_MIN_ATTEMPTS = 2,
        RELAX_MIN_FACTOR = 0.7,
        CLUSTER_FAILURE_THRESHOLD = 2
    } = context;
    let routes = []; // Track all created routes for rebalancing

    let planningInProgress = true;
    while (planningInProgress) {
        unassignedShipments.sort((a, b) => b.pallets - a.pallets);
        
        // Before creating new routes, try to fit unassigned shipments into existing routes
        if (routes.length > 0) {
            const insertedIndices = [];
            for (let i = unassignedShipments.length - 1; i >= 0; i--) {
                const shipment = unassignedShipments[i];
                if (shipment.failureReason) continue;
                
                let bestRouteIndex = -1;
                let bestScore = Infinity;
                
                for (let r = 0; r < routes.length; r++) {
                    const route = routes[r];
                    const score = checkAndScoreInsertion(shipment, route, context, log, iteration);
                    if (score !== null && score < bestScore) {
                        bestScore = score;
                        bestRouteIndex = r;
                    }
                }
                
                if (bestRouteIndex >= 0) {
                    routes[bestRouteIndex].stops.push(shipment);
                    routes[bestRouteIndex].totalPallets += shipment.pallets;
                    insertedIndices.push(i);
                }
            }
            
            // Remove inserted shipments from unassigned list
            insertedIndices.forEach(idx => unassignedShipments.splice(idx, 1));
            
            if (insertedIndices.length > 0) {
                Logger.log(`Inserted ${insertedIndices.length} shipments into existing routes`);
            }
        }
        
        let seedIndex = unassignedShipments.findIndex(s => s.attempts < MAX_PLANNING_ATTEMPTS && !s.failureReason);
        
        if (seedIndex === -1) {
            // Try rebalancing routes before finishing
            if (routes.length > 0) {
                rebalanceRoutes(routes, context);
                routes.forEach(route => {
                    writeRouteToSheet(route, planSheet, context);
                });
            }
            planningInProgress = false;
            continue;
        }

        let seedShipment = unassignedShipments.splice(seedIndex, 1)[0];
        seedShipment.attempts++;

        let bestRouteOption = findBestInitialCarrier(seedShipment, availableCarriers, context);

        if (!bestRouteOption) {
            const storeRes = restrictions[seedShipment.store];
            if (storeRes && storeRes.deliveryWindow && storeRes.deliveryWindow.trim() !== '' && storeRes.deliveryWindow.toUpperCase() !== 'N/A') {
                seedShipment.failureReason = 'Time Constraint';
            } else {
                seedShipment.failureReason = 'No Viable Carrier';
            }
            unassignedShipments.push(seedShipment);
            continue;
        }

        const currentRoute = bestRouteOption.route;
        currentRoute.seedAttempts = seedShipment.attempts;
        let routeChanged = true;
        while (routeChanged) {
            routeChanged = false;
            let bestCandidate = null, bestScore = Infinity;

            for (let i = unassignedShipments.length - 1; i >= 0; i--) {
                const candidate = unassignedShipments[i];
                if (candidate.failureReason) continue;
                if (candidate.insertionFailures === undefined) candidate.insertionFailures = 0;

                const differentCluster = currentRoute.cluster !== 'Others' && candidate.cluster !== currentRoute.cluster;
                const canRelaxCluster = differentCluster && (candidate.insertionFailures >= CLUSTER_FAILURE_THRESHOLD || currentRoute.seedAttempts >= RELAX_MIN_ATTEMPTS);
                if (differentCluster && !canRelaxCluster) {
                    candidate.insertionFailures++;
                    continue;
                }

                const score = checkAndScoreInsertion(candidate, currentRoute, context, log, iteration);
                if (score !== null && score < bestScore) {
                    bestScore = score;
                    bestCandidate = { shipment: candidate, index: i };
                } else if (score === null) {
                    candidate.insertionFailures++;
                }
            }

            if (bestCandidate) {
                const addedShipment = unassignedShipments.splice(bestCandidate.index, 1)[0];
                currentRoute.stops.push(addedShipment);
                currentRoute.totalPallets += addedShipment.pallets;
                addedShipment.insertionFailures = 0;
                if (currentRoute.cluster === 'Others' && addedShipment.cluster !== 'Others') {
                    currentRoute.cluster = addedShipment.cluster;
                }
                routeChanged = true;
            }
        }

        const relaxedMinimum = Math.max(1, Math.ceil(MIN_PALLETS_PER_ROUTE * (currentRoute.seedAttempts >= RELAX_MIN_ATTEMPTS ? RELAX_MIN_FACTOR : 1)));

        if (currentRoute.totalPallets < relaxedMinimum) {
            currentRoute.stops.forEach(stop => {
                stop.insertionFailures = (stop.insertionFailures || 0) + 1;
                unassignedShipments.push(stop);
            });
        } else {
            bestRouteOption.timeSlot.used++;
            currentRoute.routeId = `${currentRoute.carrier.name.replace(/\s+/g, '')}_${iteration++}`;
            routes.push(currentRoute); // Add to routes array instead of writing immediately
        }
    }
    
    // After main loop, try to utilize any remaining carriers with pull-forward requests
    const { remainingShipments: finalUnassigned } = utilizePullForwardRequests(unassignedShipments, availableCarriers, context, routes);
    
    return { remainingShipments: finalUnassigned, iteration };
}

function findBestInitialCarrier(shipment, carriers, context) {
    let bestOption = null;
    let bestCost = Infinity;

    for (const carrier of carriers) {
        for (const timeSlot of carrier.timeSlots) {
            if (timeSlot.used >= timeSlot.capacity) continue;
            
            const tempRoute = { carrier, time: timeSlot.time, stops: [shipment], totalPallets: shipment.pallets, cluster: shipment.cluster };
            if (checkAndScoreInsertion(shipment, { ...tempRoute, stops:[] }, context, ()=>{}, 0) === null) continue;

            const metrics = calculateRouteMetricsWithHOS([shipment], context.warehouse, context.distanceMatrix, context.detailedDurations);
            const cost = (metrics.totalDistance * carrier.costPerMile) + carrier.costPerRoute;

            if (cost < bestCost) {
                bestCost = cost;
                bestOption = { carrier, timeSlot, route: tempRoute };
            }
        }
    }
    return bestOption;
}

function checkAndScoreInsertion(shipment, route, context, log, iter) {
  const {
    restrictions,
    distanceMatrix,
    warehouse,
    OVERPLAN_FACTOR,
    MAX_ROUTE_MILEAGE,
    MAX_STOPS_GENERAL,
    MAX_DISTANCE_BETWEEN_STOPS,
    ABSOLUTE_MAX_DISTANCE_BETWEEN_STOPS
  } = context;

  // Validate input parameters
  if (!shipment || !route || !context || !distanceMatrix || !warehouse) {
    Logger.log('Error: Missing required parameters in checkAndScoreInsertion');
    return null;
  }

  const tempStops = [...route.stops, shipment];
  if ([...new Set(tempStops.map(s => s.store))].length > MAX_STOPS_GENERAL) return null;

  // Build per-store category pallet counts for the small-ambient override
  const perStoreCounts = {};
  tempStops.forEach(s => {
    if (!perStoreCounts[s.store]) perStoreCounts[s.store] = {};
    perStoreCounts[s.store][s.category] = (perStoreCounts[s.store][s.category] || 0) + (s.pallets || 0);
  });

  if (!isTempCompatible(new Set(tempStops.map(s => s.category)), { allowAllAll: context.ALLOW_ALL_ALL || false, smallAmbientThreshold: context.SMALL_AMBIENT_THRESHOLD || 4, details: { perStoreCounts } })) return null;

  if (restrictions[shipment.store] && restrictions[shipment.store].deliveryWindow && restrictions[shipment.store].deliveryWindow.trim() !== '' && restrictions[shipment.store].deliveryWindow.toUpperCase() !== 'N/A' && !isTimeInWindow(route.time, restrictions[shipment.store].deliveryWindow)) return null;

  const carrier = route.carrier;
  const requiredSize = getRequiredTrailerSize({ ...route, stops: tempStops }, restrictions);
  
  let capacity = 0;
  switch (requiredSize) {
      case '36': capacity = carrier.pallets36; break;
      case '48': capacity = carrier.pallets48; break;
      case '53': capacity = carrier.pallets53; break;
  }
  
  if (!capacity || (route.totalPallets || 0) + shipment.pallets > (capacity * OVERPLAN_FACTOR)) return null;

  const currentOptimized = advancedOptimizeStopOrder(route.stops, warehouse, distanceMatrix);
  const tempOptimized = advancedOptimizeStopOrder(tempStops, warehouse, distanceMatrix);
  const preferredLegThreshold = MAX_DISTANCE_BETWEEN_STOPS || 50;
  const hardLegLimit = ABSOLUTE_MAX_DISTANCE_BETWEEN_STOPS || Math.max(75, preferredLegThreshold);

  const currentLegDetails = getLegDistanceDetails(currentOptimized, distanceMatrix);
  if (currentLegDetails.hasMissingData) return null;

  const proposedLegDetails = getLegDistanceDetails(tempOptimized, distanceMatrix);
  if (proposedLegDetails.hasMissingData) return null;

  const proposedMaxLeg = proposedLegDetails.legDistances.length > 0 ? Math.max(...proposedLegDetails.legDistances) : 0;
  if (proposedMaxLeg > hardLegLimit) return null;

  const distancePenaltyDelta = calculateDistancePenalty(proposedLegDetails.legDistances, preferredLegThreshold)
    - calculateDistancePenalty(currentLegDetails.legDistances, preferredLegThreshold);

  const tempMetrics = calculateRouteMetricsWithHOS(tempOptimized, warehouse, distanceMatrix, context.detailedDurations);
  if (tempMetrics.totalDistance > MAX_ROUTE_MILEAGE || tempMetrics.hosStatus !== 'OK') return null;

  const originalMetrics = calculateRouteMetricsWithHOS(currentOptimized, warehouse, distanceMatrix, context.detailedDurations);

  const newCost = (tempMetrics.totalDistance * carrier.costPerMile) + carrier.costPerRoute;
  const oldCost = (originalMetrics.totalDistance * carrier.costPerMile) + carrier.costPerRoute;

  return (newCost - oldCost) + distancePenaltyDelta; // Return cost delta adjusted for leg distance penalties
}

function planRemainingShipments(shipments, context, planSheet) {
  const unplannableTime = shipments.filter(s => s.failureReason === 'Time Constraint');
  const unplannableOther = shipments.filter(s => s.failureReason && s.failureReason !== 'Time Constraint');
  const overspillShipments = shipments.filter(s => !s.failureReason);

  if (unplannableTime.length > 0) {
    writeRouteToSheet({ 
      carrier: { 
        name: 'Unplannable (Time)', 
        costPerMile: 0, 
        costPerRoute: 0, 
        costNotToUse: 0 
      }, 
      routeId: 'Time_1', 
      stops: unplannableTime, 
      totalPallets: unplannableTime.reduce((s, p) => s + p.pallets, 0), 
      cluster: 'Manual Review', 
      notes: 'Failed due to delivery window constraints.' 
    }, planSheet, context);
  }
  if (unplannableOther.length > 0) {
    writeRouteToSheet({ 
      carrier: { 
        name: 'Unplannable (Other)', 
        costPerMile: 0, 
        costPerRoute: 0, 
        costNotToUse: 0 
      }, 
      routeId: 'Other_1', 
      stops: unplannableOther, 
      totalPallets: unplannableOther.reduce((s, p) => s + p.pallets, 0), 
      cluster: 'Manual Review', 
      notes: 'Failed planning after max attempts or no viable carrier.' 
    }, planSheet, context);
  }
  
  if (overspillShipments.length === 0) return;

  let overspillRoutes = [];
  overspillShipments.sort((a,b) => b.pallets - a.pallets);

  overspillShipments.forEach(shipment => {
    let added = false;
    for (const route of overspillRoutes) {
      if (checkAndScoreInsertion(shipment, route, context, ()=>{}, 999) !== null) {
        route.stops.push(shipment);
        route.totalPallets += shipment.pallets;
        added = true;
        break;
      }
    }
    if (!added) {
      overspillRoutes.push({
        carrier: { name: 'Overspill', pallets53: 26, pallets48: 22, pallets36: 18, costPerMile: 0, costPerRoute: 0 },
        time: '23:00',
        stops: [shipment],
        totalPallets: shipment.pallets,
        cluster: 'Overspill'
      });
    }
  });
  
  overspillRoutes.forEach((route, i) => {
    route.routeId = `Overspill_${i + 1}`;
    writeRouteToSheet(route, planSheet, context);
  });
}

function writeRouteToSheet(route, planSheet, context) {
  if (!route.stops || route.stops.length === 0) return;

  const { warehouse, distanceMatrix, detailedDurations, restrictions, addressData, MAX_ROUTE_MILEAGE } = context;
  const optimizedStops = advancedOptimizeStopOrder(route.stops, warehouse, distanceMatrix);
  const metrics = calculateRouteMetricsWithHOS(optimizedStops, warehouse, distanceMatrix, detailedDurations);
  const trailerSize = getRequiredTrailerSize(route, restrictions);
  
  // Log warning for routes approaching or exceeding mileage limits
  if (metrics.totalDistance > MAX_ROUTE_MILEAGE * 0.85) {
    const stores = [...new Set(optimizedStops.map(s => s.store))].join(' -> ');
    Logger.log(`WARNING: Route ${route.routeId} has ${metrics.totalDistance} miles (${(metrics.totalDistance/MAX_ROUTE_MILEAGE*100).toFixed(1)}% of max). Stores: ${stores}`);
  }
  
  let capacity = 0;
  if (route.carrier.name.includes('Unplannable') || route.carrier.name.includes('Overspill')) {
      capacity = 999;
  } else {
      switch (trailerSize) {
          case '36': capacity = route.carrier.pallets36; break;
          case '48': capacity = route.carrier.pallets48; break;
          case '53': capacity = route.carrier.pallets53; break;
      }
  }
  const utilization = capacity > 0 && capacity < 999 ? `${(route.totalPallets / capacity * 100).toFixed(1)}%` : 'N/A';
  const cost = (metrics.totalDistance * (route.carrier.costPerMile || 0)) + (route.carrier.costPerRoute || 0);
  const mapLink = generateMapLink(optimizedStops, warehouse, addressData);

  planSheet.appendRow([
    route.carrier.name, route.routeId, route.cluster,
    [...new Set(optimizedStops.map(s => s.store))].length, formatStopSequence(optimizedStops),
    formatDetailedRoute(optimizedStops, warehouse, distanceMatrix), metrics.totalDistance,
    checkRouteForRestrictions(optimizedStops, restrictions), `${trailerSize}'`,
    'Dual Temp', route.totalPallets.toFixed(2), utilization,
    metrics.travelTime, metrics.stopTime, metrics.totalDuration,
    metrics.hosStatus,
    cost > 0 ? cost.toFixed(2) : 'N/A',
    metrics.totalDistance > MAX_ROUTE_MILEAGE ? 'OVER' : 'OK',
    route.notes || '',
    mapLink
  ]);
}

// --- ROUTE OPTIMIZATION & METRIC CALCULATION ---

function getRequiredTrailerSize(route, restrictions) {
  let requires36 = false;
  let requires48 = false;
  
  const uniqueStores = [...new Set(route.stops.map(s => s.store))];
  
  uniqueStores.forEach(store => {
    const res = restrictions[store];
    if (res) {
      const routeTimeMins = parseTimeToMinutes(route.time) || 0;
      const isNight = routeTimeMins >= 1140 || routeTimeMins < 360; 
      const equipment = isNight ? res.equipmentNight : res.equipmentDay;
      
      if (equipment) {
        if (equipment.includes("36")) requires36 = true;
        if (equipment.includes("48")) requires48 = true;
      }
    }
  });

  if (requires36) return '36';
  if (requires48) return '48';
  return '53';
}

function getLegDistanceDetails(stops, distanceMatrix) {
  if (!stops || !distanceMatrix || stops.length === 0) {
    return { legDistances: [], hasMissingData: false };
  }

  const uniqueStores = [...new Set(stops.map(s => s.store))];
  if (uniqueStores.length < 2) {
    return { legDistances: [], hasMissingData: false };
  }

  const legDistances = [];
  for (let i = 1; i < uniqueStores.length; i++) {
    const from = uniqueStores[i - 1];
    const to = uniqueStores[i];
    if (!from || !to) continue;

    const forwardLeg = distanceMatrix[from] && distanceMatrix[from][to];
    const reverseLeg = distanceMatrix[to] && distanceMatrix[to][from];
    const leg = forwardLeg || reverseLeg;

    if (!leg) {
      Logger.log(`Warning: Missing distance data between ${from} and ${to}`);
      return { legDistances: [], hasMissingData: true };
    }

    legDistances.push(leg.distance);
  }

  return { legDistances, hasMissingData: false };
}

function calculateDistancePenalty(legDistances, preferredThreshold) {
  if (!legDistances || legDistances.length === 0) return 0;
  const threshold = preferredThreshold || 0;
  return legDistances.reduce((penalty, distance) => {
    if (distance <= threshold) return penalty;
    const overage = distance - threshold;
    return penalty + (overage * overage);
  }, 0);
}

function formatDetailedRoute(orderedStops, warehouse, distanceMatrix) {
  if (!orderedStops || orderedStops.length === 0) return "";
  const uniqueStores = [...new Set(orderedStops.map(s => s.store))];
  let routeStr = "", lastLocation = warehouse;
  uniqueStores.forEach(store => {
    const dist = (distanceMatrix[lastLocation] && distanceMatrix[lastLocation][store]) ? distanceMatrix[lastLocation][store].distance : 0;
    routeStr += routeStr ? ` (${dist} mi) -> ${store}` : `${store}`;
    lastLocation = store;
  });
  return routeStr;
}

function checkRouteForRestrictions(stops, restrictions) {
  return stops.some(s => {
    const res = restrictions[s.store];
    return res && ((res.noise && res.noise !== 'N/A') || (res.equipmentNight && res.equipmentNight !== 'N/A') || (res.equipmentDay && res.equipmentDay !== 'N/A') || (res.deliveryWindow && res.deliveryWindow.trim() !== '' && res.deliveryWindow !== 'N/A'));
  }) ? "Yes" : "No";
}

function advancedOptimizeStopOrder(stops, warehouse, distanceMatrix) {
  if (stops.length <= 1) return stops;

  const stopsData = {};
  stops.forEach(s => {
    if (!stopsData[s.store]) stopsData[s.store] = { shipments: [], categories: new Set() };
    stopsData[s.store].shipments.push(s);
    stopsData[s.store].categories.add(s.category);
  });

  const ambientOnly = [], withCold = [];
  for (const store in stopsData) {
    if (stopsData[store].categories.has('Chiller') || stopsData[store].categories.has('Freezer')) withCold.push(store);
    else ambientOnly.push(store);
  }

  let finalOrder = [], lastLocation = warehouse;
  if (ambientOnly.length > 0) {
    const orderedAmbient = runNearestNeighbor(ambientOnly, warehouse, distanceMatrix);
    finalOrder.push(...orderedAmbient);
    if (orderedAmbient.length > 0) lastLocation = orderedAmbient[orderedAmbient.length - 1];
  }
  if (withCold.length > 0) {
    finalOrder.push(...runNearestNeighbor(withCold, lastLocation, distanceMatrix));
  }

  return finalOrder.flatMap(store => stopsData[store].shipments);
}

function runNearestNeighbor(items, startLocation, distanceMatrix) {
  if (!items || items.length < 2) return items;
  let ordered = [], remaining = [...items], currentLocation = startLocation;
  const isObject = typeof remaining[0] === 'object';

  while (remaining.length > 0) {
    let nearestDist = Infinity, nearestIndex = -1;
    for (let i = 0; i < remaining.length; i++) {
      const store = isObject ? remaining[i].store : remaining[i];
      const dist = (distanceMatrix[currentLocation] && distanceMatrix[currentLocation][store]) ? distanceMatrix[currentLocation][store].distance : Infinity;
      if (dist < nearestDist) {
        nearestDist = dist;
        nearestIndex = i;
      }
    }
    if (nearestIndex === -1) {
      ordered.push(...remaining); break;
    }
    const nextItem = remaining.splice(nearestIndex, 1)[0];
    ordered.push(nextItem);
    currentLocation = isObject ? nextItem.store : nextItem;
  }
  return ordered;
}

/**
 * [NEW] Calculates route metrics including Hours of Service (HOS) breaks.
 * Adds a 30-min break for every 8 hours of on-duty time.
 * @returns {object} Metrics including total duration with breaks and HOS status.
 */
function calculateRouteMetricsWithHOS(orderedStops, warehouse, distanceMatrix, detailedDurations) {
  let travelTime = 0, totalDistance = 0, stopTime = 0, currentLocation = warehouse;
  const PALLETS_PER_FLS = 26;
  const ON_DUTY_LIMIT = 14 * 60; // 14 hours
  const DRIVING_LIMIT = 11 * 60; // 11 hours
  const BREAK_THRESHOLD = 8 * 60; // 8 hours
  const BREAK_DURATION = 30; // 30 minutes

  // Calculate initial stop time at warehouse (loading)
  orderedStops.forEach(s => {
    const pDur = detailedDurations[s.store] ? detailedDurations[s.store][s.productTypeId] : null;
    if (pDur && pDur.loading) {
      stopTime += (s.pallets / PALLETS_PER_FLS) * pDur.loading;
    }
  });

  let onDutyTime = stopTime;
  let breaksTaken = 0;
  let hosStatus = 'OK';

  // Iterate through stops to calculate travel and on-site time
  [...new Set(orderedStops.map(s => s.store))].forEach(store => {
    const leg = (distanceMatrix[currentLocation] && distanceMatrix[currentLocation][store]);
    if (leg) {
      if (onDutyTime + leg.duration > (breaksTaken + 1) * BREAK_THRESHOLD) {
        onDutyTime += BREAK_DURATION;
        breaksTaken++;
      }
      travelTime += leg.duration;
      // Accumulate the leg distance for travel between stops (includes warehouse->first stop)
      totalDistance += leg.distance;
      onDutyTime += leg.duration;
    }
    
    let timeAtThisStop = 0;
    orderedStops.forEach(s => {
      if (s.store === store) {
        const pDur = detailedDurations[s.store] ? detailedDurations[s.store][s.productTypeId] : null;
        if (pDur && pDur.unloading) {
          timeAtThisStop += (s.pallets / PALLETS_PER_FLS) * pDur.unloading;
        } else {
          timeAtThisStop += 15; // Default unload time
        }
      }
    });
    stopTime += timeAtThisStop;
    onDutyTime += timeAtThisStop;
    currentLocation = store;
  });

  // Calculate return leg
  const returnLeg = (distanceMatrix[currentLocation] && distanceMatrix[currentLocation][warehouse]);
  if (returnLeg) {
    if (onDutyTime + returnLeg.duration > (breaksTaken + 1) * BREAK_THRESHOLD) {
      onDutyTime += BREAK_DURATION;
      breaksTaken++;
    }
    travelTime += returnLeg.duration;
    onDutyTime += returnLeg.duration;
    totalDistance += returnLeg.distance;
  }
  
  // Final HOS checks
  if (travelTime > DRIVING_LIMIT) hosStatus = 'DRIVING_VIOLATION';
  if (onDutyTime > ON_DUTY_LIMIT) hosStatus = 'ON_DUTY_VIOLATION';

  // Log computed metrics for debugging and validation
  Logger.log(`calculateRouteMetricsWithHOS: stops=${orderedStops.length}, travelTime=${Math.round(travelTime)}, stopTime=${Math.round(stopTime)}, totalDuration=${Math.round(onDutyTime)}, totalDistance=${Math.round(totalDistance)}, hosStatus=${hosStatus}`);

  return {
    travelTime: Math.round(travelTime),
    stopTime: Math.round(stopTime),
    totalDuration: Math.round(onDutyTime),
    totalDistance: Math.round(totalDistance),
    hosStatus: hosStatus
  };
}


// --- ROUTE OPTIMIZATION & REBALANCING FUNCTIONS ---

function rebalanceRoutes(routes, context) {
  const { OVERPLAN_FACTOR } = context;
  let improved = true;
  let iterations = 0;
  const MAX_REBALANCE_ITERATIONS = 15; // Increased from 10 to 15 for more thorough optimization
  
  Logger.log(`Starting route rebalancing with ${routes.length} routes`);
  
  while (improved && iterations < MAX_REBALANCE_ITERATIONS) {
    improved = false;
    iterations++;
    
    // First pass: Try to consolidate routes by moving all shipments from one route to another
    for (let i = routes.length - 1; i >= 0; i--) {
      if (routes[i].stops.length === 0) continue;
      
      for (let j = 0; j < routes.length; j++) {
        if (i === j) continue;
        
        const sourceRoute = routes[i];
        const targetRoute = routes[j];
        
        // Skip if routes are from different carriers or time slots
        if (sourceRoute.carrier.name !== targetRoute.carrier.name || sourceRoute.time !== targetRoute.time) continue;
        
        // Try to move all shipments from source to target
        let canMoveAll = true;
        const tempStops = [...targetRoute.stops];
        const tempPallets = targetRoute.totalPallets;
        
        for (const stop of sourceRoute.stops) {
          const testRoute = {
            ...targetRoute,
            stops: tempStops,
            totalPallets: tempPallets + stop.pallets
          };
          
          if (checkAndScoreInsertion(stop, testRoute, context, () => {}, 0) === null) {
            canMoveAll = false;
            break;
          }
          tempStops.push(stop);
        }
        
        if (canMoveAll && isRouteValid({...targetRoute, stops: tempStops, totalPallets: tempPallets + sourceRoute.totalPallets}, context)) {
          // Consolidate: move all stops from source to target
          targetRoute.stops.push(...sourceRoute.stops);
          targetRoute.totalPallets += sourceRoute.totalPallets;
          routes.splice(i, 1); // Remove empty source route
          improved = true;
          Logger.log(`Consolidated route ${sourceRoute.routeId} into ${targetRoute.routeId}`);
          break;
        }
      }
    }
    
    // Second pass: Try to swap shipments between routes to improve utilization
    for (let i = 0; i < routes.length; i++) {
      for (let j = i + 1; j < routes.length; j++) {
        const route1 = routes[i];
        const route2 = routes[j];
        
        // Skip if routes are from different carriers or time slots
        if (route1.carrier.name !== route2.carrier.name || route1.time !== route2.time) continue;
        
        // Try swapping each shipment from route1 with each from route2
        for (let si = 0; si < route1.stops.length; si++) {
          for (let sj = 0; sj < route2.stops.length; sj++) {
            const stop1 = route1.stops[si];
            const stop2 = route2.stops[sj];
            
            // Calculate current metrics
            const currentScore = calculateRouteScore(route1, context) + calculateRouteScore(route2, context);
            
            // Try swap
            route1.stops[si] = stop2;
            route2.stops[sj] = stop1;
            
            // Update pallet counts
            const newRoute1Pallets = route1.stops.reduce((sum, s) => sum + s.pallets, 0);
            const newRoute2Pallets = route2.stops.reduce((sum, s) => sum + s.pallets, 0);
            const oldRoute1Pallets = route1.totalPallets;
            const oldRoute2Pallets = route2.totalPallets;
            
            route1.totalPallets = newRoute1Pallets;
            route2.totalPallets = newRoute2Pallets;
            
            // Calculate new metrics
            const newScore = calculateRouteScore(route1, context) + calculateRouteScore(route2, context);
            
            if (newScore < currentScore && isRouteValid(route1, context) && isRouteValid(route2, context)) {
              // Keep the swap
              improved = true;
            } else {
              // Revert the swap
              route1.stops[si] = stop1;
              route2.stops[sj] = stop2;
              route1.totalPallets = oldRoute1Pallets;
              route2.totalPallets = oldRoute2Pallets;
            }
          }
        }
      }
    }
  }
  
  Logger.log(`Rebalancing complete after ${iterations} iterations. Final route count: ${routes.length}`);
}

function calculateRouteScore(route, context) {
  const metrics = calculateRouteMetricsWithHOS(route.stops, context.warehouse, context.distanceMatrix, context.detailedDurations);
  const trailerSize = getRequiredTrailerSize(route, context.restrictions);
  let capacity = 0;
  switch (trailerSize) {
    case '36': capacity = route.carrier.pallets36; break;
    case '48': capacity = route.carrier.pallets48; break;
    case '53': capacity = route.carrier.pallets53; break;
  }
  
  // Calculate actual costs
  const distanceCost = metrics.totalDistance * (route.carrier.costPerMile || 0);
  const routeCost = route.carrier.costPerRoute || 0;
  const totalCost = distanceCost + routeCost;
  
  // Penalize under-utilization more aggressively
  const utilization = route.totalPallets / capacity;
  // Much stronger penalty for under-utilization to push towards fuller trucks
  const utilizationPenalty = Math.pow(1 - utilization, 2) * 5000; // Increased from 2000 to 5000
  
  // Add penalty for high mileage routes to keep them under target
  const MAX_ROUTE_MILEAGE = context.MAX_ROUTE_MILEAGE || 425;
  const MILEAGE_TARGET = MAX_ROUTE_MILEAGE * 0.85; // Target 85% of max (361 miles for 425 max)
  let mileagePenalty = 0;
  if (metrics.totalDistance > MILEAGE_TARGET) {
    // Exponential penalty for routes exceeding target mileage
    const overage = metrics.totalDistance - MILEAGE_TARGET;
    mileagePenalty = Math.pow(overage / 10, 2) * 100; // Scaled penalty
  }
  
  // Optional: add penalty for cluster mixing (if route.stops have >1 cluster)
  let clusterPenalty = 0;
  if (route.stops && route.stops.length > 1) {
    const clusters = new Set(route.stops.map(s => s.cluster || ''));
    if (clusters.size > 1) clusterPenalty = 800; // Increased from 500 to 800
  }
  
  return totalCost + utilizationPenalty + mileagePenalty + clusterPenalty;
}

function isRouteValid(route, context) {
  const { MAX_ROUTE_MILEAGE, MAX_STOPS_GENERAL, OVERPLAN_FACTOR } = context;
  
  // Check basic constraints
  if ([...new Set(route.stops.map(s => s.store))].length > MAX_STOPS_GENERAL) return false;
  // Build per-store counts for temperature compatibility checks
  const perStoreCounts = {};
  route.stops.forEach(s => {
    if (!perStoreCounts[s.store]) perStoreCounts[s.store] = {};
    perStoreCounts[s.store][s.category] = (perStoreCounts[s.store][s.category] || 0) + (s.pallets || 0);
  });
  if (!isTempCompatible(new Set(route.stops.map(s => s.category)), { allowAllAll: context.ALLOW_ALL_ALL || false, smallAmbientThreshold: context.SMALL_AMBIENT_THRESHOLD || 4, details: { perStoreCounts } })) return false;
  
  // Check trailer capacity
  const trailerSize = getRequiredTrailerSize(route, context.restrictions);
  let capacity = 0;
  switch (trailerSize) {
    case '36': capacity = route.carrier.pallets36; break;
    case '48': capacity = route.carrier.pallets48; break;
    case '53': capacity = route.carrier.pallets53; break;
  }
  if (route.totalPallets > capacity * OVERPLAN_FACTOR) return false;
  
  // Check route metrics
  const metrics = calculateRouteMetricsWithHOS(route.stops, context.warehouse, context.distanceMatrix, context.detailedDurations);
  if (metrics.totalDistance > MAX_ROUTE_MILEAGE || metrics.hosStatus !== 'OK') return false;
  
  return true;
}

function utilizePullForwardRequests(remainingShipments, carriers, context, existingRoutes) {
  const { MIN_PALLETS_PER_ROUTE, RELAX_MIN_FACTOR = 0.8 } = context;
  let unassigned = [...remainingShipments];
  let routesCreated = [];
  // Require higher fill for pull-forward routes to minimize unplanned trucks
  const relaxedMinimum = Math.max(15, Math.ceil(MIN_PALLETS_PER_ROUTE * RELAX_MIN_FACTOR)); // Increased from 10

  // Find unused carrier slots
  carriers.forEach(carrier => {
    carrier.timeSlots.forEach(timeSlot => {
      if (timeSlot.used < timeSlot.capacity) {
        const unusedSlots = timeSlot.capacity - timeSlot.used;
        
        for (let i = 0; i < unusedSlots; i++) {
          // Try to create a route with remaining shipments
          let potentialStops = [];
          let totalPallets = 0;
          let j = 0;
          
          const extracted = [];
          while (j < unassigned.length && totalPallets < MIN_PALLETS_PER_ROUTE) {
            const shipment = unassigned[j];
            const tempRoute = {
              carrier,
              time: timeSlot.time,
              stops: [...potentialStops, shipment],
              totalPallets: totalPallets + shipment.pallets
            };

            if (checkAndScoreInsertion(shipment, tempRoute, context, () => {}, 0) !== null) {
              potentialStops.push(shipment);
              totalPallets += shipment.pallets;
              shipment.insertionFailures = 0;
              extracted.push(unassigned.splice(j, 1)[0]);
            } else {
              shipment.insertionFailures = (shipment.insertionFailures || 0) + 1;
              j++;
            }
          }

          if (totalPallets >= relaxedMinimum) {
            const route = {
              carrier,
              time: timeSlot.time,
              stops: potentialStops,
              totalPallets,
              routeId: `${carrier.name.replace(/\s+/g, '')}_PF_${routesCreated.length + 1}`,
              cluster: potentialStops[0].cluster,
              notes: 'Pull Forward Request'
            };

            routesCreated.push(route);
            timeSlot.used++;
          } else if (extracted.length > 0) {
            unassigned.push(...extracted);
          }
        }
      }
    });
  });

  // Add new routes to existing routes array
  existingRoutes.push(...routesCreated);
  
  return { remainingShipments: unassigned };
}

// --- DIAGNOSTIC & UTILITY HELPER FUNCTIONS ---

/**
 * [NEW] Generates a Google Maps URL for a given route.
 * @returns {string} A clickable URL to visualize the route.
 */
function generateMapLink(orderedStops, warehouse, addressData) {
  if (!orderedStops || orderedStops.length === 0 || !addressData) return "N/A";

  const getAddressString = (storeId) => {
    const adr = addressData[storeId];
    if (!adr) return null;
    return encodeURIComponent(`${adr.street}, ${adr.city}, ${adr.zip}`);
  };

  const origin = getAddressString(warehouse);
  const destination = getAddressString(warehouse); // Route returns to origin
  const waypoints = [...new Set(orderedStops.map(s => s.store))]
    .map(getAddressString)
    .filter(Boolean) // Remove any nulls if address not found
    .join('|');

  if (!origin) return "N/A";

  return `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${destination}&waypoints=${waypoints}`;
}

function diagnoseUnusedCarrier(carrier, shipments, context) {
  if (!carrier.timeSlots || carrier.timeSlots.length === 0) {
    return "FAIL: No time slots defined in 'Carriers_2' sheet.";
  }

  if (!shipments || shipments.length === 0) {
    return "FAIL: No shipments to evaluate";
  }

  const unusedCost = carrier.costNotToUse || 0;
  
  // Filter out invalid shipments
  const validShipments = shipments.filter(s => s && s.store && s.pallets);
  if (validShipments.length === 0) {
    return "FAIL: No valid shipments to evaluate";
  }
  
  for (const shipment of validShipments) {
    for (const timeSlot of carrier.timeSlots) {
      const tempRoute = { carrier, time: timeSlot.time };
      const failureReason = getInitialRouteFailureReason(shipment, tempRoute, context);
      if (failureReason === null) {
        return `NOTE: Carrier is valid but was not selected as the optimal choice. Cost not to use: $${unusedCost.toFixed(2)}`;
      }
    }
  }

  const firstShipment = validShipments[0];
  const firstTimeSlot = carrier.timeSlots[0];
  const representativeReason = getInitialRouteFailureReason(firstShipment, { carrier, time: firstTimeSlot.time }, context);
  
  return `FAIL: Rejected for all shipments. Example Reason: ${representativeReason}. Cost not to use: $${unusedCost.toFixed(2)}`;
}

function getInitialRouteFailureReason(shipment, route, context) {
    const { restrictions, distanceMatrix, warehouse, MAX_ROUTE_MILEAGE } = context;

    if (!shipment || !shipment.store) {
        return "Invalid shipment data";
    }

    if (!distanceMatrix || !distanceMatrix[warehouse]) {
        return "Missing distance matrix data";
    }

    if (!distanceMatrix[warehouse][shipment.store]) {
        return `Missing distance data from warehouse to store ${shipment.store}.`;
    }

    const storeRestriction = restrictions[shipment.store];
    if (storeRestriction && storeRestriction.deliveryWindow && storeRestriction.deliveryWindow.trim() !== '' && storeRestriction.deliveryWindow.toUpperCase() !== 'N/A' && !isTimeInWindow(route.time, storeRestriction.deliveryWindow)) {
        return `Time slot ${route.time} is outside store ${shipment.store}'s window (${storeRestriction.deliveryWindow}).`;
    }
    
    const requiredSize = getRequiredTrailerSize({ stops: [shipment] }, restrictions);
    let capacity = 0;
    switch (requiredSize) {
        case '36': capacity = route.carrier.pallets36; break;
        case '48': capacity = route.carrier.pallets48; break;
        case '53': capacity = route.carrier.pallets53; break;
    }
    if (!capacity || shipment.pallets > capacity) {
        return `Shipment of ${shipment.pallets} pallets exceeds trailer capacity of ${capacity}.`;
    }

    const metrics = calculateRouteMetricsWithHOS([shipment], warehouse, distanceMatrix, context.detailedDurations);
    if (metrics.hosStatus !== 'OK') {
        return `Single-stop route to ${shipment.store} violates HOS rules (${metrics.hosStatus}).`;
    }
    if (metrics.totalDistance > MAX_ROUTE_MILEAGE) {
        return `Single-stop route to ${shipment.store} exceeds max mileage (${metrics.totalDistance} > ${MAX_ROUTE_MILEAGE}).`;
    }

    return null;
}

// --- [UPDATED] DASHBOARD & ANALYTICS FUNCTION ---

function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planSheet = ss.getSheetByName('Generated Plan');
  if (!planSheet) return;

  const planData = planSheet.getDataRange().getValues();
  if (planData.length <= 1) return;

  let dashboardSheet = ss.getSheetByName('Dashboard');
  if (dashboardSheet) {
    dashboardSheet.clear();
    dashboardSheet.getCharts().forEach(chart => dashboardSheet.removeChart(chart));
  } else {
    dashboardSheet = ss.insertSheet('Dashboard');
  }

  const headers = planData[0];
  const dataRows = planData.slice(1);

  const carrierIdx = headers.indexOf('Carrier');
  const palletsIdx = headers.indexOf('Total Pallets');
  const utilizationIdx = headers.indexOf('Truck Utilization');
  const clusterIdx = headers.indexOf('Cluster');
  const mileageIdx = headers.indexOf('Total Route Mileage');
  const durationIdx = headers.indexOf('Total Duration (min)');
  const stopSeqIdx = headers.indexOf('Stop Sequence');
  const costIdx = headers.indexOf('Estimated Cost');
  const notesIdx = headers.indexOf('Notes');

  const stats = {
    totalRoutes: 0, totalPallets: 0, totalMileage: 0, totalDuration: 0, totalCost: 0,
    unplannablePallets: 0, overspillPallets: 0, unusedCapacity: 0,
    carrierData: {}, clusterData: {}, unplannableReasons: {}
  };

  dataRows.forEach(row => {
    const carrier = row[carrierIdx];
    const pallets = parseFloat(row[palletsIdx]) || 0;
    const stopSequence = row[stopSeqIdx];

    if (stopSequence === 'UNUSED CAPACITY' || stopSequence === 'NOT PLANNED') {
      stats.unusedCapacity++;
      return;
    }

    if (carrier.includes('Unplannable')) {
      stats.unplannablePallets += pallets;
      const reason = (String(row[notesIdx]).match(/Failed due to (.*)\./) || [])[1] || 'Other';
      stats.unplannableReasons[reason] = (stats.unplannableReasons[reason] || 0) + 1;
      return;
    }
    if (carrier.includes('Overspill')) stats.overspillPallets += pallets;

    stats.totalRoutes++;
    stats.totalPallets += pallets;
    stats.totalMileage += parseFloat(row[mileageIdx]) || 0;
    stats.totalDuration += parseFloat(row[durationIdx]) || 0;
    stats.totalCost += parseFloat(row[costIdx]) || 0;

    if (!stats.carrierData[carrier]) {
      stats.carrierData[carrier] = { routes: 0, pallets: 0, totalUtilization: 0, utilizationCount: 0, cost: 0 };
    }
    stats.carrierData[carrier].routes++;
    stats.carrierData[carrier].pallets += pallets;
    stats.carrierData[carrier].cost += parseFloat(row[costIdx]) || 0;
    const utilStr = String(row[utilizationIdx]);
    if (utilStr && utilStr !== 'N/A') {
      stats.carrierData[carrier].totalUtilization += parseFloat(utilStr.replace('%', ''));
      stats.carrierData[carrier].utilizationCount++;
    }

    const cluster = row[clusterIdx];
    if (cluster && cluster !== 'N/A') {
      if (!stats.clusterData[cluster]) stats.clusterData[cluster] = { routes: 0 };
      stats.clusterData[cluster].routes++;
    }
  });

  // --- Write KPIs ---
  dashboardSheet.getRange('A1:B1').merge().setValue('Overall Plan Summary').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#e0e0e0');
  const kpis = [
    ['Total Planned Routes', stats.totalRoutes],
    ['Total Planned Pallets', stats.totalPallets.toFixed(2)],
    ['Total Mileage', stats.totalMileage.toFixed(0)],
    ['Total Estimated Cost', `$${stats.totalCost.toFixed(2)}`],
    ['Avg. Cost Per Mile', `$${(stats.totalCost / stats.totalMileage || 0).toFixed(2)}`],
    ['Unplannable Pallets', stats.unplannablePallets.toFixed(2)],
    ['Overspill Pallets', stats.overspillPallets.toFixed(2)],
    ['Unused Carrier Slots', stats.unusedCapacity]
  ];
  dashboardSheet.getRange('A2:B9').setValues(kpis);
  dashboardSheet.getRange('A1:B9').setBorder(true, true, true, true, true, true);

  // --- Prepare Chart Data ---
  const carrierChartData = [['Carrier', 'Routes', 'Total Pallets', 'Total Cost']];
  for (const name in stats.carrierData) {
    carrierChartData.push([name, stats.carrierData[name].routes, stats.carrierData[name].pallets, stats.carrierData[name].cost]);
  }
  const carrierDataRange = dashboardSheet.getRange(11, 1, carrierChartData.length, carrierChartData[0].length);
  carrierDataRange.setValues(carrierChartData).setBorder(true, true, true, true, true, true);
  dashboardSheet.getRange(11, 1, 1, carrierChartData[0].length).setFontWeight('bold').setBackground('#e0e0e0');

  const clusterChartData = [['Cluster', 'Number of Routes']];
  for (const name in stats.clusterData) {
    clusterChartData.push([name, stats.clusterData[name].routes]);
  }
  const clusterDataRange = dashboardSheet.getRange(11, 6, clusterChartData.length, clusterChartData[0].length);
  clusterDataRange.setValues(clusterChartData).setBorder(true, true, true, true, true, true);
  dashboardSheet.getRange(11, 6, 1, clusterChartData[0].length).setFontWeight('bold').setBackground('#e0e0e0');

  const unplannableChartData = [['Reason', 'Count']];
  for (const reason in stats.unplannableReasons) {
    unplannableChartData.push([reason, stats.unplannableReasons[reason]]);
  }
  const unplannableDataRange = dashboardSheet.getRange(11, 9, unplannableChartData.length, unplannableChartData[0].length);
  unplannableDataRange.setValues(unplannableChartData).setBorder(true, true, true, true, true, true);
  dashboardSheet.getRange(11, 9, 1, unplannableChartData[0].length).setFontWeight('bold').setBackground('#e0e0e0');


  // --- Create Charts ---
  if (carrierChartData.length > 1) {
    dashboardSheet.insertChart(dashboardSheet.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(dashboardSheet.getRange(11, 1, carrierChartData.length, 2)).setPosition(2, 4, 0, 0)
      .setOption('title', 'Routes per Carrier').build());
    dashboardSheet.insertChart(dashboardSheet.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(dashboardSheet.getRange(11, 1, carrierChartData.length, 1))
      .addRange(dashboardSheet.getRange(11, 4, carrierChartData.length, 1))
      .setPosition(2, 11, 0, 0).setOption('title', 'Total Cost per Carrier').build());
  }
  if (clusterChartData.length > 1) {
    dashboardSheet.insertChart(dashboardSheet.newChart().setChartType(Charts.ChartType.PIE)
      .addRange(clusterDataRange).setPosition(20, 4, 0, 0)
      .setOption('title', 'Route Distribution by Cluster').build());
  }
  if (unplannableChartData.length > 1) {
    dashboardSheet.insertChart(dashboardSheet.newChart().setChartType(Charts.ChartType.PIE)
      .addRange(unplannableDataRange).setPosition(20, 11, 0, 0)
      .setOption('title', 'Reasons for Unplannable Shipments').build());
  }

  dashboardSheet.autoResizeColumns(1, 11);
}
