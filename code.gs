/**
 * @OnlyCurrentDoc
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
    if (storedSettings) {
      return JSON.parse(storedSettings);
    }
    return {
      MAX_ROUTE_MILEAGE: 400,
      MAX_STOPS_GENERAL: 5,
      MAX_STOPS_NY: 3,
      MAX_DISTANCE_BETWEEN_STOPS: 50,
      MAX_ROUTE_DURATION_MINS: 720 // Note: This is pre-HOS calculation
    };
  },
  save: function(settings) {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('routingSettings', JSON.stringify(settings));
  }
};

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
  const carrierIdx = headers.indexOf('Carrier');
  const costPerMileIdx = headers.indexOf('Cost Per Mile');
  const costPerRouteIdx = headers.indexOf('Cost Per Route');

  values.slice(1).forEach(row => {
    const carrierName = String(row[carrierIdx] || '').trim();
    if (carrierName) {
      costs[carrierName] = {
        costPerMile: parseFloat(row[costPerMileIdx]) || 0,
        costPerRoute: parseFloat(row[costPerRouteIdx]) || 0
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
      const costs = carrierCosts[carrierName] || { costPerMile: 0, costPerRoute: 0 };
      carriers[carrierName] = {
        name: carrierName,
        timeSlots: [],
        pallets36: palletCapacities[carrierName]?.pallets36 || 18,
        pallets48: palletCapacities[carrierName]?.pallets48 || 22,
        pallets53: palletCapacities[carrierName]?.pallets53 || 26,
        ...costs // Merge cost data into the carrier object
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
            if (!isNaN(maxPallets)) {
                if (size.includes("36")) capacities[carrier].pallets36 = maxPallets;
                else if (size.includes("48")) capacities[carrier].pallets48 = maxPallets;
                else if (size.includes("53")) capacities[carrier].pallets53 = maxPallets;
            }
        }
    });
    return capacities;
}


function getShipments(productTypeMap) {
  const values = getSheetData('Shipments');
  if (!values || values.length <= 1) return [];
  const headers = values[0];
  const data = values.slice(1);

  const aggregatedShipments = {};
  const storeIndex = headers.indexOf('Store');
  const palletType1Index = headers.indexOf('PalletType1');
  const palletType2Index = headers.indexOf('PalletType2');
  const palletsIndex = headers.indexOf('Pallets');

  if ([storeIndex, palletType1Index, palletType2Index, palletsIndex].includes(-1)) {
    Logger.log("Error: Missing required headers in 'Shipments' sheet.");
    return [];
  }

  data.forEach(row => {
    if (String(row[palletType2Index] || '').toUpperCase() === 'FLS') return;

    const store = row[storeIndex];
    const palletType1 = row[palletType1Index];
    const pallets = parseFloat(row[palletsIndex]);

    if (!store || !palletType1 || isNaN(pallets)) return;

    let category = 'Ambient';
    const palletTypeLower = palletType1.toLowerCase();
    if (['chiller', 'meat'].includes(palletTypeLower)) category = 'Chiller';
    else if (['freezer', 'freezertkt'].includes(palletTypeLower)) category = 'Freezer';

    const productTypeId = productTypeMap[palletType1] || null;
    const key = `${store}-${category}`;

    if (aggregatedShipments[key]) {
      aggregatedShipments[key].pallets += pallets;
    } else {
      aggregatedShipments[key] = { id: key, store, category, pallets, productTypeId, attempts: 0, failureReason: null };
    }
  });

  const shipments = Object.values(aggregatedShipments);
  Logger.log(`Fetched and aggregated ${shipments.length} unique shipments.`);
  return shipments;
}

function getDistanceMatrix(shipments, addressData, warehouseLocation, apiKey) {
  const matrix = {};
  const sheetData = getSheetData('Distance Matrix');
  
  if (sheetData && sheetData.length > 1) {
    const headers = sheetData[0];
    const fromIndex = headers.indexOf('From');
    const toIndex = headers.indexOf('To');
    const distIndex = headers.indexOf('Distance (M)');
    const durIndex = headers.indexOf('Duration (Min)');

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

function getDetailedDurations() {
  const values = getSheetData('Adress Product Type Duration');
  if (!values || values.length <= 1) return {};
  const durations = {};
  const headers = values[0];
  const [storeIndex, typeIndex, loadIndex, unloadIndex] = ['Store Number', 'Product Type', 'LoadingTimePerFLS(Min)', 'UnloadingTimePerFLS(Mins)'].map(h => headers.indexOf(h));

  const parseTime = (timeValue) => {
    if (timeValue instanceof Date) return (timeValue.getHours() * 60) + timeValue.getMinutes();
    if (typeof timeValue === 'string') {
      const parts = timeValue.split(':').map(Number);
      if (parts.length >= 2) return (parts[0] * 60) + parts[1];
    }
    return 0;
  };

  values.slice(1).forEach(row => {
    const store = row[storeIndex], type = row[typeIndex];
    if (!store || !type) return;
    if (!durations[store]) durations[store] = {};
    durations[store][type] = {
      loading: parseTime(row[loadIndex]),
      unloading: parseTime(row[unloadIndex])
    };
  });
  Logger.log("Fetched detailed loading/unloading durations.");
  return durations;
}

function getAddressData() {
  const values = getSheetData('Address Data');
  if (!values || values.length <= 1) return {};
  const headers = values[0];
  const addresses = {};
  values.slice(1).forEach(row => {
    const storeNumber = row[headers.indexOf('Store Number')];
    if (storeNumber) {
      addresses[storeNumber] = {
        name: row[headers.indexOf('Name')],
        street: row[headers.indexOf('Street')],
        city: row[headers.indexOf('City')],
        zip: row[headers.indexOf('ZipCode')],
        cluster: row[headers.indexOf('Cluster')] || 'Others'
      };
    }
  });
  Logger.log('Fetched address data.');
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

function isTempCompatible(requiredCategories) {
  return !(requiredCategories.has('Chiller') && requiredCategories.has('Freezer'));
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
    const productTypeMap = getProductTypeMap();
    shipments = getShipments(productTypeMap);
    allCarriers = getCarriers();
    restrictions = getRestrictions();
    detailedDurations = getDetailedDurations();
    addressData = getAddressData();
    distanceMatrix = getDistanceMatrix(shipments, addressData, WAREHOUSE_LOCATION, API_KEY);
    shipments.forEach(s => s.cluster = (addressData[s.store] && addressData[s.store].cluster) || 'Others');
  } catch (e) {
    ui.alert('Error loading data. Check script logs.');
    Logger.log("Data Loading Error: " + e.toString() + e.stack);
    return;
  }

  const context = { ...settings, restrictions, detailedDurations, distanceMatrix, addressData, warehouse: WAREHOUSE_LOCATION, OVERPLAN_FACTOR: 1.0, MIN_PALLETS_PER_ROUTE: 15, MAX_PLANNING_ATTEMPTS: 3 };
  
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

  ui.alert('Success', 'Planning complete! Check the Dashboard sheet for analytics.', ui.ButtonSet.OK);
}

function runPlanningLoop(shipmentsToPlan, availableCarriers, context, planSheet, log, startIteration) {
    let unassignedShipments = [...shipmentsToPlan];
    let iteration = startIteration;
    const { MAX_PLANNING_ATTEMPTS, MIN_PALLETS_PER_ROUTE, restrictions } = context;

    let planningInProgress = true;
    while (planningInProgress) {
        unassignedShipments.sort((a, b) => b.pallets - a.pallets);
        
        let seedIndex = unassignedShipments.findIndex(s => s.attempts < MAX_PLANNING_ATTEMPTS && !s.failureReason);
        
        if (seedIndex === -1) {
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
        let routeChanged = true;
        while (routeChanged) {
            routeChanged = false;
            let bestCandidate = null, bestScore = Infinity;
            
            for (let i = unassignedShipments.length - 1; i >= 0; i--) {
                const candidate = unassignedShipments[i];
                if (candidate.failureReason) continue;
                if (currentRoute.cluster !== 'Others' && candidate.cluster !== currentRoute.cluster) continue;
                
                const score = checkAndScoreInsertion(candidate, currentRoute, context, log, iteration);
                if (score !== null && score < bestScore) {
                    bestScore = score;
                    bestCandidate = { shipment: candidate, index: i };
                }
            }

            if (bestCandidate) {
                const addedShipment = unassignedShipments.splice(bestCandidate.index, 1)[0];
                currentRoute.stops.push(addedShipment);
                currentRoute.totalPallets += addedShipment.pallets;
                if (currentRoute.cluster === 'Others' && addedShipment.cluster !== 'Others') {
                    currentRoute.cluster = addedShipment.cluster;
                }
                routeChanged = true;
            }
        }

        if (currentRoute.totalPallets < MIN_PALLETS_PER_ROUTE) {
            currentRoute.stops.forEach(stop => {
                unassignedShipments.push(stop);
            });
        } else {
            bestRouteOption.timeSlot.used++;
            currentRoute.routeId = `${currentRoute.carrier.name.replace(/\s+/g, '')}_${iteration++}`;
            writeRouteToSheet(currentRoute, planSheet, context);
        }
    }
    
    return { remainingShipments: unassignedShipments, iteration };
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
  const { restrictions, distanceMatrix, warehouse, OVERPLAN_FACTOR, MAX_ROUTE_MILEAGE, MAX_STOPS_GENERAL, MAX_DISTANCE_BETWEEN_STOPS } = context;

  const tempStops = [...route.stops, shipment];
  if ([...new Set(tempStops.map(s => s.store))].length > MAX_STOPS_GENERAL) return null;
  if (getDistanceFromFirstStop(tempStops, distanceMatrix) > MAX_DISTANCE_BETWEEN_STOPS) return null;
  if (!isTempCompatible(new Set(tempStops.map(s => s.category)))) return null;
  
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

  const tempOptimized = advancedOptimizeStopOrder(tempStops, warehouse, distanceMatrix);
  const tempMetrics = calculateRouteMetricsWithHOS(tempOptimized, warehouse, distanceMatrix, context.detailedDurations);
  if (tempMetrics.totalDistance > MAX_ROUTE_MILEAGE || tempMetrics.hosStatus !== 'OK') return null;

  const originalMetrics = calculateRouteMetricsWithHOS(advancedOptimizeStopOrder(route.stops, warehouse, distanceMatrix), warehouse, distanceMatrix, context.detailedDurations);
  
  const newCost = (tempMetrics.totalDistance * carrier.costPerMile) + carrier.costPerRoute;
  const oldCost = (originalMetrics.totalDistance * carrier.costPerMile) + carrier.costPerRoute;

  return newCost - oldCost; // Return the change in cost
}

function planRemainingShipments(shipments, context, planSheet) {
  const unplannableTime = shipments.filter(s => s.failureReason === 'Time Constraint');
  const unplannableOther = shipments.filter(s => s.failureReason && s.failureReason !== 'Time Constraint');
  const overspillShipments = shipments.filter(s => !s.failureReason);

  if (unplannableTime.length > 0) {
    writeRouteToSheet({ carrier: { name: 'Unplannable (Time)' }, routeId: 'Time_1', stops: unplannableTime, totalPallets: unplannableTime.reduce((s, p) => s + p.pallets, 0), cluster: 'Manual Review', notes: 'Failed due to delivery window constraints.' }, planSheet, context);
  }
  if (unplannableOther.length > 0) {
    writeRouteToSheet({ carrier: { name: 'Unplannable (Other)' }, routeId: 'Other_1', stops: unplannableOther, totalPallets: unplannableOther.reduce((s, p) => s + p.pallets, 0), cluster: 'Manual Review', notes: 'Failed planning after max attempts or no viable carrier.' }, planSheet, context);
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

function getDistanceFromFirstStop(stops, distanceMatrix) {
  if (stops.length < 2) return 0;
  let maxDist = 0;
  const firstStopId = stops[0].store;
  for (let i = 1; i < stops.length; i++) {
    const leg = (distanceMatrix[firstStopId] && distanceMatrix[firstStopId][stops[i].store]);
    if (leg && leg.distance > maxDist) maxDist = leg.distance;
  }
  return maxDist;
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

  return {
    travelTime: Math.round(travelTime),
    stopTime: Math.round(stopTime),
    totalDuration: Math.round(onDutyTime),
    totalDistance: Math.round(totalDistance),
    hosStatus: hosStatus
  };
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

  for (const shipment of shipments) {
    for (const timeSlot of carrier.timeSlots) {
      const tempRoute = { carrier, time: timeSlot.time };
      const failureReason = getInitialRouteFailureReason(shipment, tempRoute, context);
      if (failureReason === null) {
        return "NOTE: Carrier is valid but was not selected as the optimal choice for any available routes.";
      }
    }
  }

  const firstShipment = shipments[0];
  const firstTimeSlot = carrier.timeSlots[0];
  const representativeReason = getInitialRouteFailureReason(firstShipment, { carrier, time: firstTimeSlot.time }, context);
  
  return `FAIL: Rejected for all shipments. Example Reason: ${representativeReason}`;
}

function getInitialRouteFailureReason(shipment, route, context) {
    const { restrictions, distanceMatrix, warehouse, MAX_ROUTE_MILEAGE } = context;

    if (!distanceMatrix[warehouse] || !distanceMatrix[warehouse][shipment.store]) {
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
