/**
 * [UPDATED] Fetches and processes shipment data from Shipments sheet.
 * Aggregates shipments by store and category, and filters out Standby products.
 * @param {object} productTypeMap - Map of product types from ProductTypes sheet
 * @returns {Array<object>} Array of shipment objects
 */
function getShipments(productTypeMap) {
  const values = getSheetData('Shipments');
  if (!values || values.length <= 1) return [];
  
  const headers = values[0];
  const storeIndex = headers.indexOf('Store');
  const palletType1Index = headers.indexOf('Pallet Type 1');
  const palletType2Index = headers.indexOf('Pallet Type 2');
  const palletsIndex = headers.indexOf('Pallets');
  const productStatusIndex = headers.indexOf('Status'); // For filtering Standby

  if ([storeIndex, palletType1Index, palletType2Index, palletsIndex].includes(-1)) {
    Logger.log('Error: Missing required columns in Shipments sheet');
    return [];
  }

  // Use an object to aggregate shipments by store-category
  const aggregatedShipments = {};

  values.slice(1).forEach(row => {
    const store = row[storeIndex];
    const palletType1 = row[palletType1Index];
    const palletType2 = row[palletType2Index];
    const pallets = parseFloat(row[palletsIndex]);
    const status = productStatusIndex !== -1 ? String(row[productStatusIndex] || '').toLowerCase() : '';

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
  Logger.log(`Loaded ${shipments.length} shipments from Shipments sheet.`);
  return shipments;
}