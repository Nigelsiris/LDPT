# LDPT Routing Improvements Guide

## Summary of Recent Changes

### 1. Mileage Constraint Improvements (Implemented)
- **Reduced MAX_ROUTE_MILEAGE**: Changed from 500 to 425 miles
- **Added Mileage Penalty**: Routes exceeding 361 miles (85% of max) now receive exponential penalties
- **Target Mileage**: ~361 miles per route (well below 450 mile requirement)
- **Impact**: Routes over 425 miles will be rejected during planning, and routes over 361 miles will be heavily penalized in scoring

### 2. How Mileage is Calculated

The system calculates **round-trip mileage** as follows:
```
Total Distance = Warehouse → Stop1 → Stop2 → ... → StopN → Warehouse
```

Key factors affecting mileage:
- **Warehouse Location**: Distance from warehouse to first stop
- **Stop Optimization**: The `advancedOptimizeStopOrder` function uses nearest-neighbor algorithm
- **Return Trip**: The route must return to the warehouse

### 3. Current Mileage Constraints

| Constraint | Value | Purpose |
|------------|-------|---------|
| MAX_ROUTE_MILEAGE | 425 miles | Hard limit - routes exceeding this are rejected |
| Mileage Target | 361 miles | Soft target - routes exceeding this get penalties |
| MAX_DISTANCE_BETWEEN_STOPS | 60 miles | Preferred leg distance |
| ABSOLUTE_MAX_DISTANCE_BETWEEN_STOPS | 75 miles | Hard limit between consecutive stops |

## Potential Issues and Solutions

### Issue 1: Routes Still Exceeding Target Mileage

**Possible Causes:**
1. **Warehouse Location**: The warehouse (US0007) might be far from store clusters
2. **Sparse Store Distribution**: Stores might be geographically spread out
3. **Distance Matrix Accuracy**: Distances in the Distance Matrix sheet might be incorrect

**Solutions:**
1. **Verify Distance Matrix Data**:
   - Check the "Distance Matrix" sheet for accuracy
   - Ensure distances are in miles, not kilometers
   - Look for unusually large distances (>200 miles between any two points)

2. **Consider Multiple Warehouse Strategy**:
   - If stores are in multiple geographic regions, consider:
     - Creating separate planning runs for different regions
     - Adjusting the WAREHOUSE_LOCATION variable for different clusters

3. **Strengthen Clustering**:
   - The code already has cluster penalties (800 points)
   - Consider adding cluster data to the "Addresses" sheet if not already present

### Issue 2: Data Performance (Moving from Sheets to Code)

**Current Sheet Dependencies:**
```javascript
- Shipments: Dynamic data (changes frequently) - KEEP IN SHEET
- Carriers_2: Semi-static carrier capacity - COULD MOVE TO CODE
- Carrier Costs: Semi-static cost data - COULD MOVE TO CODE
- Carrier Inventory: Static pallet capacities - CANDIDATE FOR CODE
- ProductTypes: Static product definitions - GOOD CANDIDATE FOR CODE
- Restrictions: Semi-static delivery restrictions - COULD MOVE TO CODE
- Addresses: Semi-static location data - COULD MOVE TO CODE
- Distance Matrix: Calculated/cached distances - KEEP IN SHEET
- Adress Product Type Duration: Static duration data - GOOD CANDIDATE FOR CODE
```

**Recommended Approach:**
1. **Keep in Sheets** (changes frequently):
   - Shipments
   - Distance Matrix (acts as cache)
   - Carrier capacity adjustments

2. **Move to Code Constants** (rarely changes):
   - ProductTypes
   - Adress Product Type Duration
   - Carrier Inventory (pallet capacities)

3. **Hybrid Approach** (semi-static):
   - Keep in sheets for easy editing
   - Cache in script properties for faster access
   - Add "Refresh Cache" button in sidebar

**Example Implementation:**
```javascript
// Option 1: Hard-code in constants
const PRODUCT_TYPES = {
  'Ambient': 'AMB',
  'Chiller': 'CHL',
  'Freezer': 'FRZ',
  // ... more types
};

// Option 2: Cache with refresh
const CachedDataService = {
  get: function(key) {
    const cache = PropertiesService.getScriptProperties();
    const cached = cache.getProperty(key);
    if (cached) return JSON.parse(cached);
    return null;
  },
  set: function(key, value) {
    const cache = PropertiesService.getScriptProperties();
    cache.setProperty(key, JSON.stringify(value));
  },
  refresh: function() {
    // Reload from sheets and update cache
  }
};
```

## Machine Learning Feedback System (Future Enhancement)

**Concept:**
Add a feedback loop to learn from actual route performance and adjust planning parameters.

**Possible Implementation:**
1. **Data Collection Sheet** ("Route Performance"):
   ```
   Columns: Route ID, Planned Miles, Actual Miles, Planned Duration, 
            Actual Duration, Utilization, Issues, Date
   ```

2. **Learning Algorithm** (Simple):
   ```javascript
   function adjustParametersFromFeedback() {
     // Analyze historical performance
     const routes = getSheetData('Route Performance');
     const avgVariance = calculateMileageVariance(routes);
     
     // Adjust planning parameters
     if (avgVariance > 10%) {
       // Routes consistently over/under by 10%+
       settings.MAX_ROUTE_MILEAGE *= (1 - avgVariance/100);
     }
   }
   ```

3. **Automatic Tuning**:
   - Track utilization vs. mileage trade-offs
   - Identify optimal cluster sizes
   - Adjust penalties based on actual outcomes

**Advanced ML Approach:**
- Export route data to CSV
- Use external ML tools (Python/TensorFlow) for analysis
- Import optimized parameters back into settings

## Recommended Next Steps

1. **Immediate** (This Week):
   - [x] Reduce MAX_ROUTE_MILEAGE to 425 miles ✅
   - [x] Add mileage penalties to route scoring ✅
   - [ ] Review Distance Matrix sheet for accuracy
   - [ ] Run test planning cycle and verify route mileages

2. **Short Term** (Next 2 Weeks):
   - [ ] Move ProductTypes to code constants
   - [ ] Move Adress Product Type Duration to code constants
   - [ ] Add cluster verification to Addresses sheet
   - [ ] Implement cached data service for semi-static data

3. **Medium Term** (Next Month):
   - [ ] Create Route Performance tracking sheet
   - [ ] Implement basic feedback system
   - [ ] Add "Refresh Cache" functionality to sidebar
   - [ ] Consider multi-warehouse routing strategy

4. **Long Term** (2-3 Months):
   - [ ] Implement machine learning feedback loop
   - [ ] Add route simulation/preview before execution
   - [ ] Create performance dashboard with trends
   - [ ] Develop route comparison tools

## Testing the Changes

### Verification Steps:
1. Run `generatePlan()` with current data
2. Check the diagnostic report in logs:
   ```
   Logger: Routes Over Mileage: X (should be 0 or minimal)
   Logger: Avg Utilization: Y% (target: 92-94%)
   ```
3. Review the "Generated Plan" sheet:
   - Sort by "Total Route Mileage" column
   - Verify all routes are < 425 miles
   - Check "Mileage Status" column (should show "OK")

### Expected Results:
- Total routes: Should decrease slightly (consolidation)
- Avg utilization: Should increase to 90%+
- Max route mileage: Should be ≤ 425 miles
- Most routes: Should be in 320-380 mile range
- Routes over 425: Should be marked as "OVER" and appear minimal/zero

## Troubleshooting

### If routes still exceed 450 miles:
1. Check Distance Matrix data accuracy
2. Verify warehouse location is correct
3. Review store cluster assignments
4. Consider lowering MAX_ROUTE_MILEAGE further (to 400 or 375)

### If utilization drops below 85%:
1. The mileage penalty might be too aggressive
2. Consider adjusting MILEAGE_TARGET to 90% instead of 85%
3. Review MIN_PALLETS_PER_ROUTE setting

### If too many "Unplannable" routes:
1. The constraints might be too strict
2. Consider relaxing ABSOLUTE_MAX_DISTANCE_BETWEEN_STOPS to 80-85 miles
3. Review carrier capacity vs. shipment volume

## Contact & Support

For issues or questions:
1. Check the Logger output in Apps Script editor (View > Logs)
2. Review the diagnostic report at the end of plan generation
3. Examine the "Generated Plan" sheet for detailed route information
