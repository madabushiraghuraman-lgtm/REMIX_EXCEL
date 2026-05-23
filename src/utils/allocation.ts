/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { ExcelRowData, MslConfig, AllocationProposal, ActionStrategy } from '../types';

// Helper to normalize gender strings to match our categories
export function normalizeGender(genderStr: string = ''): 'boys' | 'girls' | 'kids' | 'ladies' | 'gents' {
  const s = genderStr.toLowerCase();
  if (s.includes('boy')) return 'boys';
  if (s.includes('girl')) return 'girls';
  if (s.includes('kid') || s.includes('child')) return 'kids';
  if (s.includes('lady') || s.includes('ladies') || s.includes('women') || s.includes('female')) return 'ladies';
  return 'gents'; // Default to gents
}

// Get the MSL quantity based on gender category and size
export function getMslValue(gender: string, size: string, config: MslConfig): number {
  const category = normalizeGender(gender);
  const catConfig = config[category];
  
  // Normalize size key (remove spaces, lowercase)
  const normSize = size.trim().toLowerCase();
  
  // Clean custom size keys (e.g. "size-7" or "7" or "6-1" or "6")
  // Let's search inside the keys of catConfig.sizes
  for (const sizeKey of Object.keys(catConfig.sizes)) {
    const cleanKey = sizeKey.trim().toLowerCase();
    if (cleanKey === normSize || normSize === cleanKey.split('-')[0]) {
      return catConfig.sizes[sizeKey];
    }
  }
  
  // Fallback to direct size match or standard default
  if (catConfig.sizes[size] !== undefined) {
    return catConfig.sizes[size];
  }
  
  return catConfig.defaultVal;
}

// Initialize default MSL configuration
export function getDefaultMslConfig(): MslConfig {
  return {
    boys: {
      defaultVal: 2,
      sizes: {}
    },
    girls: {
      defaultVal: 2,
      sizes: {}
    },
    kids: {
      defaultVal: 2,
      sizes: {}
    },
    ladies: {
      defaultVal: 2, // Default size val, override specific ones
      sizes: {
        '5': 1,
        '6': 1,
        '7': 2,
        '8': 2,
        '9': 2,
        '10': 2,
        '11': 1,
        '12': 1,
        '13': 1,
        '14': 1
      }
    },
    gents: {
      defaultVal: 2,
      sizes: {
        '6': 1,
        '7': 1,
        '8': 2,
        '9': 2,
        '10': 2,
        '11': 1,
        '12': 1,
        '13': 1,
        '14': 1
      }
    }
  };
}

/**
 * Perform Stock Allocation
 * 
 * @param excelData Parsed rows from Excel sheet
 * @param searchList List of target item codes/barcodes or item names (from textarea/notepad). If empty, run on all.
 * @param sourceStores Selected source branches
 * @param destinationStores Selected destination branches
 * @param mslConfig Current user-configured Minimum Stock Levels
 * @param defaultStrategy How to handle non-sales items by default ('automatic' | 'fresh_allocate' | 'ignore')
 */
export function runAllocationAlgorithm(
  excelData: ExcelRowData[],
  searchList: string[],
  sourceStores: string[],
  destinationStores: string[],
  mslConfig: MslConfig,
  defaultStrategy: ActionStrategy
): AllocationProposal[] {
  if (!excelData || excelData.length === 0) return [];
  if (sourceStores.length === 0 || destinationStores.length === 0) return [];

  const proposals: AllocationProposal[] = [];

  // 1. Filter rows by selected source and destination stores
  const sourceRows = excelData.filter(r => sourceStores.includes(r.BRANCH));
  const destRows = excelData.filter(r => destinationStores.includes(r.BRANCH));

  // 2. Identify the target item set
  // An item can be identified by its barcode or item name.
  // Let's find all unique (BARCODE, ITEM NAME, CATEGORY, STYLE, GENDER, BRAND, SIZE, MRP) in source rows first
  // that match our searchList (if searchList is not empty).
  const cleanSearchKeys = searchList
    .map(s => s.trim().toLowerCase())
    .filter(s => s.length > 0);

  // Filter items in the source pool that match our target search keys.
  // If search list is empty, we consider all items in source.
  const targetItemsMap = new Map<string, {
    barcode: string;
    itemName: string;
    gender: string;
    style: string;
    brand: string;
    size: string;
    mrp: number;
  }>();

  sourceRows.forEach(row => {
    const bcode = (row.BARCODE || '').toString().trim();
    const iname = (row["ITEM NAME"] || '').toString().trim();
    const size = (row.SIZE || '').toString().trim();
    
    // Check if matches search
    let isMatch = false;
    if (cleanSearchKeys.length === 0) {
      isMatch = true;
    } else {
      isMatch = cleanSearchKeys.some(key => {
        return bcode.toLowerCase() === key || 
               iname.toLowerCase().includes(key) || 
               (bcode && key.includes(bcode.toLowerCase()));
      });
    }

    if (isMatch && bcode) {
      const key = `${bcode}::${size}`;
      if (!targetItemsMap.has(key)) {
        targetItemsMap.set(key, {
          barcode: bcode,
          itemName: iname,
          gender: (row.GENDER || '').toString().trim(),
          style: (row.STYLE || '').toString().trim(),
          brand: (row.BRAND || '').toString().trim(),
          size: size,
          mrp: Number(row.MRP) || 0
        });
      }
    }
  });

  // Unique item entities (barcode + size grouping)
  const targetItems = Array.from(targetItemsMap.values());

  // 3. For each target item, run the prioritization and transfer simulation
  targetItems.forEach((item, itemIdx) => {
    // Collect source stock of this EXACT barcode & size, group/sum by branch to handle different lots
    const sourcesByBranch = new Map<string, { branch: string; closingStock: number; originalData: ExcelRowData }>();
    
    sourceRows
      .filter(r => (r.BARCODE || '').toString().trim() === item.barcode && (r.SIZE || '').toString().trim() === item.size)
      .forEach(r => {
        const branch = r.BRANCH;
        const stockQty = Number(r["Closing Stock QTY"]) || 0;
        if (stockQty > 0) {
          const existing = sourcesByBranch.get(branch);
          if (existing) {
            existing.closingStock += stockQty;
          } else {
            sourcesByBranch.set(branch, {
              branch,
              closingStock: stockQty,
              originalData: r
            });
          }
        }
      });

    const availableSources = Array.from(sourcesByBranch.values());

    // Sum of stock available from sources
    let totalSourceStock = availableSources.reduce((sum, s) => sum + s.closingStock, 0);
    if (totalSourceStock <= 0) return; // No items to transfer

    // Sort source stores descending by closing stock qty so we draw from the most slow-moving first
    availableSources.sort((a, b) => b.closingStock - a.closingStock);

    // Get the configured MSL for this item gender + size
    const targetMsl = getMslValue(item.gender, item.size, mslConfig);

    // For this item, find existing data (sales/stock/different lots) in each selected destination store
    // Even if no row exists in standard destination rows, we must represent it as a potential destination
    // with 0 Sales and 0 Stock.
    const destinationsForThisItem = destinationStores.map(destBranch => {
      // Find all rows in excel matching destination branch, barcode, and size
      const matchingRows = destRows.filter(
        r => r.BRANCH === destBranch && 
             (r.BARCODE || '').toString().trim() === item.barcode && 
             (r.SIZE || '').toString().trim() === item.size
      );

      const existingRow = matchingRows[0] || null;

      // Sum values across all lot rows in the destination branch
      const netSales = matchingRows.reduce((sum, r) => sum + (Number(r["NET SALE QTY"]) || 0), 0);
      const closingStock = matchingRows.reduce((sum, r) => sum + (Number(r["Closing Stock QTY"]) || 0), 0);

      // We determine if this branch is already carrying other sizes of this item name, which represents a "size gap" if current size is 0
      const sameItemRows = destRows.filter(
        r => r.BRANCH === destBranch && 
             (r["ITEM NAME"] || '').toString().trim().toLowerCase() === item.itemName.toLowerCase()
      );

      const sameItemOtherSizesRows = sameItemRows.filter(
        r => (r.SIZE || '').toString().trim() !== item.size
      );

      const totalStockOtherSizes = sameItemOtherSizesRows.reduce((sum, r) => sum + (Number(r["Closing Stock QTY"]) || 0), 0);
      const totalSalesOfItem = sameItemRows.reduce((sum, r) => sum + (Number(r["NET SALE QTY"]) || 0), 0);

      const hasSizeGap = (closingStock === 0) && (totalStockOtherSizes > 0);
      const isNewStore = (closingStock === 0) && (totalStockOtherSizes === 0);

      // Classify if it is "Fast-Moving" automatically:
      // "Fast moving item name store is store with greater then 1 qty of item name sales and with out closing stock qty"
      const isAutoEligible = netSales > 1 && closingStock <= 0;

      // Calculate priority score for sorting
      let priorityScore = 0;
      if (isAutoEligible) {
        priorityScore += 50000;
      }

      if (hasSizeGap) {
        if (netSales > 0 || totalSalesOfItem > 0) {
          priorityScore += 20000;
        } else {
          priorityScore += 10000;
        }
      } else if (closingStock > 0) {
        priorityScore += 5000;
      } else if (isNewStore) {
        if (netSales > 0 || totalSalesOfItem > 0) {
          priorityScore += 2000;
        } else {
          priorityScore += 1000;
        }
      }

      // Add a small tie-breaker for sales to prioritize branches with relative demand
      priorityScore += (netSales * 10) + (totalSalesOfItem / 100);

      return {
        branch: destBranch,
        netSales,
        closingStock,
        isAutoEligible,
        existingRow,
        hasSizeGap,
        isNewStore,
        totalStockOtherSizes,
        totalSalesOfItem,
        priorityScore
      };
    });

    // Sort destinations by priority:
    // 1. Size Gaps / Fast-moving first
    // 2. New stores are considered only after size-gap stores have been addressed
    destinationsForThisItem.sort((a, b) => {
      return b.priorityScore - a.priorityScore;
    });

    // Filter destination candidates by our allocation decisions
    destinationsForThisItem.forEach(dest => {
      if (totalSourceStock <= 0) return;

      // Determine allocation strategy for this destination candidate:
      let actionType: 'automatic' | 'fresh' | 'ignored' = 'ignored';
      let reason = '';

      const gapInfo = dest.hasSizeGap 
        ? `[Size Gap: has stock in other sizes of "${item.itemName}" (${dest.totalStockOtherSizes} qty) but 0 in size ${item.size}]`
        : dest.isNewStore
        ? `[New Store: no stock across any size of "${item.itemName}"]`
        : `[Active: has current stock ${dest.closingStock}]`;

      if (dest.isAutoEligible) {
        actionType = 'automatic';
        reason = `Auto-Eligible: ${gapInfo} Sales in ${dest.branch} are ${dest.netSales} and stock is 0.`;
      } else {
        // Here we apply the decision for no-previous-sales or non-auto cases.
        // It's either "Fresh Allocate", "Automatic" (which skips non-auto), or "Ignore" (which ignores non-auto).
        if (defaultStrategy === 'fresh_allocate') {
          actionType = 'fresh';
          reason = `Fresh Allocate: ${gapInfo} Force-allocated up to MSL as per user preset strategy.`;
        } else if (defaultStrategy === 'automatic') {
          // Automatic mode means we ONLY transfer if they are fast-moving (meaning they have sales > 1 and stock = 0).
          // Otherwise, we ignore them
          actionType = 'ignored';
          reason = `Ignored: ${gapInfo} Does not meet fast-moving criteria (sales > 1 & stock = 0).`;
        } else {
          actionType = 'ignored';
          reason = `Ignored: ${gapInfo} Excluded as per user ignore strategy.`;
        }
      }

      // If ignored, create an 'ignored' record with 0 allocated so the user can see it in tabular logs and choose to force-transfer it!
      if (actionType === 'ignored') {
        proposals.push({
          id: `${item.barcode}_${dest.branch}_${item.size}_${itemIdx}`,
          barcode: item.barcode,
          itemName: item.itemName,
          gender: item.gender,
          style: item.style,
          brand: item.brand,
          size: item.size,
          mrp: item.mrp,
          sourceStore: availableSources[0]?.branch || sourceStores[0],
          sourceStoreClosingQty: availableSources[0]?.closingStock || 0,
          destinationStore: dest.branch,
          destinationStoreClosingQty: dest.closingStock,
          destinationStoreSales: dest.netSales,
          allocatedQty: 0,
          allocationVal: 0,
          type: 'ignored',
          reason: reason
        });
        return;
      }

      // If we allocate: we target up to MSL. We check if they already have stock
      const targetAllocationQty = Math.max(0, targetMsl - dest.closingStock);
      if (targetAllocationQty <= 0) {
        proposals.push({
          id: `${item.barcode}_${dest.branch}_${item.size}_${itemIdx}`,
          barcode: item.barcode,
          itemName: item.itemName,
          gender: item.gender,
          style: item.style,
          brand: item.brand,
          size: item.size,
          mrp: item.mrp,
          sourceStore: availableSources[0]?.branch || sourceStores[0],
          sourceStoreClosingQty: availableSources[0]?.closingStock || 0,
          destinationStore: dest.branch,
          destinationStoreClosingQty: dest.closingStock,
          destinationStoreSales: dest.netSales,
          allocatedQty: 0,
          allocationVal: 0,
          type: 'ignored',
          reason: `No Transfer: Destination current stock (${dest.closingStock}) already matches or exceeds MSL (${targetMsl}).`
        });
        return;
      }

      const totalSourceStockBeforeDraw = availableSources.reduce((sum, s) => sum + s.closingStock, 0);

      // Collect this quantity from our available source pools step by step
      let remainingToAllocate = targetAllocationQty;
      let actualAllocated = 0;
      let usedSourceBranches: { branch: string; drawn: number }[] = [];

      for (const src of availableSources) {
        if (remainingToAllocate <= 0 || src.closingStock <= 0) break;

        const drawQty = Math.min(src.closingStock, remainingToAllocate);
        src.closingStock -= drawQty;
        totalSourceStock -= drawQty;
        remainingToAllocate -= drawQty;
        actualAllocated += drawQty;

        usedSourceBranches.push({
          branch: src.branch,
          drawn: drawQty
        });
      }

      if (actualAllocated > 0) {
        // Create an allocation proposal record
        // If drawn from multiple, we can list the primary draw source branch or create multiple entries.
        // Let's create an entry for each source drawer to make calculations precise, or sum them up and list primary helper!
        // To keep headers clear: S.No | STORE (dest branch) | GENDER | STYLE | ITEM CODE | COMPANY | ITEM NAME | SIZE | MRP | HEAD OFFICE (HO) Stock/Source Store | ...
        // We can list the source store names in the HO Stock/Source Store column!
        const sourceStoresDisplay = usedSourceBranches.map(b => `${b.branch} (-${b.drawn})`).join(', ');

        proposals.push({
          id: `${item.barcode}_${dest.branch}_${item.size}_${itemIdx}`,
          barcode: item.barcode,
          itemName: item.itemName,
          gender: item.gender,
          style: item.style,
          brand: item.brand,
          size: item.size,
          mrp: item.mrp,
          sourceStore: sourceStoresDisplay,
          sourceStoreClosingQty: totalSourceStockBeforeDraw, // represents total slow inventory available in sources BEFORE draw
          destinationStore: dest.branch,
          destinationStoreClosingQty: dest.closingStock,
          destinationStoreSales: dest.netSales,
          allocatedQty: actualAllocated,
          allocationVal: actualAllocated * item.mrp,
          type: actionType,
          reason: reason + ` Drew ${actualAllocated} inventory from: ${sourceStoresDisplay}.`
        });
      } else {
        proposals.push({
          id: `${item.barcode}_${dest.branch}_${item.size}_${itemIdx}`,
          barcode: item.barcode,
          itemName: item.itemName,
          gender: item.gender,
          style: item.style,
          brand: item.brand,
          size: item.size,
          mrp: item.mrp,
          sourceStore: availableSources[0]?.branch || sourceStores[0],
          sourceStoreClosingQty: availableSources[0]?.closingStock || 0,
          destinationStore: dest.branch,
          destinationStoreClosingQty: dest.closingStock,
          destinationStoreSales: dest.netSales,
          allocatedQty: 0,
          allocationVal: 0,
          type: 'ignored',
          reason: `Stock Depleted: Checked source stores but slow-moving inventory is fully exhausted.`
        });
      }
    });
  });

  return proposals;
}
