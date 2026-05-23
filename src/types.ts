/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

// Excel Row Data structure from the uploaded spreadsheet
export interface ExcelRowData {
  SNO?: number;
  BRANCH: string; // Store name
  BARCODE: string; // Item Code
  CATEGORY: string;
  STYLE: string;
  GENDER: string;
  BRAND: string; // This maps also to "COMPANY"
  "ITEM NAME": string;
  SIZE: string;
  MRP: number;
  "Opening Balance"?: number;
  "OPENING VALUE-2"?: number;
  Purchase?: number;
  "Misc Repts"?: number;
  "Transfer in"?: number;
  "Pur Return"?: number;
  "NET SALE QTY": number; // Sales amount
  "MRP VALUE"?: number;
  "BILL VALUE"?: number;
  "Transfer Out"?: number;
  ISSUE?: number;
  "Closing Stock QTY": number; // Head Office or Source store stock
  "Value(MRP)"?: number;
  "LOT NUMBER"?: string;
  "LOT CODE"?: string;
  "LOT SUPPLIER NAME"?: string;
  "ORGINAL LOT CODE"?: string;
}

// MSL Configuration Map
export interface MslSizeMap {
  [size: string]: number;
}

export interface MslCategoryConfig {
  defaultVal: number;
  sizes: MslSizeMap;
}

export interface MslConfig {
  boys: MslCategoryConfig;
  girls: MslCategoryConfig;
  kids: MslCategoryConfig;
  ladies: MslCategoryConfig;
  gents: MslCategoryConfig;
}

export type ActionStrategy = 'automatic' | 'fresh_allocate' | 'ignore';

// Single proposed item transfer record for review and reporting
export interface AllocationProposal {
  id: string; // Unique combination key
  barcode: string;
  itemName: string;
  gender: string;
  style: string;
  brand: string; // maps to company
  size: string;
  mrp: number;
  sourceStore: string;
  sourceStoreClosingQty: number; // Current HO stock
  destinationStore: string;
  destinationStoreClosingQty: number; // Current destination stock
  destinationStoreSales: number; // Net Sale QTY in destination
  allocatedQty: number;
  allocationVal: number;
  type: 'automatic' | 'fresh' | 'ignored';
  reason: string;
}
