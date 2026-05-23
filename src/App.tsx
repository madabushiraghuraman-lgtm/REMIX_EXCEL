/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Settings,
  RefreshCw,
  FileSpreadsheet,
  FileCode,
  Trash2,
  Play,
  Download,
  AlertCircle,
  Filter,
  Check,
  Search,
  Building2,
  Sparkles,
  Clock,
  ArrowRight,
  Database,
  ListFilter,
  Zap,
  Ban,
  PlusCircle,
  TrendingUp,
  AlertTriangle,
  Layers,
  ChevronRight,
  Calculator,
  Sliders,
  Sparkle
} from 'lucide-react';

import { ExcelRowData, MslConfig, AllocationProposal, ActionStrategy } from './types';
import { runAllocationAlgorithm, getDefaultMslConfig, getMslValue } from './utils/allocation';
import MslEditor from './components/MslEditor';
import { NeonCard, NeonBadge, NeonButton, FileDropzone } from './components/NeonWidgets';

export default function App() {
  // Application Data States
  const [parsedData, setParsedData] = useState<ExcelRowData[]>([]);
  const [dataFileName, setDataFileName] = useState<string>('');
  const [searchItemsText, setSearchItemsText] = useState<string>('');
  const [searchItemsList, setSearchItemsList] = useState<string[]>([]);
  const [notepadFileName, setNotepadFileName] = useState<string>('');

  // Store Selection States
  const [sourceStores, setSourceStores] = useState<string[]>([]);
  const [destinationStores, setDestinationStores] = useState<string[]>([]);
  const [sourceSearchTerm, setSourceSearchTerm] = useState<string>('');
  const [destSearchTerm, setDestSearchTerm] = useState<string>('');

  // Custom Allocation Parameters
  const [mslConfig, setMslConfig] = useState<MslConfig>(getDefaultMslConfig());
  const [defaultStrategy, setDefaultStrategy] = useState<ActionStrategy>('automatic');

  // Outputs
  const [proposals, setProposals] = useState<AllocationProposal[]>([]);
  const [manualOverrides, setManualOverrides] = useState<Map<string, { type: 'automatic' | 'fresh' | 'ignored'; qty: number }>>(new Map());
  const [hasCalculated, setHasCalculated] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [timeStr, setTimeStr] = useState<string>('');

  // Local Table Filtering
  const [tableSearchTerm, setTableSearchTerm] = useState<string>('');
  const [tableStatusFilter, setTableStatusFilter] = useState<'all' | 'allocated' | 'ignored'>('all');

  // Tick simulated engine actions in terminal
  const [systemLogs, setSystemLogs] = useState<string[]>([
    '⚡ Cyberspace Allocation Core online v4.1.0',
    '📊 Idle - Awaiting Excel Inventory upload and parameters configuration'
  ]);

  // Clock runner
  useEffect(() => {
    const tick = () => {
      const pad = (n: number) => n.toString().padStart(2, '0');
      const now = new Date();
      setTimeStr(`${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`);
    };
    tick();
    const interval = setInterval(tick, 1000);
    return () => clearInterval(interval);
  }, []);

  // Sync manual textarea content into list array
  const parseTextareaInput = (text: string) => {
    const codes = text
      .split(/[\n,]+/)
      .map(line => line.trim())
      .filter(line => line.length > 0);
    setSearchItemsList(codes);
  };

  const handleTextareaChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const val = e.target.value;
    setSearchItemsText(val);
    parseTextareaInput(val);
  };

  const clearNotepadInput = () => {
    setSearchItemsText('');
    setSearchItemsList([]);
    setNotepadFileName('');
    addLog('🗑️ Notepad search criteria cache cleared.');
  };

  const addLog = (msg: string) => {
    const time = new Date().toLocaleTimeString();
    setSystemLogs(prev => [`[${time}] ${msg}`, ...prev.slice(0, 49)]);
  };

  // Parse Excel File on drop/select
  const handleExcelUpload = (file: File) => {
    setIsLoading(true);
    addLog(`📂 Commencing ingest of Excel file: ${file.name}`);
    
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const bstr = e.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        
        // Convert to JSON
        const rawJson: any[] = XLSX.utils.sheet_to_json(ws);
        
        if (rawJson.length === 0) {
          throw new Error('No dataset lines discovered in spreadsheet sheet 1.');
        }

        addLog(`📦 Parsed ${rawJson.length} raw records from spreadsheet.`);

        // Normalize rows to standard schema
        const normalized: ExcelRowData[] = rawJson.map((row, idx) => {
          // Normalize case-insensitive keys
          const fetch = (keys: string[], fallback: any = '') => {
            for (const key of Object.keys(row)) {
              if (keys.map(k => k.toLowerCase().replace(/[\s\._-]/g, '')).includes(key.toLowerCase().replace(/[\s\._-]/g, ''))) {
                return row[key];
              }
            }
            return fallback;
          };

          return {
            SNO: Number(fetch(['sno', 'sno.', 's.no', 'sr', 'srno'], idx + 1)),
            BRANCH: String(fetch(['branch', 'store', 'storename', 'branchname', 'location', 'outlet'], 'HO_STOCK')).trim(),
            BARCODE: String(fetch(['barcode', 'itemcode', 'code', 'id', 'barcodecode', 'sku'], '')).trim(),
            CATEGORY: String(fetch(['category', 'dept', 'department', 'cat'], '-')).trim(),
            STYLE: String(fetch(['style', 'design', 'stylecode', 'designcode'], '-')).trim(),
            GENDER: String(fetch(['gender', 'sex', 'gendergroup'], 'Gents')).trim(),
            BRAND: String(fetch(['brand', 'company', 'supplier', 'lot supplier', 'lot supplier name'], '-')).trim(),
            "ITEM NAME": String(fetch(['item name', 'itemname', 'description', 'name', 'desc'], '-')).trim(),
            SIZE: String(fetch(['size', 'size_name'], 'U')).trim(),
            MRP: Number(fetch(['mrp', 'price', 'retailprice', 'm.r.p.'], 0)),
            "NET SALE QTY": Number(fetch(['net sale qty', 'netsaleqty', 'netsales', 'salesqty', 'sales'], 0)),
            "Closing Stock QTY": Number(fetch(['closing stock qty', 'closingstockqty', 'closingstock', 'stockqty', 'currentstock', 'stock'], 0)),
          };
        });

        // Filter out empty rows of barcode
        const cleanNormalized = normalized.filter(row => row.BARCODE && row.BRANCH);

        setParsedData(cleanNormalized);
        setDataFileName(file.name);
        addLog(`✅ Successfully loaded ${cleanNormalized.length} qualified matrix entries.`);

        // Automatically assign default stores based on parsed data to save time!
        const storeSet = new Set(cleanNormalized.map(r => r.BRANCH));
        const stores = Array.from(storeSet).sort();
        
        // Pre-select Source Store as typically stores with HO / higher inventories or user selects
        // Here we default to setting all discovered stores available. No forced pre-selection so user feels fully in control.
        addLog(`🏭 Identified ${stores.length} unique branch entities across the dataset.`);
      } catch (err: any) {
        alert(`Failed to parse Excel: ${err.message || err}`);
        addLog(`❌ Critial error parsing Excel sheet: ${err.message}`);
      } finally {
        setIsLoading(false);
      }
    };
    reader.onerror = () => {
      addLog(`❌ File reader IO interrupt.`);
      setIsLoading(false);
    };
    reader.readAsBinaryString(file);
  };

  // Parse Notepad file of Item codes
  const handleNotepadUpload = (file: File) => {
    setIsLoading(true);
    addLog(`📄 Uploading notepad query criteria text file: ${file.name}`);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const text = e.target?.result as string;
        setSearchItemsText(text);
        setNotepadFileName(file.name);
        
        const codes = text
          .split(/[\r\n,]+/)
          .map(line => line.trim())
          .filter(line => line.length > 0);
        
        setSearchItemsList(codes);
        addLog(`✅ Successfully loaded ${codes.length} barcode/item filter tokens from notepad.`);
      } catch (err: any) {
        addLog(`❌ Error loading notepad: ${err.message}`);
      } finally {
        setIsLoading(false);
      }
    };
    reader.readAsText(file);
  };

  // Unique list of store branches
  const uniqueBranches = useMemo(() => {
    const storeSet = new Set(parsedData.map(r => r.BRANCH));
    return Array.from(storeSet).sort();
  }, [parsedData]);

  // Compute stats of parsed excel data on the fly
  const dataSummary = useMemo(() => {
    if (parsedData.length === 0) return { totalRows: 0, totalStock: 0, totalSales: 0, valMRP: 0 };
    let totalStock = 0;
    let totalSales = 0;
    let valMRP = 0;
    parsedData.forEach(row => {
      totalStock += row["Closing Stock QTY"] || 0;
      totalSales += row["NET SALE QTY"] || 0;
      valMRP += (row["Closing Stock QTY"] || 0) * (row.MRP || 0);
    });
    return {
      totalRows: parsedData.length,
      totalStock,
      totalSales,
      valMRP
    };
  }, [parsedData]);

  // Branch multi-selection helpers
  const handleToggleSourceStore = (branch: string) => {
    setSourceStores(prev =>
      prev.includes(branch) ? prev.filter(b => b !== branch) : [...prev, branch]
    );
  };

  const handleToggleDestStore = (branch: string) => {
    setDestinationStores(prev =>
      prev.includes(branch) ? prev.filter(b => b !== branch) : [...prev, branch]
    );
  };

  const handleSelectAllSources = () => {
    const filtered = uniqueBranches.filter(b => b.toLowerCase().includes(sourceSearchTerm.toLowerCase()));
    setSourceStores(prev => {
      const next = new Set([...prev, ...filtered]);
      return Array.from(next);
    });
    addLog(`📢 Allocated all filtered stores as active SOURCE repositories.`);
  };

  const handleDeselectAllSources = () => {
    const filtered = uniqueBranches.filter(b => b.toLowerCase().includes(sourceSearchTerm.toLowerCase()));
    setSourceStores(prev => prev.filter(b => !filtered.includes(b)));
    addLog(`📢 Cleared custom filters for SOURCE repositories.`);
  };

  const handleSelectAllDests = () => {
    const filtered = uniqueBranches.filter(b => b.toLowerCase().includes(destSearchTerm.toLowerCase()));
    setDestinationStores(prev => {
      const next = new Set([...prev, ...filtered]);
      return Array.from(next);
    });
    addLog(`📢 Allocated all filtered stores as active RECEIVING destinations.`);
  };

  const handleDeselectAllDests = () => {
    const filtered = uniqueBranches.filter(b => b.toLowerCase().includes(destSearchTerm.toLowerCase()));
    setDestinationStores(prev => prev.filter(b => !filtered.includes(b)));
    addLog(`📢 Cleared custom filters for RECEIVING destinations.`);
  };

  // Perform core stock allocation calculations
  const handleTriggerAllocation = () => {
    if (parsedData.length === 0) {
      alert("Please upload the Inventory excelsheet first.");
      return;
    }
    if (sourceStores.length === 0) {
      alert("Please select at least one Source store.");
      return;
    }
    if (destinationStores.length === 0) {
      alert("Please select at least one Destination store.");
      return;
    }

    setIsLoading(true);
    addLog(`🚀 Initiating transfer calculation. Strategy mode: ${defaultStrategy.toUpperCase()}`);
    
    // Simulate real calculations in state
    setTimeout(() => {
      try {
        const computedProposals = runAllocationAlgorithm(
          parsedData,
          searchItemsList,
          sourceStores,
          destinationStores,
          mslConfig,
          defaultStrategy
        );

        setProposals(computedProposals);
        setManualOverrides(new Map()); // Reset manual revisions
        setHasCalculated(true);
        addLog(`📊 Allocation computed! Generated ${computedProposals.length} proposed trace paths.`);
      } catch (err: any) {
        addLog(`❌ Allocation calculation crash: ${err.message}`);
      } finally {
        setIsLoading(false);
      }
    }, 450);
  };

  // Reset MSL to hardcoded values
  const handleResetMsl = () => {
    setMslConfig(getDefaultMslConfig());
    addLog(`⚡ Reset MSL variables back to fixed standard: boys/girls/kids 2 each; Ladies/Gents size ranges matched.`);
  };

  // Apply row level overrides directly
  const handleRowTypeOverride = (id: string, nextType: 'automatic' | 'fresh' | 'ignored') => {
    const target = proposals.find(p => p.id === id);
    if (!target) return;

    // We recalculate quantity allocation based on decision
    let calculatedQty = 0;
    if (nextType !== 'ignored') {
      const targetMslVal = getMslValue(target.gender, target.size, mslConfig);
      // Allocate MSL - current stock
      calculatedQty = Math.max(0, targetMslVal - target.destinationStoreClosingQty);

      // Make sure we have source inventory
      // We look at original source inventory, but for dynamic row-specific click overrides,
      // let's restore original source closing qty if it was depleted, or allow the quantity override
      if (calculatedQty === 0) {
        calculatedQty = targetMslVal; // force allocate 
      }
    }

    const nextOverrides = new Map(manualOverrides);
    nextOverrides.set(id, { type: nextType, qty: calculatedQty });
    setManualOverrides(nextOverrides);
    addLog(`✏️ Manual override row ${target.barcode} size ${target.size} dest ${target.destinationStore} set to ${nextType.toUpperCase()} (Qty: ${calculatedQty})`);
  };

  const handleRowQtyChange = (id: string, qty: number) => {
    const target = proposals.find(p => p.id === id);
    if (!target) return;

    const currentOverride = manualOverrides.get(id) || { type: target.type, qty: target.allocatedQty };
    const nextOverrides = new Map(manualOverrides);
    nextOverrides.set(id, {
      type: qty > 0 ? (currentOverride.type === 'ignored' ? 'fresh' : currentOverride.type) : 'ignored',
      qty: Math.max(0, qty)
    });
    setManualOverrides(nextOverrides);
  };

  // Complete consolidated proposals listing manually edited values on the fly
  const finalizedProposals = useMemo(() => {
    return proposals.map(prop => {
      const override = manualOverrides.get(prop.id);
      if (override) {
        return {
          ...prop,
          type: override.type,
          allocatedQty: override.qty,
          allocationVal: override.qty * prop.mrp
        };
      }
      return prop;
    });
  }, [proposals, manualOverrides]);

  // Computed results metrics
  const allocationSummary = useMemo(() => {
    const filtered = finalizedProposals.filter(p => p.allocatedQty > 0);
    const sumQty = filtered.reduce((sum, p) => sum + p.allocatedQty, 0);
    const sumVal = filtered.reduce((sum, p) => sum + p.allocationVal, 0);
    const distinctBarcodes = new Set(filtered.map(p => p.barcode)).size;
    const targetsStores = new Set(filtered.map(p => p.destinationStore)).size;
    return {
      activeTransferRows: filtered.length,
      sumQty,
      sumVal,
      distinctBarcodes,
      targetsStores
    };
  }, [finalizedProposals]);

  // Filtering table results
  const filteredAndSortedProposals = useMemo(() => {
    return finalizedProposals.filter(prop => {
      // Filter status
      if (tableStatusFilter === 'allocated' && prop.allocatedQty === 0) return false;
      if (tableStatusFilter === 'ignored' && prop.allocatedQty > 0) return false;

      // Filter text
      if (!tableSearchTerm) return true;
      const term = tableSearchTerm.toLowerCase();
      return prop.barcode.toLowerCase().includes(term) ||
             prop.itemName.toLowerCase().includes(term) ||
             prop.destinationStore.toLowerCase().includes(term) ||
             prop.gender.toLowerCase().includes(term) ||
             (prop.style && prop.style.toLowerCase().includes(term));
    });
  }, [finalizedProposals, tableStatusFilter, tableSearchTerm]);

  // Compile final allocation dataset and download as Excel (.xlsx) file
  const handleDownloadExcelReport = () => {
    const exportRows = finalizedProposals
      .filter(prop => prop.allocatedQty > 0)
      .map((p, idx) => {
        return {
          "S.No": idx + 1,
          "STORE": p.destinationStore,
          "GENDER": p.gender,
          "STYLE": p.style,
          "ITEM CODE": p.barcode,
          "COMPANY": p.brand || "N/A",
          "ITEM NAME": p.itemName,
          "SIZE": p.size,
          "MRP": p.mrp,
          "SOURCE QTY": p.sourceStoreClosingQty,
          "Destination store Current Closing Stock": p.destinationStoreClosingQty,
          "destination store Net Sale Qty": p.destinationStoreSales,
          "Allocated QTY": p.allocatedQty,
          "Allocation VAL": p.allocationVal,
          "HEAD OFFICE (HO) Stock/Source Store": p.sourceStore
        };
      });

    if (exportRows.length === 0) {
      alert("No active allocations to download! Add items, configure stores, and run allocation.");
      return;
    }

    addLog(`📥 Generating final transfer report workbook with ${exportRows.length} allocation rows...`);

    try {
      // Core sheet config using XLSX
      const ws = XLSX.utils.json_to_sheet(exportRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Transfers_Report");
      
      // Save triggers system download dialog
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const outputFilename = `Retail_Transfers_Report_${timestamp}.xlsx`;
      
      addLog(`💾 Invoking browser pipeline to save file: ${outputFilename}`);
      XLSX.writeFile(wb, outputFilename);
      
      addLog(`✨ Report download completed perfectly! Check your system downloads directory.`);
    } catch (err: any) {
      addLog(`❌ Failed writing workbook file downloads: ${err.message}`);
      alert(`Error writing Excel report: ${err.message}`);
    }
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 font-sans grid-bg p-4 md:p-8 selection:bg-rose-500 selection:text-white pb-36">
      
      {/* GLOW DECORATIONS */}
      <div className="pointer-events-none fixed -top-40 left-1/4 h-[500px] w-[500px] rounded-full bg-cyan-500/5 blur-[120px] glow-pulser-cyan" />
      <div className="pointer-events-none fixed -bottom-40 right-1/4 h-[500px] w-[500px] rounded-full bg-rose-500/5 blur-[120px]" />

      {/* RETAIL WORKSPACE CONTAINER */}
      <div className="mx-auto max-w-7xl">
        
        {/* APP TITLE / HEADER */}
        <header className="relative mb-6 flex flex-wrap items-center justify-between gap-6 border-b border-cyan-500/20 bg-slate-900/60 p-6 rounded-2xl glow-cyan backdrop-blur-xl">
          <div className="flex items-center gap-4">
            <div className="relative flex h-14 w-14 items-center justify-center rounded-2xl bg-cyan-500/10 border border-cyan-400 shadow-[0_0_15px_rgba(6,182,212,0.3)] animate-pulse">
              <Database className="h-7 w-7 text-cyan-400" />
              <div className="absolute -top-1 -right-1 h-3 w-3 rounded-full bg-rose-500 animate-ping" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-mono tracking-widest text-cyan-400 font-bold uppercase p-0.5 bg-cyan-950/40 border border-cyan-500/20 rounded">
                  ENGINE CORE V4.1
                </span>
                <span className="flex h-2 w-2 rounded-full bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.8)]" />
                <span className="text-[10px] text-slate-400 font-mono">LIVE MATRIX PIPELINE</span>
              </div>
              <h1 className="font-sans text-2xl font-bold tracking-tight text-white uppercase mt-0.5">
                Excel Stock Transfer & <span className="bg-gradient-to-r from-cyan-400 via-rose-400 to-emerald-400 bg-clip-text text-transparent">Allocation Portal</span>
              </h1>
            </div>
          </div>

          {/* DYNAMIC DIGITAL CLOCK TOCK */}
          <div className="flex flex-col items-end gap-1 font-mono text-xs">
            <div className="flex items-center gap-1.5 bg-slate-950/60 px-3.5 py-1.5 rounded-lg border border-slate-800">
              <Clock className="h-3.5 w-3.5 text-rose-500 animate-spin-slow" />
              <span className="text-[10px] text-slate-400 font-bold tracking-widest">SYSTEM DATE:</span>
              <span className="text-rose-400 font-bold tracking-widest">{timeStr || '08:03:52'}</span>
            </div>
            <span className="text-[9px] text-slate-500 font-mono">UTC: 2026-05-20 08:03:52</span>
          </div>
        </header>

        {/* METRICS AND STATS TICKER */}
        <section className="mb-6 grid grid-cols-2 md:grid-cols-4 gap-4">
          <div className="rounded-xl border border-slate-800/80 bg-slate-900/40 p-4 transition-all hover:border-cyan-500/30">
            <div className="flex items-center justify-between">
              <span className="text-[10px] uppercase font-bold tracking-widest text-slate-400">Total Matrix Records</span>
              <FileSpreadsheet className="h-4 w-4 text-cyan-500" />
            </div>
            <div className="text-xl font-mono font-bold text-white mt-1 animate-fade-in">
              {dataSummary.totalRows.toLocaleString()} <span className="text-xs text-slate-500">rows</span>
            </div>
            <p className="text-[10px] text-slate-500 mt-1 truncate">
              {dataFileName ? `Source: ${dataFileName}` : 'Awaiting sheet upload...'}
            </p>
          </div>

          <div className="rounded-xl border border-slate-800/80 bg-slate-900/40 p-4 transition-all hover:border-rose-500/30">
            <div className="flex items-center justify-between">
              <span className="text-[10px] uppercase font-bold tracking-widest text-slate-400">Inventory Stock Count</span>
              <Database className="h-4 w-4 text-rose-500" />
            </div>
            <div className="text-xl font-mono font-bold text-white mt-1">
              {dataSummary.totalStock.toLocaleString()} <span className="text-xs text-slate-500">pcs</span>
            </div>
            <p className="text-[10px] text-slate-500 mt-1 font-mono">
              Value: {new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 }).format(dataSummary.valMRP)}
            </p>
          </div>

          <div className="rounded-xl border border-slate-800/80 bg-slate-900/40 p-4 transition-all hover:border-emerald-500/30">
            <div className="flex items-center justify-between">
              <span className="text-[10px] uppercase font-bold tracking-widest text-slate-400">Target Filters Count</span>
              <ListFilter className="h-4 w-4 text-emerald-500" />
            </div>
            <div className="text-xl font-mono font-bold text-white mt-1">
              {searchItemsList.length > 0 ? (
                <>
                  {searchItemsList.length} <span className="text-xs text-slate-500">items</span>
                </>
              ) : (
                <span className="text-emerald-400 text-sm">ALL SHEET ITEMS</span>
              )}
            </div>
            <p className="text-[10px] text-slate-500 mt-1 truncate">
              {notepadFileName ? `Notepad: ${notepadFileName}` : 'Manual keywords or Notepad upload.'}
            </p>
          </div>

          <div className="rounded-xl border border-slate-800/80 bg-slate-900/40 p-4 transition-all hover:border-amber-500/30">
            <div className="flex items-center justify-between">
              <span className="text-[10px] uppercase font-bold tracking-widest text-slate-400">Stores Discovered</span>
              <Building2 className="h-4 w-4 text-amber-500" />
            </div>
            <div className="text-xl font-mono font-bold text-white mt-1">
              {uniqueBranches.length} <span className="text-xs text-slate-400">stores</span>
            </div>
            <p className="text-[10px] text-slate-500 mt-1">
              {sourceStores.length} Selected Src | {destinationStores.length} Selected Dest
            </p>
          </div>
        </section>

        {/* WORKSTATION GRID */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
          
          {/* LEFT COMMAND PANEL: COMPILATION ELEMENTS AND UPLOADS */}
          <div className="lg:col-span-12 xl:col-span-5 flex flex-col gap-6">

            {/* NEON STEP 1 CARD: FILE PARSERS */}
            <NeonCard color="cyan" className="shadow-[0_0_15px_rgba(6,182,212,0.1)]">
              <div className="mb-4 flex items-center justify-between">
                <span className="text-xs font-mono font-bold tracking-widest text-cyan-400">STEP 01</span>
                <span className="text-xs bg-cyan-950/60 font-medium border border-cyan-500/20 text-cyan-300 px-2 py-0.5 rounded">
                  DATA INGESTION
                </span>
              </div>
              <h2 className="text-lg font-bold text-white uppercase mb-1">
                Upload Master Sheets
              </h2>
              <p className="text-slate-400 text-xs mb-5">
                Drop your raw Retail Excel file with standard attributes. It parses unlimited records, barcodes, genders, sizes, stock levels & billing logs.
              </p>

              {/* DUAL FILE DROPZONES */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <FileDropzone
                  onFileSelect={handleExcelUpload}
                  accept=".xlsx, .xls, .csv"
                  label="Master Excel Sheet"
                  subLabel="Excel .xlsx with branch stock and sales"
                  isLoaded={parsedData.length > 0}
                  loadedInfo={`${parsedData.length.toLocaleString()} rows verified`}
                  color="cyan"
                  iconType="spreadsheet"
                />

                <FileDropzone
                  onFileSelect={handleNotepadUpload}
                  accept=".txt"
                  label="Notepad Item codes"
                  subLabel="Notepad (.txt) with target barcodes"
                  isLoaded={searchItemsList.length > 0 && !!notepadFileName}
                  loadedInfo={`${searchItemsList.length} barcodes loaded`}
                  color="green"
                  iconType="code"
                />
              </div>

              {/* MANUAL KEYWORD FORM FIELD */}
              <div className="mt-5 border-t border-slate-800 pt-4">
                <div className="flex items-center justify-between mb-2">
                  <label className="text-xs text-slate-300 font-bold uppercase tracking-wider block">
                    ⚡ Paste Target Codes or Barcodes (Alternative)
                  </label>
                  {searchItemsList.length > 0 && (
                    <button
                      onClick={clearNotepadInput}
                      className="text-slate-500 hover:text-rose-400 transition-colors flex items-center gap-1 text-[11px]"
                    >
                      <Trash2 className="h-3 w-3" /> Clear filters
                    </button>
                  )}
                </div>

                <textarea
                  className="w-full h-24 bg-slate-950 text-emerald-400 placeholder:text-slate-600 font-mono text-xs rounded-xl border border-slate-800 p-3 focus:border-cyan-400 outline-none focus:shadow-[0_0_10px_rgba(6,182,212,0.2)]"
                  placeholder="Paste barcodes or partial item names here, one per line. We will search the Master sheet for matching criteria. Leave dry to match ALL data."
                  value={searchItemsText}
                  onChange={handleTextareaChange}
                />
                
                <div className="mt-2 flex items-center justify-between">
                  <span className="text-[10px] text-slate-500 font-mono">
                    💡 Parsed {searchItemsList.length} search tokens.
                  </span>
                  {searchItemsList.length > 0 && (
                    <span className="bg-emerald-950 text-emerald-400 border border-emerald-500/20 px-2 py-0.5 text-[9px] rounded font-bold font-mono">
                      Active: {searchItemsList.length} matches
                    </span>
                  )}
                </div>
              </div>
            </NeonCard>

            {/* NEON STEP 2 CARD: SOURCE STORES selection */}
            <NeonCard color="amber" className="shadow-[0_0_15px_rgba(245,158,11,0.1)]">
              <div className="mb-3 flex items-center justify-between">
                <span className="text-xs font-mono font-bold tracking-widest text-amber-500">STEP 02</span>
                <span className="text-xs bg-amber-950/60 font-medium border border-amber-500/20 text-amber-300 px-2 py-0.5 rounded">
                  STORES DIRECTORY
                </span>
              </div>
              <h2 className="text-lg font-bold text-white uppercase mb-1">
                Select Source & Destination
              </h2>
              <p className="text-slate-400 text-xs mb-4">
                <strong className="text-amber-400">Sources:</strong> SLOW moving stores (stock &gt; 0). <strong className="text-cyan-400">Destinations:</strong> FAST moving target stores (sales &gt; 1 & zero stock). Select multiple of each.
              </p>

              {parsedData.length === 0 ? (
                <div className="text-center py-10 bg-slate-950/50 rounded-xl border border-slate-800 text-xs text-slate-500 italic">
                  Upload inventory Excel data above to scan the store directories.
                </div>
              ) : (
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  
                  {/* SOURCE SELECTION WINDOW */}
                  <div className="bg-slate-950/50 border border-slate-850 rounded-xl p-3">
                    <div className="flex items-center justify-between mb-2 pb-1.5 border-b border-slate-900">
                      <span className="text-[11px] font-bold text-amber-400 uppercase tracking-widest block">
                        Source Stores ({sourceStores.length})
                      </span>
                    </div>

                    <input
                      type="text"
                      placeholder="Filter source store..."
                      className="w-full text-[11px] bg-slate-900 border border-slate-800 rounded px-2 py-1 mb-2 text-white outline-none focus:border-amber-500"
                      value={sourceSearchTerm}
                      onChange={(e) => setSourceSearchTerm(e.target.value)}
                    />

                    <div className="flex gap-1.5 mb-2">
                      <button
                        onClick={handleSelectAllSources}
                        className="flex-1 text-[9px] font-bold bg-amber-500/10 text-amber-300 hover:bg-amber-500/20 py-0.5 rounded border border-amber-500/20 transition-all"
                      >
                        Select All
                      </button>
                      <button
                        onClick={handleDeselectAllSources}
                        className="flex-1 text-[9px] font-bold bg-slate-900 text-slate-500 hover:text-rose-400 py-0.5 rounded border border-slate-800 transition-all"
                      >
                        Clear
                      </button>
                    </div>

                    {/* SELECTABLE BRANCH CHIPS */}
                    <div className="h-44 overflow-y-auto pr-1 flex flex-col gap-1">
                      {uniqueBranches
                        .filter(b => b.toLowerCase().includes(sourceSearchTerm.toLowerCase()))
                        .map(b => {
                          const isSelected = sourceStores.includes(b);
                          return (
                            <button
                              key={b}
                              onClick={() => handleToggleSourceStore(b)}
                              className={`flex items-center justify-between text-left px-2.5 py-1.5 rounded text-xs transition-all ${
                                isSelected
                                  ? 'bg-amber-500/15 text-amber-300 border border-amber-500/35 shadow-[0_0_8px_rgba(245,158,11,0.15)] font-bold'
                                  : 'bg-slate-900/60 hover:bg-slate-900 text-slate-400 border border-transparent'
                              }`}
                            >
                              <span>{b}</span>
                              {isSelected && <Check className="h-3.5 w-3.5 text-amber-400" />}
                            </button>
                          );
                        })}
                    </div>
                  </div>

                  {/* DESTINATION SELECTION WINDOW */}
                  <div className="bg-slate-950/50 border border-slate-850 rounded-xl p-3">
                    <div className="flex items-center justify-between mb-2 pb-1.5 border-b border-slate-900">
                      <span className="text-[11px] font-bold text-cyan-400 uppercase tracking-widest block">
                        Dest Stores ({destinationStores.length})
                      </span>
                    </div>

                    <input
                      type="text"
                      placeholder="Filter dest store..."
                      className="w-full text-[11px] bg-slate-900 border border-slate-800 rounded px-2 py-1 mb-2 text-white outline-none focus:border-cyan-500"
                      value={destSearchTerm}
                      onChange={(e) => setDestSearchTerm(e.target.value)}
                    />

                    <div className="flex gap-1.5 mb-2">
                      <button
                        onClick={handleSelectAllDests}
                        className="flex-1 text-[9px] font-bold bg-cyan-500/10 text-cyan-300 hover:bg-cyan-500/20 py-0.5 rounded border border-cyan-500/20 transition-all"
                      >
                        Select All
                      </button>
                      <button
                        onClick={handleDeselectAllDests}
                        className="flex-1 text-[9px] font-bold bg-slate-900 text-slate-500 hover:text-rose-400 py-0.5 rounded border border-slate-800 transition-all"
                      >
                        Clear
                      </button>
                    </div>

                    {/* SELECTABLE BRANCH CHIPS */}
                    <div className="h-44 overflow-y-auto pr-1 flex flex-col gap-1">
                      {uniqueBranches
                        .filter(b => b.toLowerCase().includes(destSearchTerm.toLowerCase()))
                        .map(b => {
                          const isSelected = destinationStores.includes(b);
                          return (
                            <button
                              key={b}
                              onClick={() => handleToggleDestStore(b)}
                              className={`flex items-center justify-between text-left px-2.5 py-1.5 rounded text-xs transition-all ${
                                isSelected
                                  ? 'bg-cyan-500/15 text-cyan-200 border border-cyan-500/35 shadow-[0_0_8px_rgba(6,182,212,0.15)] font-bold'
                                  : 'bg-slate-900/60 hover:bg-slate-900 text-slate-400 border border-transparent'
                              }`}
                            >
                              <span>{b}</span>
                              {isSelected && <Check className="h-3.5 w-3.5 text-cyan-400" />}
                            </button>
                          );
                        })}
                    </div>
                  </div>

                </div>
              )}
            </NeonCard>

          </div>

          {/* MAIN CONFIGS AND DYNAMIC SIMULATION ROW */}
          <div className="lg:col-span-12 xl:col-span-7 flex flex-col gap-6">

            {/* NEON STEP 3 PANEL: MSL CONFIGURATOR */}
            <MslEditor
              config={mslConfig}
              onChange={setMslConfig}
              onReset={handleResetMsl}
            />

            {/* NEON STEP 4 CARD: STRATEGY CHOOSERS & RUNTRIGGER */}
            <NeonCard color="green" className="shadow-[0_0_15px_rgba(16,185,129,0.1)]">
              <div className="mb-4 flex items-center justify-between">
                <span className="text-xs font-mono font-bold tracking-widest text-emerald-400">STEP 03</span>
                <span className="text-xs bg-emerald-950/60 font-medium border border-emerald-500/20 text-emerald-300 px-2 py-0.5 rounded">
                  ALLOCATION TRIGGER
                </span>
              </div>
              
              <h2 className="text-lg font-bold text-white uppercase mb-1">
                Execution Settings & Run Simulation
              </h2>
              <p className="text-slate-400 text-xs mb-4">
                Choose the default allocation strategy for cases where items do not have sales in the destination store. Once finished, click on Run to generate proposals.
              </p>

              {/* THREE MAIN BUTTON STRATEGIES */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mb-5">
                <button
                  type="button"
                  onClick={() => {
                    setDefaultStrategy('automatic');
                    addLog('🎯 Strategy adjusted: AUTOMATIC. High sales prioritization only.');
                  }}
                  className={`flex flex-col text-left p-3 rounded-xl border transition-all duration-300 ${
                    defaultStrategy === 'automatic'
                      ? 'bg-cyan-500/10 border-cyan-400 shadow-[0_0_12px_rgba(6,182,212,0.25)] text-cyan-300'
                      : 'bg-slate-950/40 border-slate-800 text-slate-400 hover:bg-slate-900 hover:text-slate-200'
                  }`}
                >
                  <div className="flex items-center gap-1.5 font-bold uppercase text-xs tracking-wider mb-1">
                    <Zap className="h-4 w-4 animate-bounce text-cyan-400" />
                    Automatic Mode
                  </div>
                  <span className="text-[10px] opacity-75 leading-normal">
                    Allocates strictly when destination store sales &gt; 1 AND stock is depleted. Priority given on sales volume.
                  </span>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setDefaultStrategy('fresh_allocate');
                    addLog('🎯 Strategy adjusted: FRESH ALLOCATE. Global forced distribution.');
                  }}
                  className={`flex flex-col text-left p-3 rounded-xl border transition-all duration-300 ${
                    defaultStrategy === 'fresh_allocate'
                      ? 'bg-rose-500/10 border-rose-400 shadow-[0_0_12px_rgba(244,63,94,0.25)] text-rose-300'
                      : 'bg-slate-950/40 border-slate-800 text-slate-400 hover:bg-slate-900 hover:text-slate-200'
                  }`}
                >
                  <div className="flex items-center gap-1.5 font-bold uppercase text-xs tracking-wider mb-1">
                    <PlusCircle className="h-4 w-4 text-rose-400" />
                    Fresh Allocate
                  </div>
                  <span className="text-[10px] opacity-75 leading-normal">
                    Allocates size items up to their MSL automatically even if there is no previous sales record in destination stores.
                  </span>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setDefaultStrategy('ignore');
                    addLog('🎯 Strategy adjusted: IGNORE. Excluded non-sales pairings.');
                  }}
                  className={`flex flex-col text-left p-3 rounded-xl border transition-all duration-300 ${
                    defaultStrategy === 'ignore'
                      ? 'bg-amber-500/10 border-amber-400 shadow-[0_0_12px_rgba(245,158,11,0.25)] text-amber-300'
                      : 'bg-slate-950/40 border-slate-800 text-slate-400 hover:bg-slate-900 hover:text-slate-200'
                  }`}
                >
                  <div className="flex items-center gap-1.5 font-bold uppercase text-xs tracking-wider mb-1">
                    <Ban className="h-4 w-4 text-amber-400" />
                    Ignore Mode
                  </div>
                  <span className="text-[10px] opacity-75 leading-normal">
                    Excludes non-sales destinations from proposal. Ideal for optimizing only active selling units.
                  </span>
                </button>
              </div>

              {/* ACTION EXECUTE TRIGGER */}
              <div className="flex flex-wrap items-center justify-between gap-4 border-t border-slate-800 pt-4">
                <div className="flex items-center gap-2">
                  <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 font-mono text-xs font-bold">
                    i
                  </div>
                  <span className="text-[11px] text-slate-400 max-w-sm">
                    Sources are sorted automatically and drawn based on slow-moving volumes first. Destination stores are prioritised by historical net sales volume.
                  </span>
                </div>

                <button
                  onClick={handleTriggerAllocation}
                  disabled={isLoading || parsedData.length === 0}
                  className={`flex items-center gap-2 font-sans font-bold text-xs uppercase tracking-wider px-6 py-3.5 rounded-xl border transition-all duration-300 group ${
                    parsedData.length === 0
                      ? 'bg-slate-900 text-slate-600 border-slate-850 cursor-not-allowed'
                      : 'bg-emerald-600 border-emerald-500 text-white hover:bg-emerald-500 hover:shadow-[0_0_20px_rgba(16,185,129,0.5)] transform hover:-translate-y-0.5 active:translate-y-0 shadow-lg'
                  }`}
                >
                  <Play className={`h-4 w-4 animate-pulse group-hover:scale-110 ${isLoading ? 'animate-spin' : ''}`} />
                  {isLoading ? 'Processing Ingest...' : '🚀 Execute Master Allocation'}
                </button>
              </div>
            </NeonCard>

          </div>

          {/* ENGINE METRICS & SYSTEM OUT TERMINAL LOGS */}
          <div className="lg:col-span-12">
            
            {/* TERMINAL LOGS SCREEN */}
            <div className="rounded-2xl border border-indigo-500/20 bg-slate-950 p-4 font-mono select-none">
              <div className="flex items-center justify-between border-b border-slate-900 pb-2 mb-2 text-[10px]">
                <div className="flex items-center gap-1.5">
                  <span className="h-2 w-2 rounded-full bg-indigo-500" />
                  <span className="text-slate-400 font-bold">CYBER CORE CONSOLE MONITOR</span>
                </div>
                <span className="text-indigo-400 animate-pulse font-bold">[ONLINE STATUS: STABLE]</span>
              </div>
              <div className="h-28 overflow-y-auto flex flex-col gap-1 text-[11px] text-slate-400 pr-2">
                {systemLogs.map((log, index) => (
                  <div key={index} className="leading-relaxed hover:text-[white] transition-colors truncate">
                    <span className="text-slate-600 select-none mr-2">&gt;&gt;</span>
                    {log}
                  </div>
                ))}
              </div>
            </div>

          </div>

          {/* SIMULATION RESULTS PANEL PANEL */}
          {hasCalculated && (
            <div className="lg:col-span-12 animate-fade-in">
              <div className="relative rounded-2xl border border-cyan-500/40 bg-slate-900/60 p-6 shadow-[0_0_25px_rgba(6,182,212,0.15)] backdrop-blur-xl">
                
                {/* Header */}
                <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-800 pb-5 mb-5">
                  <div>
                    <div className="flex items-center gap-2 text-cyan-400 font-mono text-xs uppercase tracking-widest font-bold">
                      <Sparkles className="h-4 w-4 animate-spin-slow" />
                      CALCULATED STOCK TRANSFER SUGGESTIONS
                    </div>
                    <h3 className="text-xl font-bold font-sans text-white uppercase mt-0.5">
                      Inter-Store Transfer Matrix Logs
                    </h3>
                  </div>

                  <div className="flex items-center gap-3">
                    <button
                      onClick={handleDownloadExcelReport}
                      className="flex items-center gap-2 font-sans font-bold text-xs uppercase tracking-wider px-5 py-3 rounded-lg border border-emerald-500 text-emerald-400 hover:bg-emerald-500/10 hover:shadow-[0_0_20px_rgba(16,185,129,0.3)] transition-all bg-emerald-950/20"
                    >
                      <Download className="h-4 w-4 animate-bounce" />
                      📥 Download Excel Report (.xlsx)
                    </button>
                  </div>
                </div>

                {/* Proposals stats indicators */}
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 bg-slate-950/50 rounded-xl border border-slate-850 p-4 mb-6">
                  <div>
                    <span className="text-[10px] text-slate-500 uppercase tracking-wider font-bold">Active Transfer Rows:</span>
                    <div className="text-lg font-mono font-bold text-emerald-400 mt-0.5">{allocationSummary.activeTransferRows} rows</div>
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-500 uppercase tracking-wider font-bold">Total Allocated Units:</span>
                    <div className="text-lg font-mono font-bold text-cyan-400 mt-0.5">{allocationSummary.sumQty.toLocaleString()} units</div>
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-500 uppercase tracking-wider font-bold">Total Allocation Valuation:</span>
                    <div className="text-lg font-mono font-bold text-purple-400 mt-0.5">
                      {new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 }).format(allocationSummary.sumVal)}
                    </div>
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-500 uppercase tracking-wider font-bold">Beneficiary Store Outlets:</span>
                    <div className="text-lg font-mono font-bold text-amber-500 mt-0.5">{allocationSummary.targetsStores} target outlets</div>
                  </div>
                </div>

                {/* Table search filters and tabs controls */}
                <div className="flex flex-wrap items-center justify-between gap-4 mb-4">
                  <div className="flex items-center gap-2 bg-slate-950/40 px-3.5 py-1.5 rounded-lg border border-slate-800">
                    <Filter className="h-3.5 w-3.5 text-slate-400" />
                    <span className="text-xs text-slate-400 font-bold uppercase tracking-wider">Show:</span>
                    <div className="flex items-center gap-1 ml-1.5">
                      <button
                        onClick={() => setTableStatusFilter('all')}
                        className={`text-[10px] font-bold px-2.5 py-0.5 rounded transition-all ${
                          tableStatusFilter === 'all' ? 'bg-cyan-500 text-white shadow-[0_0_8px_rgba(6,182,212,0.4)]' : 'text-slate-500 hover:text-slate-300'
                        }`}
                      >
                        ALL PROPOSALS
                      </button>
                      <button
                        onClick={() => setTableStatusFilter('allocated')}
                        className={`text-[10px] font-bold px-2.5 py-0.5 rounded transition-all ${
                          tableStatusFilter === 'allocated' ? 'bg-emerald-500 text-white shadow-[0_0_8px_rgba(16,185,129,0.4)]' : 'text-slate-500 hover:text-slate-300'
                        }`}
                      >
                        ACTIVE ONLY ({allocationSummary.activeTransferRows})
                      </button>
                      <button
                        onClick={() => setTableStatusFilter('ignored')}
                        className={`text-[10px] font-bold px-2.5 py-0.5 rounded transition-all ${
                          tableStatusFilter === 'ignored' ? 'bg-amber-500 text-white shadow-[0_0_8px_rgba(245,158,11,0.4)]' : 'text-slate-500 hover:text-slate-300'
                        }`}
                      >
                        IGNORES / EXCLUDED
                      </button>
                    </div>
                  </div>

                  {/* SEARCH FIELD */}
                  <div className="relative flex-1 max-w-sm">
                    <Search className="absolute left-3 top-2.5 h-3.5 w-3.5 text-slate-500" />
                    <input
                      type="text"
                      placeholder="Search simulation log (Barcode, store, name)..."
                      className="w-full text-xs bg-slate-950 text-white border border-slate-800 rounded-lg pl-9 pr-4 py-2 outline-none focus:border-cyan-400 focus:shadow-[0_0_10px_rgba(6,182,212,0.15)]"
                      value={tableSearchTerm}
                      onChange={(e) => setTableSearchTerm(e.target.value)}
                    />
                  </div>
                </div>

                {/* THE PRIMARY TABLE OUT */}
                <div className="overflow-x-auto rounded-xl border border-slate-800">
                  <table className="w-full table-auto text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-slate-950 border-b border-slate-800/80 text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                        <th className="p-3 font-mono text-center w-12">S.No</th>
                        <th className="p-3">STORE (Dest)</th>
                        <th className="p-3">GENDER</th>
                        <th className="p-3">STYLE</th>
                        <th className="p-3">ITEM CODE</th>
                        <th className="p-3">COMPANY</th>
                        <th className="p-3">ITEM NAME / SIZE</th>
                        <th className="p-3 text-right">MRP</th>
                        <th className="p-3 text-center">SRC STORE (Slow Stock)</th>
                        <th className="p-3 text-center">DEST STOCK</th>
                        <th className="p-3 text-center">DEST SALES</th>
                        <th className="p-3 text-center w-32">ALLOCATED QTI</th>
                        <th className="p-3 text-right">ALLOC VAL</th>
                        <th className="p-3 text-center w-52">STATE ACTIONS</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-850 bg-slate-900/10">
                      {filteredAndSortedProposals.length === 0 ? (
                        <tr>
                          <td colSpan={14} className="p-10 text-center text-xs text-slate-500 italic">
                            No matching stock transfer proposals discovered matching the active search.
                          </td>
                        </tr>
                      ) : (
                        filteredAndSortedProposals.map((prop, idx) => {
                          const hasSales = prop.destinationStoreSales > 0;
                          return (
                            <tr
                              key={prop.id}
                              className={`transition-all hover:bg-slate-900/50 ${
                                prop.allocatedQty > 0
                                  ? 'bg-emerald-950/5 border-l-2 border-l-emerald-500'
                                  : 'opacity-65 text-slate-500 border-l-2 border-l-slate-800'
                              }`}
                            >
                              <td className="p-3 text-center text-[10px] font-mono text-slate-500">
                                {idx + 1}
                              </td>
                              <td className="p-3 font-sans font-bold text-white tracking-wide">
                                {prop.destinationStore}
                              </td>
                              <td className="p-3 font-mono text-slate-300">
                                {prop.gender}
                              </td>
                              <td className="p-3 font-mono text-slate-400">
                                {prop.style}
                              </td>
                              <td className="p-3 font-mono font-bold text-cyan-300">
                                {prop.barcode}
                              </td>
                              <td className="p-3 font-sans truncate max-w-[100px]" title={prop.brand}>
                                {prop.brand}
                              </td>
                              <td className="p-3">
                                <div className="font-sans font-medium text-slate-200 line-clamp-1">{prop.itemName}</div>
                                <div className="text-[10px] text-zinc-500 font-mono">Size override: <span className="text-zinc-300 font-bold">{prop.size}</span></div>
                              </td>
                              <td className="p-3 text-right font-mono text-slate-300">
                                ₹{prop.mrp.toLocaleString()}
                              </td>
                              <td className="p-3 text-center font-mono">
                                <div className="text-[11px] text-amber-300 flex flex-col font-bold">
                                  <span>{prop.sourceStore}</span>
                                </div>
                              </td>
                              <td className="p-3 text-center font-mono font-semibold text-slate-300">
                                {prop.destinationStoreClosingQty}
                              </td>
                              <td className="p-3 text-center font-mono font-bold">
                                <span className={hasSales ? 'text-cyan-400' : 'text-slate-500'}>
                                  {prop.destinationStoreSales}
                                </span>
                              </td>
                              <td className="p-3 text-center">
                                {prop.allocatedQty > 0 ? (
                                  <div className="inline-flex items-center gap-1.5 justify-center">
                                    <input
                                      type="number"
                                      min="0"
                                      value={prop.allocatedQty}
                                      onChange={(e) => handleRowQtyChange(prop.id, Number(e.target.value))}
                                      className="w-14 text-center font-mono font-bold bg-slate-950 text-emerald-400 border border-emerald-500/30 rounded px-1.5 py-1 text-xs focus:border-emerald-500 outline-none hover:shadow-[0_0_8px_rgba(16,185,129,0.2)]"
                                    />
                                    <span className="text-[10px] text-slate-500">pcs</span>
                                  </div>
                                ) : (
                                  <span className="text-xs text-slate-500 font-mono italic">Excluded</span>
                                )}
                              </td>
                              <td className="p-3 text-right font-mono font-bold text-white">
                                {prop.allocatedQty > 0 ? `₹${prop.allocationVal.toLocaleString()}` : '₹0'}
                              </td>
                              <td className="p-3 text-center">
                                {/* State changer buttons: Automatic, Fresh Allocate, Ignore */}
                                <div className="flex items-center gap-1 justify-center">
                                  
                                  <button
                                    onClick={() => handleRowTypeOverride(prop.id, 'fresh')}
                                    className={`px-2 py-1 rounded text-[9px] font-bold uppercase tracking-wider border transition-all ${
                                      prop.type === 'fresh' && prop.allocatedQty > 0
                                        ? 'bg-rose-500/20 border-rose-500 text-rose-300 shadow-[0_0_8px_rgba(244,63,94,0.3)]'
                                        : 'bg-slate-950/40 border-slate-800 text-slate-400 hover:text-slate-300 hover:bg-slate-900/60'
                                    }`}
                                    title="Fresh Allocate up to custom MSL regardless of existing store item records"
                                  >
                                    Fresh Allocate
                                  </button>

                                  <button
                                    onClick={() => handleRowTypeOverride(prop.id, 'ignored')}
                                    className={`px-2 py-1 rounded text-[9px] font-bold uppercase tracking-wider border transition-all ${
                                      prop.allocatedQty === 0
                                        ? 'bg-amber-500/20 border-amber-500 text-amber-300 shadow-[0_0_8px_rgba(245,158,11,0.3)]'
                                        : 'bg-slate-950/40 border-slate-800 text-slate-400 hover:text-slate-300 hover:bg-slate-900/60'
                                    }`}
                                    title="Ignore and discard this product transfer pairing"
                                  >
                                    Ignore
                                  </button>

                                </div>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>

                <div className="flex flex-wrap items-center justify-between gap-4 mt-5">
                  <div className="flex items-center gap-2">
                    <AlertTriangle className="h-4 w-4 text-amber-500" />
                    <span className="text-xs text-zinc-400 font-mono">
                      * Fresh manual overrides dynamically apply custom MSL configurations over selected destinations.
                    </span>
                  </div>

                  <button
                    onClick={handleDownloadExcelReport}
                    className="flex items-center gap-2 font-sans font-bold text-xs uppercase tracking-wider px-6 py-3.5 rounded-lg border border-cyan-500 text-cyan-300 bg-cyan-950/20 shadow-lg hover:shadow-[0_0_20px_rgba(6,182,212,0.4)] transition-all hover:bg-cyan-500/10"
                  >
                    <Download className="h-4.5 w-4.5" />
                    📤 Download Final Allocation Spreadsheet
                  </button>
                </div>

              </div>
            </div>
          )}

        </div>

      </div>
    </div>
  );
}
