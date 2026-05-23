/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { MslConfig } from '../types';
import { Settings, RefreshCw, Layers, CheckCircle2 } from 'lucide-react';

interface MslEditorProps {
  config: MslConfig;
  onChange: (newConfig: MslConfig) => void;
  onReset: () => void;
}

export default function MslEditor({ config, onChange, onReset }: MslEditorProps) {
  const [activeCategory, setActiveCategory] = useState<'boys' | 'girls' | 'kids' | 'ladies' | 'gents'>('gents');

  const handleDefaultValChange = (category: keyof MslConfig, val: number) => {
    const updated = { ...config };
    updated[category] = {
      ...updated[category],
      defaultVal: Math.max(0, val)
    };
    onChange(updated);
  };

  const handleSizeValChange = (category: keyof MslConfig, size: string, val: number) => {
    const updated = { ...config };
    const cleanSize = size.trim();
    updated[category] = {
      ...updated[category],
      sizes: {
        ...updated[category].sizes,
        [cleanSize]: Math.max(0, val)
      }
    };
    onChange(updated);
  };

  const addCustomSizeRow = (category: keyof MslConfig, sizeKey: string) => {
    if (!sizeKey.trim()) return;
    const updated = { ...config };
    updated[category] = {
      ...updated[category],
      sizes: {
        ...updated[category].sizes,
        [sizeKey.trim()]: 2
      }
    };
    onChange(updated);
  };

  const removeSizeOverride = (category: keyof MslConfig, sizeKey: string) => {
    const updated = { ...config };
    const nextSizes = { ...updated[category].sizes };
    delete nextSizes[sizeKey];
    updated[category] = {
      ...updated[category],
      sizes: nextSizes
    };
    onChange(updated);
  };

  const [newSizeInput, setNewSizeInput] = useState('');

  const categoriesList: { id: keyof MslConfig; label: string; desc: string }[] = [
    { id: 'gents', label: 'Gents', desc: 'Sizes: 6-13 overrides' },
    { id: 'ladies', label: 'Ladies', desc: 'Sizes: 5-14 overrides' },
    { id: 'boys', label: 'Boys', desc: 'Uniform count for sizes' },
    { id: 'girls', label: 'Girls', desc: 'Uniform count for sizes' },
    { id: 'kids', label: 'Kids', desc: 'Uniform count for sizes' },
  ];

  const currentConfig = config[activeCategory];

  return (
    <div className="relative overflow-hidden rounded-2xl border border-rose-500/30 bg-slate-900/80 p-5 shadow-[0_0_20px_rgba(244,63,94,0.15)] backdrop-blur-md">
      {/* Glow Effects */}
      <div className="absolute -top-10 -right-10 h-32 w-32 rounded-full bg-rose-500/10 blur-2xl pointer-events-none" />
      <div className="absolute -bottom-10 -left-10 h-32 w-32 rounded-full bg-indigo-500/10 blur-2xl pointer-events-none" />

      {/* Header */}
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3 border-b border-slate-800 pb-3">
        <div className="flex items-center gap-2">
          <Settings className="h-5 w-5 text-rose-500 animate-spin-slow" />
          <h3 className="font-sans text-base font-semibold tracking-wide text-rose-400 uppercase">
            ⚡ Variable MSL Target Configuration
          </h3>
        </div>
        <button
          onClick={onReset}
          className="flex items-center gap-1 text-xs font-semibold text-rose-400 hover:text-rose-300 border border-rose-500/30 rounded px-2.5 py-1 transition-all bg-rose-500/10 hover:shadow-[0_0_10px_rgba(244,63,94,0.4)]"
        >
          <RefreshCw className="h-3 w-3" />
          Reset Defaults
        </button>
      </div>

      <p className="text-slate-400 text-xs mb-4">
        Change these values dynamically. If no modifications are made, the allocation engine uses standard fixed Retail MSLs.
      </p>

      {/* Category Selection Tabs & Editors */}
      <div className="grid grid-cols-1 md:grid-cols-12 gap-4">
        {/* Sidebar Tabs */}
        <div className="md:col-span-4 flex flex-row md:flex-col gap-2 overflow-x-auto pb-2 md:pb-0 scrollbar-none">
          {categoriesList.map((cat) => (
            <button
              key={cat.id}
              onClick={() => {
                setActiveCategory(cat.id);
                setNewSizeInput('');
              }}
              className={`flex-1 md:flex-none text-left p-2.5 rounded-lg border transition-all duration-300 shrink-0 ${
                activeCategory === cat.id
                  ? 'bg-rose-500/20 border-rose-500/80 text-rose-300 shadow-[0_0_10px_rgba(244,63,94,0.3)]'
                  : 'bg-slate-950/40 border-slate-800 text-slate-400 hover:text-slate-300 hover:border-slate-700'
              }`}
            >
              <div className="font-medium text-xs uppercase tracking-wider">{cat.label}</div>
              <div className="text-[10px] opacity-60 hidden md:block">{cat.desc}</div>
            </button>
          ))}
        </div>

        {/* Editor Body */}
        <div className="md:col-span-8 bg-slate-950/40 rounded-xl border border-slate-800 p-4">
          <div className="mb-4 flex items-center justify-between">
            <span className="text-xs font-semibold uppercase tracking-wider text-rose-400">
              Editing: <span className="text-white font-bold">{activeCategory}</span>
            </span>
            <div className="flex items-center gap-2">
              <span className="text-[11px] text-slate-400">Default Quantity:</span>
              <input
                type="number"
                min="0"
                value={currentConfig.defaultVal}
                onChange={(e) => handleDefaultValChange(activeCategory, Number(e.target.value))}
                className="w-14 text-center text-xs font-mono font-bold bg-slate-900 text-white border border-slate-800 rounded px-1.5 py-1 focus:border-rose-500 outline-none"
              />
            </div>
          </div>

          <div className="h-px bg-slate-800/80 my-3" />

          {/* Sizes Grid */}
          <div className="max-h-56 overflow-y-auto pr-1">
            <h4 className="text-[10px] font-bold uppercase tracking-widest text-slate-500 mb-2">
              Size-Specific Overrides
            </h4>

            {Object.keys(currentConfig.sizes).length === 0 ? (
              <div className="text-center py-4 text-xs text-slate-500 italic">
                No custom size overrides active. All sizes will allocate {currentConfig.defaultVal} units.
              </div>
            ) : (
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                {Object.entries(currentConfig.sizes).map(([size, qty]) => (
                  <div
                    key={size}
                    className="flex items-center justify-between p-1.5 rounded bg-slate-900 border border-slate-800/60 transition-all hover:border-slate-700"
                  >
                    <span className="text-xs font-mono text-slate-300 pl-1 font-bold">Size {size}:</span>
                    <div className="flex items-center gap-1">
                      <input
                        type="number"
                        min="0"
                        value={qty}
                        onChange={(e) => handleSizeValChange(activeCategory, size, Number(e.target.value))}
                        className="w-10 text-center text-xs font-mono font-semibold bg-slate-950 text-rose-300 border border-slate-800 rounded px-1 py-0.5 outline-none focus:border-rose-500"
                      />
                      <button
                        onClick={() => removeSizeOverride(activeCategory, size)}
                        className="text-slate-500 hover:text-rose-400 transition-colors px-1 text-[10px]"
                        title="Remove override"
                      >
                        ×
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Add a size override */}
          <div className="mt-4 pt-3 border-t border-slate-800 flex items-center gap-2">
            <input
              type="text"
              placeholder="Size input (e.g. 15)"
              value={newSizeInput}
              onChange={(e) => setNewSizeInput(e.target.value)}
              className="flex-1 text-xs bg-slate-900 text-white border border-slate-800 rounded px-2.5 py-1.5 focus:border-rose-500 outline-none"
            />
            <button
              onClick={() => {
                const s = newSizeInput.trim();
                if (s) {
                  addCustomSizeRow(activeCategory, s);
                  setNewSizeInput('');
                }
              }}
              className="text-xs font-bold text-white bg-rose-600 hover:bg-rose-500 px-3 py-1.5 rounded transition-all transform hover:-translate-y-0.5 active:translate-y-0 shadow-lg shadow-rose-600/30 hover:shadow-rose-600/50"
            >
              + Override
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
