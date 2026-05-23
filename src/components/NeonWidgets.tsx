/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import { UploadCloud, FileSpreadsheet, FileCode, CheckCircle, AlertTriangle } from 'lucide-react';

interface NeonCardProps {
  children: React.ReactNode;
  color?: 'cyan' | 'rose' | 'green' | 'amber' | 'violet';
  className?: string;
  id?: string;
}

export function NeonCard({ children, color = 'cyan', className = '', id }: NeonCardProps) {
  // Define shadow gradients matching the color
  const colorShadows = {
    cyan: 'shadow-[0_0_20px_rgba(6,182,212,0.15)] border-cyan-500/30 hover:border-cyan-400/70 hover:shadow-[0_0_25px_rgba(6,182,212,0.35)]',
    rose: 'shadow-[0_0_20px_rgba(244,63,94,0.15)] border-rose-500/30 hover:border-rose-400/70 hover:shadow-[0_0_25px_rgba(244,63,94,0.35)]',
    green: 'shadow-[0_0_20px_rgba(16,185,129,0.15)] border-emerald-500/30 hover:border-emerald-400/70 hover:shadow-[0_0_25px_rgba(16,185,129,0.35)]',
    amber: 'shadow-[0_0_20px_rgba(245,158,11,0.15)] border-amber-500/30 hover:border-amber-400/70 hover:shadow-[0_0_25px_rgba(245,158,11,0.35)]',
    violet: 'shadow-[0_0_20px_rgba(139,92,246,0.15)] border-violet-500/30 hover:border-violet-400/70 hover:shadow-[0_0_25px_rgba(139,92,246,0.35)]',
  };

  return (
    <div
      id={id}
      className={`relative rounded-2xl bg-slate-900/40 backdrop-blur-xl border p-6 transition-all duration-300 ease-out hover:-translate-y-1 ${colorShadows[color]} ${className}`}
    >
      {/* Visual background gradient accents */}
      <div className={`absolute top-0 left-0 w-2 h-10 rounded-tr-md rounded-bl-md transition-colors ${
        color === 'cyan' ? 'bg-cyan-500' :
        color === 'rose' ? 'bg-rose-500' :
        color === 'green' ? 'bg-emerald-500' :
        color === 'amber' ? 'bg-amber-500' : 'bg-violet-500'
      }`} />
      {children}
    </div>
  );
}

interface NeonBadgeProps {
  label: string;
  variant?: 'cyan' | 'rose' | 'green' | 'amber' | 'slate';
}

export function NeonBadge({ label, variant = 'cyan' }: NeonBadgeProps) {
  const styles = {
    cyan: 'bg-cyan-950/40 text-cyan-400 border border-cyan-500/30 shadow-[0_0_6px_rgba(6,182,212,0.1)]',
    rose: 'bg-rose-950/40 text-rose-400 border border-rose-500/30 shadow-[0_0_6px_rgba(244,63,94,0.1)]',
    green: 'bg-emerald-950/40 text-emerald-400 border border-emerald-500/30 shadow-[0_0_6px_rgba(16,185,129,0.1)]',
    amber: 'bg-amber-950/40 text-amber-400 border border-amber-500/30 shadow-[0_0_6px_rgba(245,158,11,0.1)]',
    slate: 'bg-slate-900 text-slate-400 border border-slate-800'
  };

  return (
    <span className={`inline-flex items-center px-2 py-0.5 rounded text-[10px] font-bold uppercase tracking-wider ${styles[variant]}`}>
      {label}
    </span>
  );
}

interface NeonButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  children: React.ReactNode;
  color?: 'cyan' | 'rose' | 'green' | 'amber';
  glow?: boolean;
  className?: string;
}

export function NeonButton({ children, color = 'cyan', glow = true, className = '', ...props }: NeonButtonProps) {
  const styles = {
    cyan: 'bg-cyan-950/20 text-cyan-400 border border-cyan-500/60 hover:bg-cyan-500/10 active:bg-cyan-500/20 shadow-[0_0_12px_rgba(6,182,212,0.1)] hover:shadow-[0_0_20px_rgba(6,182,212,0.4)] hover:border-cyan-400',
    rose: 'bg-rose-950/20 text-rose-400 border border-rose-500/60 hover:bg-rose-500/10 active:bg-rose-500/20 shadow-[0_0_12px_rgba(244,63,94,0.1)] hover:shadow-[0_0_20px_rgba(244,63,94,0.4)] hover:border-rose-400',
    green: 'bg-emerald-950/20 text-emerald-400 border border-emerald-500/60 hover:bg-emerald-500/10 active:bg-emerald-500/20 shadow-[0_0_12px_rgba(16,185,129,0.1)] hover:shadow-[0_0_20px_rgba(16,185,129,0.4)] hover:border-emerald-400',
    amber: 'bg-amber-950/20 text-amber-400 border border-amber-500/60 hover:bg-amber-500/10 active:bg-amber-500/20 shadow-[0_0_12px_rgba(245,158,11,0.1)] hover:shadow-[0_0_20px_rgba(245,158,11,0.4)] hover:border-amber-400',
  };

  return (
    <button
      className={`font-sans text-xs uppercase tracking-wider font-bold px-4 py-2.5 rounded-lg transition-all duration-300 transform active:scale-95 ${styles[color]} ${className}`}
      {...props}
    >
      {children}
    </button>
  );
}

interface FileDropzoneProps {
  onFileSelect: (file: File) => void;
  accept: string;
  label: string;
  subLabel?: string;
  isLoaded: boolean;
  loadedInfo?: string;
  color?: 'cyan' | 'rose' | 'green' | 'amber';
  iconType?: 'spreadsheet' | 'code';
}

export function FileDropzone({
  onFileSelect,
  accept,
  label,
  subLabel = "Drag & Drop or Click to browse",
  isLoaded,
  loadedInfo = "",
  color = "cyan",
  iconType = "spreadsheet"
}: FileDropzoneProps) {
  const [isDragActive, setIsDragActive] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setIsDragActive(true);
    } else if (e.type === "dragleave") {
      setIsDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      onFileSelect(e.dataTransfer.files[0]);
    }
  };

  const handleClick = () => {
    fileInputRef.current?.click();
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      onFileSelect(e.target.files[0]);
    }
  };

  const colorStyles = {
    cyan: {
      borderActive: 'border-cyan-400 bg-cyan-950/20 shadow-[0_0_20px_rgba(6,182,212,0.3)]',
      borderDefault: 'border-slate-800 hover:border-cyan-500/50 hover:bg-cyan-950/5',
      iconColor: 'text-cyan-400',
      badgeColor: 'bg-cyan-500/10 text-cyan-300 font-mono border-cyan-500/20'
    },
    rose: {
      borderActive: 'border-rose-400 bg-rose-950/20 shadow-[0_0_20px_rgba(244,63,94,0.3)]',
      borderDefault: 'border-slate-800 hover:border-rose-500/50 hover:bg-rose-950/5',
      iconColor: 'text-rose-400',
      badgeColor: 'bg-rose-500/10 text-rose-300 font-mono border-rose-500/20'
    },
    green: {
      borderActive: 'border-emerald-400 bg-emerald-950/20 shadow-[0_0_20px_rgba(16,185,129,0.3)]',
      borderDefault: 'border-slate-800 hover:border-emerald-500/50 hover:bg-emerald-950/5',
      iconColor: 'text-emerald-400',
      badgeColor: 'bg-emerald-500/10 text-emerald-300 font-mono border-emerald-500/20'
    },
    amber: {
      borderActive: 'border-amber-400 bg-amber-950/20 shadow-[0_0_20px_rgba(245,158,11,0.3)]',
      borderDefault: 'border-slate-800 hover:border-amber-500/50 hover:bg-amber-950/5',
      iconColor: 'text-amber-400',
      badgeColor: 'bg-amber-500/10 text-amber-300 font-mono border-amber-500/20'
    }
  };

  const style = colorStyles[color];

  return (
    <div
      onDragEnter={handleDrag}
      onDragOver={handleDrag}
      onDragLeave={handleDrag}
      onDrop={handleDrop}
      onClick={handleClick}
      className={`relative flex flex-col items-center justify-center border-2 border-dashed rounded-2xl p-6 text-center cursor-pointer transition-all duration-300 ${
        isDragActive ? style.borderActive : style.borderDefault
      }`}
    >
      <input
        ref={fileInputRef}
        type="file"
        accept={accept}
        onChange={handleInputChange}
        className="hidden"
      />

      {isLoaded ? (
        <div className="flex flex-col items-center animate-fade-in">
          <div className="relative mb-3 flex h-16 w-16 items-center justify-center rounded-2xl bg-emerald-950/30 border border-emerald-500/60 shadow-[0_0_15px_rgba(16,185,129,0.4)] text-emerald-400">
            <CheckCircle className="h-8 w-8 animate-pulse" />
          </div>
          <span className="text-sm font-sans font-semibold text-white tracking-wide uppercase mb-1">
            File Loaded Successfully
          </span>
          <span className={`px-2.5 py-1 text-xs rounded border mt-2 ${style.badgeColor}`}>
            {loadedInfo}
          </span>
        </div>
      ) : (
        <div className="flex flex-col items-center">
          <div className={`relative mb-4 flex h-14 w-14 items-center justify-center rounded-2xl bg-slate-900 border border-slate-800 text-slate-400 ${style.iconColor} transition-all group-hover:scale-110`}>
            {iconType === 'spreadsheet' ? (
              <FileSpreadsheet className="h-7 w-7" />
            ) : (
              <FileCode className="h-7 w-7" />
            )}
          </div>
          <span className="text-sm font-sans font-bold text-slate-300 uppercase tracking-wider mb-1">
            {label}
          </span>
          <span className="text-xs text-slate-500 max-w-xs leading-relaxed">
            {subLabel}
          </span>
        </div>
      )}
    </div>
  );
}
