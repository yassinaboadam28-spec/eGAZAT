
import React, { useState, useMemo, useCallback, useRef, useEffect } from 'react';
import type { DataRow, SortConfig, EmployeeRecord } from './types';
import { AmiriFont } from './AmiriFont';
import { saveEmployeeData, getEmployeeData, firebaseConfig } from './firebase';
// Make external libraries available, assuming they're loaded from a CDN.
declare const XLSX: any;
declare const jspdf: any;
declare const html2canvas: any;
declare const pdfjsLib: any; // For handling PDF files

// --- New types for structured summary ---
interface LeaveSummary {
  type: string;
  dayCount: number;
  hourCount: number;
  dateDetails: string;
  monthYear?: string;
}

interface EmployeeSummary {
  name: string;
  leaves: LeaveSummary[];
  initialBalance?: number; // Represents the truly initial balance before any deductions.
  currentBalance?: number; // Represents the final, calculated balance.
  photo?: string;
  workplace?: string;
}

// --- Type for Monthly Report ---
interface MonthlyReportRow {
    name: string;
    workplace: string;
    initialBalance: number;
    regularLeaves: { count: number; dates: string };
    hourlyLeaves: { days: number; hours: number };
    sickLeave: { count: number; dateRange: string };
    longLeave: { count: number; dateRange: string };
    travelLeave: { count: number; dateRange: string };
    absentLeave: { count: number; dateRange: string };
    finalBalance: number;
}

interface UploadHistoryEntry {
    fileName: string;
    uploadId: number;
}
  
// --- Storage Keys ---
const EMPLOYEES_STORAGE_KEY = 'employeeLeaveBalances';
const LEAVE_DATA_STORAGE_KEY = 'leaveDataRecords';
const PROCESSED_FILES_STORAGE_KEY = 'processedFileNames';
const LAST_BALANCE_UPDATE_KEY = 'lastBalanceUpdateMarker';
const THEME_STORAGE_KEY = 'appTheme';
const UPLOAD_HISTORY_KEY = 'uploadHistory';
const CURRENT_USER_STORAGE_KEY = 'currentUserSession';


// --- SVG Icon Components (defined outside the main component to prevent re-creation on re-renders) ---

const SunIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 3v2.25m6.364.386l-1.591 1.591M21 12h-2.25m-.386 6.364l-1.591-1.591M12 18.75V21m-4.773-4.227l-1.591 1.591M5.25 12H3m4.227-4.773L5.636 5.636M15.75 12a3.75 3.75 0 11-7.5 0 3.75 3.75 0 017.5 0z" />
    </svg>
);
  
const MoonIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M21.752 15.002A9.718 9.718 0 0118 15.75c-5.385 0-9.75-4.365-9.75-9.75 0-1.33.266-2.597.748-3.752A9.753 9.753 0 003 11.25C3 16.635 7.365 21 12.75 21a9.753 9.753 0 009.002-5.998z" />
    </svg>
);

const UploadIcon: React.FC<{ className?: string }> = ({ className }) => (
  <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
    <path strokeLinecap="round" strokeLinejoin="round" d="M12 16.5V9.75m0 0l-3.75 3.75M12 9.75l3.75 3.75M3 17.25V21h18v-3.75M4.5 12.75l7.5-7.5 7.5 7.5" />
  </svg>
);

const DownloadIcon: React.FC<{ className?: string }> = ({ className }) => (
  <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
    <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3" />
  </svg>
);

const TrashIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M14.74 9l-.346 9m-4.788 0L9.26 9m9.968-3.21c.342.052.682.107 1.022.166m-1.022-.165L18.16 19.673a2.25 2.25 0 01-2.244 2.077H8.084a2.25 2.25 0 01-2.244-2.077L4.772 5.79m14.456 0a48.108 48.108 0 00-3.478-.397m-12 .562c.34-.059.68-.114 1.022-.165m0 0a48.11 48.11 0 013.478-.397m7.5 0v-.916c0-1.18-.91-2.144-2.09-2.201a51.964 51.964 0 00-3.32 0c-1.18.057-2.09 1.022-2.09 2.201v.916m7.5 0a48.667 48.667 0 00-7.5 0" />
    </svg>
);


const SearchIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607z" />
    </svg>
);

const DocumentTextIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
    </svg>
);

const TableCellsIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M3.375 19.5h17.25m-17.25 0a1.125 1.125 0 01-1.125-1.125v-1.5c0-.621.504-1.125 1.125-1.125h17.25c.621 0 1.125.504 1.125 1.125v1.5c0 .621-.504 1.125-1.125 1.125m-17.25 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm-16.5-3.375h17.25m-17.25 0a1.125 1.125 0 01-1.125-1.125v-1.5c0-.621.504-1.125 1.125-1.125h17.25c.621 0 1.125.504 1.125 1.125v1.5c0 .621-.504 1.125-1.125 1.125m-17.25 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm-16.5-3.375h17.25m-17.25 0A1.125 1.125 0 012.25 8.25v-1.5c0-.621.504-1.125 1.125-1.125h17.25c.621 0 1.125.504 1.125 1.125v1.5c0 .621-.504 1.125-1.125 1.125m-17.25 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75z" />
    </svg>
);

const ClipboardDocumentIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M15.666 3.888A2.25 2.25 0 0013.5 2.25h-3c-1.03 0-1.9.693-2.166 1.638m7.332 0c.055.194.084.4.084.612v0a2.25 2.25 0 01-2.25 2.25h-1.5a2.25 2.25 0 01-2.25-2.25v0c0-.212.03-.418.084-.612m7.332 0c.646.049 1.288.11 1.927.184 1.1.128 1.907 1.077 1.907 2.185V19.5a2.25 2.25 0 01-2.25 2.25H6.75A2.25 2.25 0 014.5 19.5V6.257c0-1.108.806-2.057 1.907-2.185a48.208 48.208 0 011.927-.184" />
    </svg>
);

const ClockIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M12 6v6h4.5m4.5 0a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
);

const BarsArrowUpIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M3 4.5h14.25M3 9h9.75M3 13.5h5.25m5.25-.75L17.25 9m0 0L21 12.75M17.25 9v12" />
    </svg>
);

const AlphabeticalSortIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25H12" />
    </svg>
);

const BriefcaseIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M20.25 14.15v4.075c0 1.313-.943 2.5-2.206 2.5H6.081c-1.262 0-2.206-1.187-2.206-2.5v-4.075m16.35 0c.225.045.45.08.675.112v-4.075c0-1.313-.943-2.5-2.206-2.5h-12.17c-1.263 0-2.206 1.187-2.206 2.5v4.075c.225-.032.45-.067.675-.112M16.5 7.5h-9" />
    </svg>
);

const HeartIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M21 8.25c0-2.485-2.099-4.5-4.688-4.5-1.935 0-3.597 1.126-4.312 2.733-.715-1.607-2.377-2.733-4.313-2.733C5.1 3.75 3 5.765 3 8.25c0 7.22 9 12 9 12s9-4.78 9-12z" />
    </svg>
);

const PrinterIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M6.72 13.829c-.24.03-.48.062-.72.096m.72-.096a42.415 42.415 0 0110.56 0m-10.56 0L6 3.369m0 0c.071.01.141.02.21.031m-1.215 10.518a42.597 42.597 0 01-2.073-.09m2.073.09a42.415 42.415 0 0010.56 0m-10.56 0c.316.05.632.095.948.135M16.5 10.5V6.75a4.5 4.5 0 10-9 0v3.75m-.75 11.25h10.5a2.25 2.25 0 002.25-2.25v-6.75a2.25 2.25 0 00-2.25-2.25H6.75a2.25 2.25 0 00-2.25 2.25v6.75a2.25 2.25 0 002.25 2.25z" />
    </svg>
);

const PdfIcon: React.FC<{ className?: string }> = ({ className }) => (
  <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m.158 10.302L9 18.25m1.5-3.948L10.5 18.25m0 0l.158.25m-.158-.25L9.342 16.5m.208 1.75h.208m-3.375 0h3.375m-3.375 0h.008v.015h-.008v-.015z" />
    <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
  </svg>
);

// --- New Icons for Summary Cards ---
const UserGroupIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M18 18.72a9.094 9.094 0 00-12 0m12 0a9.094 9.094 0 00-12 0m12 0v.006M18 18.72v.006m-12 0v.006m0-10.74a6 6 0 0112 0v1.5a6 6 0 10-12 0v-1.5zm12 0a6 6 0 00-12 0v1.5a6 6 0 1012 0v-1.5z" />
    </svg>
);

const UserCircleIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M17.982 18.725A7.488 7.488 0 0012 15.75a7.488 7.488 0 00-5.982 2.975m11.963 0a9 9 0 10-11.963 0m11.963 0A8.966 8.966 0 0112 21a8.966 8.966 0 01-5.982-2.275M15 9.75a3 3 0 11-6 0 3 3 0 016 0z" />
    </svg>
);


const UsersIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M15 19.128a9.38 9.38 0 002.625.372 9.337 9.337 0 004.121-.952 4.125 4.125 0 00-7.533-2.493M15 19.128v-.003c0-1.113-.285-2.16-.786-3.07M15 19.128v.106A12.318 12.318 0 018.624 21c-2.331 0-4.512-.645-6.374-1.766l-.001-.109a6.375 6.375 0 0111.964-4.663M12 3.375c-3.418 0-6.167 2.023-6.167 4.5s2.75 4.5 6.167 4.5 6.167-2.023 6.167-4.5S15.418 3.375 12 3.375z" />
    </svg>
);

const PencilIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M16.862 4.487l1.687-1.688a1.875 1.875 0 112.652 2.652L10.582 16.07a4.5 4.5 0 01-1.897 1.13L6 18l.8-2.685a4.5 4.5 0 011.13-1.897l8.932-8.931zm0 0L19.5 7.125M18 14v4.75A2.25 2.25 0 0115.75 21H5.25A2.25 2.25 0 013 18.75V8.25A2.25 2.25 0 015.25 6H10" />
    </svg>
);

const CalendarDaysIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M6.75 3v2.25M17.25 3v2.25M3 18.75V7.5a2.25 2.25 0 012.25-2.25h13.5A2.25 2.25 0 0121 7.5v11.25m-18 0A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75m-18 0h18" />
    </svg>
);

const TrophyIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M16.5 18.75h-9a9 9 0 119 0zM16.5 18.75a9 9 0 10-9 0m9 0h-9m9 0h-9M9 13.5v6.375m6-6.375v6.375m-3-3.375V18.75m0-12.75a3 3 0 00-3-3H9a3 3 0 00-3 3v.75m6 .75v-3.75m-6 3.75v-3.75m3 3.75v-3.75M9 3h6" />
    </svg>
);

const ArrowPathIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M16.023 9.348h4.992v-.001M2.985 19.644v-4.992m0 0h4.992m-4.993 0l3.181 3.183a8.25 8.25 0 0011.664 0l3.181-3.183m-4.991-2.696a8.25 8.25 0 00-11.664 0l-3.181 3.183" />
    </svg>
);

const MinusCircleIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M15 12H9m12 0a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
);

const LogoutIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M15.75 9V5.25A2.25 2.25 0 0013.5 3h-6a2.25 2.25 0 00-2.25 2.25v13.5A2.25 2.25 0 007.5 21h6a2.25 2.25 0 002.25-2.25V15m3 0l3-3m0 0l-3-3m3 3H9" />
    </svg>
);

const Bars3Icon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5" />
    </svg>
);


const SortIcon: React.FC<{ direction?: 'ascending' | 'descending' }> = ({ direction }) => {
  if (!direction) {
    return (
      <svg className="w-4 h-4 text-gray-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M8.25 15L12 18.75 15.75 15m-7.5-6L12 5.25 15.75 9" />
      </svg>
    );
  }
  if (direction === 'ascending') {
    return (
      <svg className="w-4 h-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M4.5 15.75l7.5-7.5 7.5 7.5" />
      </svg>
    );
  }
  return (
    <svg className="w-4 h-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 8.25l-7.5 7.5-7.5-7.5" />
    </svg>
  );
};

/**
 * Normalizes Arabic names for consistent matching.
 */
const normalizeName = (name: string): string => {
    if (!name) return '';
    return name
        .trim()
        .replace(/\s+/g, '') // Remove all whitespace
        .replace(/[أإآ]/g, 'ا')  // Normalize Alef
        .replace(/ؤ/g, 'و')    // Normalize Waw
        .replace(/[ئى]/g, 'ي')    // Normalize Yeh and Alef Maksura
        .replace(/ة/g, 'ه')    // Normalize Teh Marbuta
        .replace(/[\u064B-\u0652]/g, ''); // Remove diacritics/tashkeel
};


/**
 * Converts Western/ASCII numerals in a string to Eastern Arabic numerals.
 */
const toArabicNumerals = (input: string | number): string => {
    const strInput = String(input);
    const arabicNumerals = ['٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩'];
    return strInput.replace(/[0-9]/g, (digit) => arabicNumerals[parseInt(digit)]);
};

/**
 * Formats a list of dates by grouping them by month and year.
 * e.g., "٥، ٨/١/٢٠٢٤ | ١٢/٢/٢٠٢٤"
 */
const formatDatesGroupedByMonth = (dates: Date[]): string => {
    if (dates.length === 0) return '';
    const datesByMonthYear: Record<string, number[]> = {};

    dates.forEach(date => {
        if (date && !isNaN(date.getTime())) {
            const month = date.getUTCMonth() + 1;
            const year = date.getUTCFullYear();
            const key = `${year}-${String(month).padStart(2, '0')}`;
            if (!datesByMonthYear[key]) {
                datesByMonthYear[key] = [];
            }
            datesByMonthYear[key].push(date.getUTCDate());
        }
    });

    const formattedDateParts = Object.keys(datesByMonthYear).sort().map(key => {
        const [year, month] = key.split('-');
        const days = datesByMonthYear[key];
        days.sort((a, b) => a - b);
        const arabicDays = days.map(d => toArabicNumerals(d)).join('،');
        const arabicMonth = toArabicNumerals(parseInt(month, 10));
        const arabicYear = toArabicNumerals(year);
        return `${arabicDays}/${arabicMonth}/${arabicYear}`;
    });

    return formattedDateParts.join(' | ');
};

/**
 * Formats a list of dates into ranges for consecutive days.
 * e.g., "١٠-١٢/٣/٢٠٢٤ | ١٥/٣/٢٠٢٤"
 */
const formatDateRanges = (dates: Date[]): string => {
    if (dates.length === 0) return '';
    
    const sortedDates = dates.filter(d => d && !isNaN(d.getTime())).sort((a, b) => a.getTime() - b.getTime());
    if (sortedDates.length === 0) return '';

    const periods: { start: Date, end: Date }[] = [];
    let currentPeriod = { start: sortedDates[0], end: sortedDates[0] };

    for (let i = 1; i < sortedDates.length; i++) {
        const currentDate = sortedDates[i];
        const prevDate = currentPeriod.end;
        const diffDays = Math.round((currentDate.getTime() - prevDate.getTime()) / (1000 * 60 * 60 * 24));

        if (diffDays === 1) {
            currentPeriod.end = currentDate;
        } else {
            periods.push(currentPeriod);
            currentPeriod = { start: currentDate, end: currentDate };
        }
    }
    periods.push(currentPeriod);

    return periods.map(period => {
        const startDay = period.start.getUTCDate();
        const startMonth = period.start.getUTCMonth() + 1;
        const startYear = period.start.getUTCFullYear();

        if (period.start.getTime() === period.end.getTime()) {
            return `${toArabicNumerals(startDay)}/${toArabicNumerals(startMonth)}/${toArabicNumerals(startYear)}`;
        } else {
            const endDay = period.end.getUTCDate();
            const endMonth = period.end.getUTCMonth() + 1;
            const endYear = period.end.getUTCFullYear();

            if (startYear === endYear && startMonth === endMonth) {
                return `${toArabicNumerals(startDay)}-${toArabicNumerals(endDay)}/${toArabicNumerals(startMonth)}/${toArabicNumerals(startYear)}`;
            } else if (startYear === endYear) {
                return `${toArabicNumerals(startDay)}/${toArabicNumerals(startMonth)} - ${toArabicNumerals(endDay)}/${toArabicNumerals(endMonth)}/${toArabicNumerals(startYear)}`;
            } else {
                return `${toArabicNumerals(startDay)}/${toArabicNumerals(startMonth)}/${toArabicNumerals(startYear)} - ${toArabicNumerals(endDay)}/${toArabicNumerals(endMonth)}/${toArabicNumerals(endYear)}`;
            }
        }
    }).join(' | ');
};

/**
 * Generates a random alphanumeric password.
 */
const generateRandomPassword = (length: number = 7): string => {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
};

/**
 * Generates a random two-digit username.
 */
const generateRandomUsername = (): string => {
  return String(Math.floor(Math.random() * 100)).padStart(2, '0');
};

/**
 * Converts a file to a base64 encoded string.
 */
const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result as string);
        reader.onerror = error => reject(error);
    });
};


/**
 * Formats a number of days into a grammatically correct Arabic string.
 */
const formatDaysArabic = (count: number): string => {
    const num = Number(count);
    if (isNaN(num)) return ' ';
    if (num === 0) return '٠ يوم';
    if (num === 1) return 'يوم واحد';
    if (num === 2) return 'يومان';
    if (num >= 3 && num <= 10) {
        const words = ['ثلاثة', 'أربعة', 'خمسة', 'ستة', 'سبعة', 'ثمانية', 'تسعة', 'عشرة'];
        return `${words[num - 3]} أيام`;
    }
    // For numbers > 10, the format is [number] + singular noun (accusative)
    return `${toArabicNumerals(num)} يومًا`;
};

/**
 * Formats leave counts (days and hours) into an Arabic string with Arabic numerals.
 */
const formatLeaveCount = (days: number, hours: number): string => {
    const parts: string[] = [];
    if (days > 0) parts.push(`${toArabicNumerals(days)} يوم`);
    const formattedHours = parseFloat(hours.toFixed(2));
    if (formattedHours > 0) parts.push(`${toArabicNumerals(formattedHours)} ساعة`);
    return parts.length > 0 ? parts.join(' و ') : 'لا يوجد';
};


/**
 * Processes raw data from SheetJS to handle messy and unstructured Excel files.
 * It identifies columns by content, not just headers, and fills in missing employee names
 * for consecutive rows.
 */
const processAndCleanData = (rawData: DataRow[]): { cleanData: DataRow[], cleanHeaders: string[] } => {
    if (rawData.length < 1) {
        return { cleanData: [], cleanHeaders: [] };
    }

    const daysOfWeek = ['الاحد', 'الاثنين', 'الثلاثاء', 'الاربعاء', 'الخميس', 'الجمعة', 'السبت'];
    const leaveTypesKeywords = ['اجازة', 'زمنية', 'رصد'];
    const headers = Object.keys(rawData[0] || {});

    // Step 1: Score each column based on content matching
    const headerScores: Record<string, Record<string, number>> = {};
    headers.forEach(h => {
        headerScores[h] = { name: 0, date: 0, day: 0, type: 0, value: 0, id: 0, other: 0 };
    });

    const sampleSize = Math.min(rawData.length, 50); // Use a larger sample size for better accuracy

    for (let i = 0; i < sampleSize; i++) {
        const row = rawData[i];
        for (const header of headers) {
            const val = row[header];
            if (val == null || String(val).trim() === '') continue;

            const valueStr = String(val).trim();

            if (val instanceof Date || valueStr.match(/^\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}$/)) {
                headerScores[header].date++;
            } else if (daysOfWeek.includes(valueStr)) {
                headerScores[header].day++;
            } else if (leaveTypesKeywords.some(keyword => valueStr.toLowerCase().includes(keyword))) {
                headerScores[header].type++;
            } else if (!isNaN(parseFloat(valueStr)) && isFinite(Number(val))) {
                const num = Number(val);
                
                if (valueStr.includes('.')) {
                    headerScores[header].value += 3; 
                } else { 
                    if (num > 0 && num <= 24) {
                        headerScores[header].value++;
                    }
                    
                    if (num > 24) {
                        headerScores[header].id += 2;
                    } else if (num >= 1) {
                        headerScores[header].id++;
                    }
                }
            } else if (valueStr.includes(' ') && isNaN(Number(valueStr)) && valueStr.length > 5) {
                headerScores[header].name++;
            } else {
                headerScores[header].other++;
            }
        }
    }
    
    // Step 2: Find the best matching header for each data type, preventing duplicates
    const findBestHeader = (type: 'name' | 'date' | 'day' | 'type' | 'value', usedHeaders: Set<string>): string | null => {
        let bestHeader: string | null = null;
        let maxScore = 0;

        for (const header of headers) {
            if (usedHeaders.has(header)) continue;

            if (headerScores[header][type] > maxScore) {
                maxScore = headerScores[header][type];
                bestHeader = header;
            }
        }
        
        return maxScore > 0 ? bestHeader : null;
    };

    const usedHeaders = new Set<string>();
    const headerMapping: Record<string, string | null> = {};
    const typesToFind: ('date' | 'type' | 'name' | 'day' | 'value')[] = ['date', 'type', 'name', 'day', 'value'];

    typesToFind.forEach(type => {
        const header = findBestHeader(type, usedHeaders);
        headerMapping[type] = header;
        if (header) {
            usedHeaders.add(header);
        }
    });

    const { name: nameHeader, date: dateHeader, type: typeHeader, day: dayHeader, value: valueHeader } = headerMapping;
    
    // Step 3: Validate that required columns were found
    if (!nameHeader || !dateHeader || !typeHeader) {
        const missing = [];
        if (!nameHeader) missing.push("'الاسم'");
        if (!dateHeader) missing.push("'التاريخ'");
        if (!typeHeader) missing.push("'نوع الاجازة'");
        throw new Error(`لم نتمكن من تحديد الأعمدة التالية تلقائيًا: ${missing.join('، ')}. يرجى التأكد من أن الملف يحتوي على هذه الأعمدة.`);
    }

    // Step 4: Process and clean the data row by row
    const cleanData: DataRow[] = [];
    let currentName: string = '';

    rawData.forEach(row => {
        const name = (row[nameHeader!] ?? '').toString().trim();
        const date = row[dateHeader!];
        const type = (row[typeHeader!] ?? '').toString().trim();

        if (name) {
            currentName = name;
        }

        if (date && type) {
            let formattedDate: string;
            if (date instanceof Date) {
                const d = String(date.getDate()).padStart(2, '0');
                const m = String(date.getMonth() + 1).padStart(2, '0');
                const y = date.getFullYear();
                formattedDate = `${d}-${m}-${y}`;
            } else {
                // Handle Excel's numeric date format
                if (typeof date === 'number' && date > 1) {
                    const excelEpoch = new Date(1899, 11, 30);
                    const jsDate = new Date(excelEpoch.getTime() + date * 86400000);
                    const d = String(jsDate.getUTCDate()).padStart(2, '0');
                    const m = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
                    const y = jsDate.getUTCFullYear();
                    formattedDate = `${d}-${m}-${y}`;
                } else {
                     formattedDate = String(date);
                }
            }

            const newRow: DataRow = {
                'الاسم': currentName,
                'التاريخ': formattedDate,
                'يوم العمل': dayHeader && row[dayHeader] ? row[dayHeader] : ' ',
                'نوع الاجازة': type,
                'القيمة': valueHeader && row[valueHeader] != null ? row[valueHeader] : ' ',
            };
            cleanData.push(newRow);
        }
    });

    const cleanHeaders = ['الاسم', 'التاريخ', 'يوم العمل', 'نوع الاجازة', 'القيمة'];
    return { cleanData, cleanHeaders };
};

const leaveTypeLabels: Record<string, string> = {
    regular: 'اعتيادية',
    hourly: 'زمنية',
    sick: 'مرضية',
    long: 'طويلة',
    travel: 'سفر',
    absent: 'غائب',
};


// --- Main Application Component ---

export const App: React.FC = () => {
  const [data, setData] = useState<DataRow[]>([]);
  const [processedFiles, setProcessedFiles] = useState<string[]>([]);
  const [headers, setHeaders] = useState<string[]>(['الاسم', 'التاريخ', 'يوم العمل', 'نوع الاجازة', 'القيمة']);
  const [searchTerm, setSearchTerm] = useState('');
  const [sortConfig, setSortConfig] = useState<SortConfig | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(true); // Start as true to handle initial session load
  type ViewType = 'table' | 'summary' | 'rankedSummary' | 'employeeManagement' | 'upload' | 'monthlyReport';
  const [view, setView] = useState<ViewType>('upload');
  const [summary, setSummary] = useState<EmployeeSummary[]>([]);
  const [showCopyNotification, setShowCopyNotification] = useState<boolean>(false);
  const [rankedSortOrder, setRankedSortOrder] = useState<'byDays' | 'alphabetical'>('byDays');
  
  // --- New State for Summary Filtering ---
  const [summarySearchTerm, setSummarySearchTerm] = useState('');
  const [selectedYear, setSelectedYear] = useState('all');
  const [selectedMonth, setSelectedMonth] = useState('all');
  
  // --- New State for Monthly Report ---
  const [reportYear, setReportYear] = useState(new Date().getFullYear());
  const [reportMonth, setReportMonth] = useState(new Date().getMonth() + 1);
  const [monthlyReportData, setMonthlyReportData] = useState<MonthlyReportRow[]>([]);
  const [alphabeticalReportData, setAlphabeticalReportData] = useState<MonthlyReportRow[]>([]);
  const [showAlphabeticalReport, setShowAlphabeticalReport] = useState<boolean>(false);
  const [reportWorkplaceFilter, setReportWorkplaceFilter] = useState<string>('all');
  const [reportColumnVisibility, setReportColumnVisibility] = useState({
    regular: true,
    hourly: true,
    sick: true,
    long: true,
    travel: true,
    absent: true,
  });

  // --- Authentication State ---
  const [currentUser, setCurrentUser] = useState<EmployeeRecord | 'admin' | 'supervisor' | null>(null);
  const [usernameInput, setUsernameInput] = useState<string>('');
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [loginError, setLoginError] = useState<string | null>(null);
  const ADMIN_USERNAME = 'admin';
  const ADMIN_PASSWORD = '2781';
  const SUPERVISOR_USERNAME = '168';
  const SUPERVISOR_PASSWORD = 'ياسر78';

  // --- Employee Management State ---
  const [employees, setEmployees] = useState<EmployeeRecord[]>([]);
  const [editingEmployee, setEditingEmployee] = useState<EmployeeRecord | null>(null);
  const [newEmployeeName, setNewEmployeeName] = useState('');
  const [newEmployeeWorkplace, setNewEmployeeWorkplace] = useState('');
  const [newEmployeeBalance, setNewEmployeeBalance] = useState('');
  const [newEmployeeWorkdayHours, setNewEmployeeWorkdayHours] = useState('7');
  const [newEmployeeUsername, setNewEmployeeUsername] = useState(generateRandomUsername());
  const [newEmployeePassword, setNewEmployeePassword] = useState(generateRandomPassword());
  const [newEmployeePhoto, setNewEmployeePhoto] = useState<string | null>(null);
  const [newEmployeePhotoFile, setNewEmployeePhotoFile] = useState<File | null>(null);
  const [employeeSearchTerm, setEmployeeSearchTerm] = useState('');
  
  // --- Theme State ---
  const [theme, setTheme] = useState(localStorage.getItem(THEME_STORAGE_KEY) || 'light');
  
  // --- Data Deletion State ---
  const [uploadHistory, setUploadHistory] = useState<UploadHistoryEntry[]>([]);
  const [isDeleteMenuOpen, setIsDeleteMenuOpen] = useState(false);
  
  // --- Responsive UI State ---
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);

  const priorHourlyBalances: Record<string, number> = {
    'اثيب عبد الزهرة مجيد': 0,
    'احمد امير حسين': 2,
    'احمد سمير محمد': 2,
    'احمد ناجح رزاق': 0,
    'ازهر ثامر ربح': 0,
    'استبرق جابر حميد': 3,
    'امير خالد هادي': 3,
    'انغام عبد الزهرة مجيد': 5,
    'انوار فاضل طراد': 5,
    'باسم عباس حسين': 6,
    'باسم علي محمد': 3,
    'باقر عادل احمد': 6,
    'تقى عبد الحسن حمودي': 2,
    'جاسم حازم محمد': 0,
    'جلال حسن هادي': 0,
    'حسن كريم باجي': 0,
    'حسنين عبد الرزاق': 0,
    'حسين علي عبد الأمير': 3,
    'حكمت شوكت عبد الحمزة': 3,
    'حمزة سعد حمودي': 0,
    'حيدر عباس حسن': 1,
    'حيدر كاظم محمد سعيد': 0,
    'خليل كريم عباس': 0,
    'رائد كامل عبد اليمة': 0,
    'راضي حمودي سلطان': 1,
    'رغد عبد الزهرة مجيد': 0,
    'سجاد محمد علي طالب': 0,
    'سعد حمودي سلطان': 0,
    'سعيد عبد الحسن حمودي': 0,
    'سهاد جابر محمد': 4,
    'ضرغام جهادي خضير': 0,
    'عباس سعد حمودي': 2,
    'عباس كاظم عبد الاخوه': 5,
    'عبد الرضا نجم عبد': 0,
    'عقيل مسلم عبد المحسن': 0,
    'علي باسم حسن': 3,
    'علي رسول محي': 4,
    'علي عبد الكريم حميد': 0,
    'علي كاظم قربون': 2,
    'علي محمد باجي': 1,
    'عمار ناصر محمد': 4,
    'غسان نزار ضياء': 0,
    'قاسم محمد رضا مجيد': 4,
    'كرار عدي عبد الحسين': 5,
    'كواكب عزيز حمزة': 2,
    'كوثر هادي عطيه': 3,
    'ليث محمد علي حمودي': 0,
    'مجيد محمد رضا مجيد': 5,
    'محمد جواد كاظم عباس': 2,
    'محمد راضي حمودي': 4,
    'محمد رضا عبود': 0,
    'محمد صلاح محمد': 0,
    'محمد عبد الرضا تركي': 1,
    'محمد علاء محمد': 0,
    'محمد فاضل شاكر': 2,
    'محمد محمد رضا مجيد': 6,
    'مرتضى كريم موسى': 3,
    'مروة محمد علي': 3,
    'مسلم إبراهيم حسن': 5,
    'مصطفى راضي حمودي': 1,
    'مصطفى علي محمد': 5,
    'مصطفى محسن يعقوب': 5,
    'مها سعد حمودي': 0,
    'مهدي صالح هادي': 4,
    'مهند محمد رشاد جعفر': 3,
    'ناجح كاظم قربون': 0,
    'هدى عبد الحسن حمودي': 6,
    'هناء علي عبد الرضا': 0,
    'ياسر فائز جاسم': 0,
    'ياسر ياسين ناجي': 1,
    'ياسين رياض احمد': 2,
    'زهراء طه تقي': 0,
    'يحيى فارس محمد': 5,
  };

  const fileInputRef = useRef<HTMLInputElement>(null);
  const summaryContentRef = useRef<HTMLDivElement>(null);
  const deleteMenuRef = useRef<HTMLDivElement>(null);
  
  const initialEmployeesData: Omit<EmployeeRecord, 'id' | 'photo' | 'priorHourlyBalance'>[] = [
    { name: 'اثيب عبد الزهرة مجيد', balance: 16, workplace: 'الشعبة الثقافية', workdayHours: 7, username: '100', password: 'اثيب10' },
    { name: 'احمد امير حسين', balance: 26, workplace: 'شعبة المشاريع الهندسية', workdayHours: 7, username: '101', password: 'احمد11' },
    { name: 'احمد سمير محمد', balance: 38, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '102', password: 'احمد12' },
    { name: 'احمد ناجح رزاق', balance: 61, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '103', password: 'احمد13' },
    { name: 'ازهر ثامر ربح', balance: 36, workplace: 'شعبة الصيانة الهندسية', workdayHours: 7, username: '104', password: 'ازهر14' },
    { name: 'استبرق جابر حميد', balance: 4, workplace: 'شعبة تكنلوجيا المعلومات', workdayHours: 7, username: '105', password: 'استبرق15' },
    { name: 'امير خالد هادي', balance: 14, workplace: 'شعبة الإعلام والعلاقات العامة', workdayHours: 7, username: '106', password: 'امير16' },
    { name: 'انغام عبد الزهرة مجيد', balance: 20, workplace: 'شعبة الشؤون النسوية', workdayHours: 7, username: '107', password: 'انغام17' },
    { name: 'انوار فاضل طراد', balance: 9, workplace: 'شعبة الشؤون النسوية', workdayHours: 7, username: '108', password: 'انوار18' },
    { name: 'باسم عباس حسين', balance: 41, workplace: 'شعبة حفظ النظام', workdayHours: 7, username: '109', password: 'باسم19' },
    { name: 'باسم علي محمد', balance: 33, workplace: 'شعبة حفظ النظام', workdayHours: 7, username: '110', password: 'باسم20' },
    { name: 'باقر عادل احمد', balance: 27, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '111', password: 'باقر21' },
    { name: 'تقى عبد الحسن حمودي', balance: 22, workplace: 'شعبة الشؤون النسوية', workdayHours: 7, username: '112', password: 'تقى22' },
    { name: 'جاسم حازم محمد', balance: 13, workplace: 'شعبة الإعلام والعلاقات العامة', workdayHours: 7, username: '113', password: 'جاسم23' },
    { name: 'جلال حسن هادي', balance: 54, workplace: 'شعبة المخازن', workdayHours: 7, username: '114', password: 'جلال24' },
    { name: 'حسن كريم باجي', balance: 44, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '115', password: 'حسن25' },
    { name: 'حسنين عبد الرزاق', balance: 67, workplace: 'شعبة المشاريع الهندسية', workdayHours: 7, username: '116', password: 'حسنين26' },
    { name: 'حسين علي عبد الأمير', balance: 17, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '117', password: 'حسين27' },
    { name: 'حكمت شوكت عبد الحمزة', balance: 26, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '118', password: 'حكمت28' },
    { name: 'حمزة سعد حمودي', balance: 27, workplace: 'شعبة الإعلام والعلاقات العامة', workdayHours: 7, username: '119', password: 'حمزة29' },
    { name: 'حيدر عباس حسن', balance: 48, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '120', password: 'حيدر30' },
    { name: 'حيدر كاظم محمد سعيد', balance: 67, workplace: 'المعاون المالي', workdayHours: 7, username: '121', password: 'حيدر31' },
    { name: 'خليل كريم عباس', balance: 51, workplace: 'شعبة الصيانة الهندسية', workdayHours: 7, username: '122', password: 'خليل32' },
    { name: 'رائد كامل عبد اليمة', balance: 21, workplace: 'شعبة الشؤون المالية', workdayHours: 7, username: '123', password: 'رائد33' },
    { name: 'راضي حمودي sultan', balance: 39, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '124', password: 'راضي34' },
    { name: 'رغد عبد الزهرة مجيد', balance: 15, workplace: 'شعبة الشؤون النسوية', workdayHours: 7, username: '125', password: 'رغد35' },
    { name: 'زهراء طه تقي', balance: 12, workplace: 'شعبة الشؤون النسوية', workdayHours: 7, username: '171', password: 'زهراء81' },
    { name: 'سجاد محمد علي طالب', balance: 10, workplace: 'شعبة الصيانة الهندسية', workdayHours: 7, username: '126', password: 'سجاد36' },
    { name: 'سعد حمودي سلطان', balance: 27, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '127', password: 'سعد37' },
    { name: 'سعيد عبد الحسن حمودي', balance: 51, workplace: 'نائب الأمين الخاص', workdayHours: 7, username: '128', password: 'سعيد38' },
    { name: 'سهاد جابر محمد', balance: 4, workplace: 'الشعبة الثقافية', workdayHours: 7, username: '129', password: 'سهاد39' },
    { name: 'ضرغام جهادي خضير', balance: 19, workplace: 'المعاون الثقافي', workdayHours: 7, username: '130', password: 'ضرغام40' },
    { name: 'عباس سعد حمودي', balance: 18, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '131', password: 'عباس41' },
    { name: 'عباس كاظم عبد الاخوه', balance: 9, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '132', password: 'عباس42' },
    { name: 'عبد الرضا نجم عبد', balance: 24, workplace: 'شعبة الصيانة الهندسية', workdayHours: 7, username: '133', password: 'عبد43' },
    { name: 'عقيل مسلم عبد المحسن', balance: 6, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '134', password: 'عقيل44' },
    { name: 'علي باسم حسن', balance: 5, workplace: 'شعبة التدقيق', workdayHours: 7, username: '135', password: 'علي45' },
    { name: 'علي رسول محي', balance: 25, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '136', password: 'علي46' },
    { name: 'علي عبد الكريم حميد', balance: 21, workplace: 'الشعبة الثقافية', workdayHours: 7, username: '137', password: 'علي47' },
    { name: 'علي كاظم قربون', balance: 30, workplace: 'شعبة حفظ النظام', workdayHours: 7, username: '138', password: 'علي48' },
    { name: 'علي محمد باجي', balance: 24, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '139', password: 'علي49' },
    { name: 'عمار ناصر محمد', balance: -1, workplace: 'الشعبة الثقافية', workdayHours: 7, username: '140', password: 'عمار50' },
    { name: 'غسان نزار ضياء', balance: 21, workplace: 'الشعبة القانونية', workdayHours: 7, username: '141', password: 'غسان51' },
    { name: 'قاسم محمد رضا مجيد', balance: 6, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '142', password: 'قاسم52' },
    { name: 'كرار عدي عبد الحسين', balance: 0, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '143', password: 'كرار53' },
    { name: 'كواكب عزيز حمزة', balance: 13, workplace: 'شعبة الشؤون النسوية', workdayHours: 7, username: '144', password: 'كواكب54' },
    { name: 'كوثر هادي عطيه', balance: -13, workplace: 'شعبة الشؤون النسوية', workdayHours: 7, username: '145', password: 'كوثر55' },
    { name: 'ليث محمد علي حمودي', balance: 27, workplace: 'شعبة حفظ النظام', workdayHours: 7, username: '146', password: 'ليث56' },
    { name: 'مجيد محمد رضا مجيد', balance: 16, workplace: 'شعبة المخازن', workdayHours: 7, username: '147', password: 'مجيد57' },
    { name: 'محمد جواد كاظم عباس', balance: 53, workplace: 'شعبة الصيانة الهندسية', workdayHours: 7, username: '148', password: 'محمد58' },
    { name: 'محمد راضي حمودي', balance: 35, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '149', password: 'محمد59' },
    { name: 'محمد رضا عبود', balance: 51, workplace: 'المعاون الإداري', workdayHours: 7, username: '150', password: 'محمد60' },
    { name: 'محمد صلاح محمد', balance: 54, workplace: 'شعبة الإعلام والعلاقات العامة', workdayHours: 7, username: '151', password: 'محمد61' },
    { name: 'محمد عبد الرضا تركي', balance: 19, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '152', password: 'محمد62' },
    { name: 'محمد علاء محمد', balance: 67, workplace: 'شعبة المشاريع الهندسية', workdayHours: 7, username: '153', password: 'محمد63' },
    { name: 'محمد فاضل شاكر', balance: 5, workplace: 'شعبة الشؤون الإدارية', workdayHours: 7, username: '154', password: 'محمد64' },
    { name: 'محمد محمد رضا مجيد', balance: 22, workplace: 'الشعبة الثقافية', workdayHours: 7, username: '155', password: 'محمد65' },
    { name: 'مرتضى كريم موسى', balance: 1, workplace: 'شعبة حفظ النظام', workdayHours: 7, username: '156', password: 'مرتضى66' },
    { name: 'مروة محمد علي', balance: 10, workplace: 'شعبة حفظ النظام', workdayHours: 7, username: '157', password: 'مروة67' },
    { name: 'مسلم إبراهيم حسن', balance: 26, workplace: 'شعبة حفظ النظام', workdayHours: 7, username: '158', password: 'مسلم68' },
    { name: 'مصطفى راضي حمودي', balance: -1, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '159', password: 'مصطفى69' },
    { name: 'مصطفى علي محمد', balance: 16, workplace: 'شعبة الصيانة الهندسية', workdayHours: 7, username: '160', password: 'مصطفى70' },
    { name: 'مصطفى محسن يعقوب', balance: 14, workplace: 'الشعبة القانونية', workdayHours: 7, username: '161', password: 'مصطفى71' },
    { name: 'مها سعد حمودي', balance: 11, workplace: 'شعبة الشؤون النسوية', workdayHours: 7, username: '162', password: 'مها72' },
    { name: 'مهدي صالح هادي', balance: 14, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '163', password: 'مهدي73' },
    { name: 'مهند محمد رشاد جعفر', balance: 6, workplace: 'شعبة الصيانة الهندسية', workdayHours: 7, username: '164', password: 'مهند74' },
    { name: 'ناجح كاظم قربون', balance: 56, workplace: 'شعبة حفظ النظام', workdayHours: 7, username: '165', password: 'ناجح75' },
    { name: 'هدى عبد الحسن حمودي', balance: 28, workplace: 'شعبة حفظ النظام', workdayHours: 7, username: '166', password: 'هدى76' },
    { name: 'هناء علي عبد الرضا', balance: 18, workplace: 'الشعبة الثقافية', workdayHours: 7, username: '167', password: 'هناء77' },
    { name: 'ياسر فائز جاسم', balance: 37, workplace: 'شعبة الشؤون الإدارية', workdayHours: 7, username: '168', password: 'ياسر78' },
    { name: 'ياسر ياسين ناجي', balance: -9, workplace: 'الشعبة الثقافية', workdayHours: 7, username: '169', password: 'ياسر79' },
    { name: 'ياسين رياض احمد', balance: -6, workplace: 'شعبة الشؤون الإدارية', workdayHours: 7, username: '170', password: 'ياسين80' },
    { name: 'يحيى فارس محمد', balance: 0, workplace: 'شعبة الشؤون الخدمية', workdayHours: 7, username: '172', password: 'يحيى82' },
  ];

  // --- Theme Management ---
  useEffect(() => {
    const root = window.document.documentElement;
    if (theme === 'dark') {
      root.classList.add('dark');
    } else {
      root.classList.remove('dark');
    }
    localStorage.setItem(THEME_STORAGE_KEY, theme);
  }, [theme]);

  const toggleTheme = () => {
    setTheme(prevTheme => (prevTheme === 'light' ? 'dark' : 'light'));
  };

  // --- Data Persistence ---

  // Load all data from localStorage on initial mount
  useEffect(() => {
    const initializeApp = async () => {
        try {
            // --- Session Restoration ---
            const storedUserJson = localStorage.getItem(CURRENT_USER_STORAGE_KEY);
            if (storedUserJson) {
                const storedUser = JSON.parse(storedUserJson);
                if (storedUser === 'admin' || storedUser === 'supervisor') {
                    setCurrentUser(storedUser);
                } else if (typeof storedUser === 'object' && storedUser.id) {
                    // Fetch latest data on reload to ensure sync with DB
                    const firestoreData = await getEmployeeData(storedUser.id);
                    setCurrentUser(firestoreData || storedUser); // Use fresh data or fallback to stored
                }
            }

            // Employees
            let storedEmployees = localStorage.getItem(EMPLOYEES_STORAGE_KEY);
            if (!storedEmployees || JSON.parse(storedEmployees).length === 0) {
                const initialEmployees: EmployeeRecord[] = initialEmployeesData.map((emp, index) => {
                   const priorHours = priorHourlyBalances[emp.name] ?? 0;
                   return {
                     ...emp,
                     id: `${Date.now()}-${index}`,
                     priorHourlyBalance: priorHours,
                   };
                }).sort((a, b) => a.name.localeCompare(b.name, 'ar'));
                
                storedEmployees = JSON.stringify(initialEmployees);
                localStorage.setItem(EMPLOYEES_STORAGE_KEY, storedEmployees);
            }
            setEmployees(JSON.parse(storedEmployees));

            // Leave Data
            const storedData = localStorage.getItem(LEAVE_DATA_STORAGE_KEY);
            const parsedData = storedData ? JSON.parse(storedData) : [];
            setData(parsedData);

            // Processed Files
            const storedFiles = localStorage.getItem(PROCESSED_FILES_STORAGE_KEY);
            setProcessedFiles(storedFiles ? JSON.parse(storedFiles) : []);
            
            // Upload History
            const storedHistory = localStorage.getItem(UPLOAD_HISTORY_KEY);
            setUploadHistory(storedHistory ? JSON.parse(storedHistory) : []);
            
            // Set initial view based on loaded data
            if (parsedData.length > 0 && currentUser && (currentUser === 'admin' || currentUser === 'supervisor')) {
                setView('summary');
            }
        } catch (err) {
            console.error("Failed to load data from localStorage", err);
            setError("فشل تحميل البيانات المحفوظة.");
        } finally {
            setIsLoading(false); // Stop loading after initialization
        }
    };
    
    initializeApp();
  }, []); // Empty dependency array means this runs only once on mount

  // Save data to localStorage whenever it changes
  useEffect(() => {
    try {
      localStorage.setItem(EMPLOYEES_STORAGE_KEY, JSON.stringify(employees));
    } catch (err) {
      console.error("Failed to save employees to localStorage", err);
    }
  }, [employees]);
  
  useEffect(() => {
    try {
      localStorage.setItem(LEAVE_DATA_STORAGE_KEY, JSON.stringify(data));
    } catch (err) {
      console.error("Failed to save leave data to localStorage", err);
    }
  }, [data]);

  useEffect(() => {
    try {
      localStorage.setItem(PROCESSED_FILES_STORAGE_KEY, JSON.stringify(processedFiles));
    } catch (err) {
      console.error("Failed to save processed files list to localStorage", err);
    }
  }, [processedFiles]);
  
  useEffect(() => {
    try {
      localStorage.setItem(UPLOAD_HISTORY_KEY, JSON.stringify(uploadHistory));
    } catch (err) {
      console.error("Failed to save upload history to localStorage", err);
    }
  }, [uploadHistory]);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (deleteMenuRef.current && !deleteMenuRef.current.contains(event.target as Node)) {
        setIsDeleteMenuOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);
  
  const parseDate = useCallback((dateStr: string): Date => {
      const parts = dateStr.split(/[-/]/);
      if (parts.length === 3) {
          const [day, month, year] = parts.map(p => parseInt(p, 10));
          if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
              // Handle DD-MM-YYYY or similar formats. Assume year > 1000 is 4-digit.
              if (year > 1000) {
                  return new Date(Date.UTC(year, month - 1, day));
              }
              // Handle 2-digit years like '24' for 2024
              const fullYear = year > 50 ? 1900 + year : 2000 + year;
              return new Date(Date.UTC(fullYear, month - 1, day));
          }
      }
      return new Date(0); // Invalid date
  }, []);

  const getRegularAndTimeBasedTotalDays = useCallback((employee: EmployeeSummary): number => {
    let total = 0;
    employee.leaves.forEach(l => {
        if (l.type === 'اجازة اعتيادية' || l.type === 'ملخص الزمنيات') {
            total += l.dayCount;
        }
    });
    return total;
  }, []);

  const generateSummary = useCallback((summaryData: DataRow[], employeeRecords: EmployeeRecord[]): EmployeeSummary[] => {
    const normalizedEmployeeMap = new Map<string, EmployeeRecord>();
    employeeRecords.forEach(emp => {
        normalizedEmployeeMap.set(normalizeName(emp.name), emp);
    });

    const nameResolutionCache = new Map<string, string | null>();

    const resolveCanonicalName = (sheetName: string): string | null => {
        if (!sheetName) return null;
        const cached = nameResolutionCache.get(sheetName);
        if (cached !== undefined) {
            return cached;
        }

        const normalizedSheetName = normalizeName(sheetName);

        if (normalizedEmployeeMap.has(normalizedSheetName)) {
            const canonicalName = normalizedEmployeeMap.get(normalizedSheetName)!.name;
            nameResolutionCache.set(sheetName, canonicalName);
            return canonicalName;
        }

        const potentialMatches = employeeRecords.filter(emp => {
            const normalizedEmpName = normalizeName(emp.name);
            return normalizedEmpName.includes(normalizedSheetName) || normalizedSheetName.includes(normalizedEmpName);
        });

        if (potentialMatches.length === 1) {
            const canonicalName = potentialMatches[0].name;
            nameResolutionCache.set(sheetName, canonicalName);
            return canonicalName;
        } else if (potentialMatches.length > 1) {
            // If multiple matches, prefer the one with the closest length
            potentialMatches.sort((a, b) =>
                Math.abs(normalizeName(a.name).length - normalizedSheetName.length) -
                Math.abs(normalizeName(b.name).length - normalizedSheetName.length)
            );
            const canonicalName = potentialMatches[0].name;
            nameResolutionCache.set(sheetName, canonicalName);
            return canonicalName;
        }

        nameResolutionCache.set(sheetName, null);
        return null;
    };

    const groupedData: Record<string, Record<string, { date: string, value: number }[]>> = {};

    summaryData.forEach(row => {
        const nameFromSheet = row['الاسم'];
        if (!nameFromSheet) return;

        const canonicalName = resolveCanonicalName(nameFromSheet);
        const date = row['التاريخ'];
        let type = (row['نوع الاجازة'] || '').toString().trim();
        const value = parseFloat(String(row['القيمة']).trim()) || 0;

        if (!canonicalName || !date || !type) {
            return;
        }

        if (type.startsWith('زمنية')) {
            type = 'اجازة زمنية';
        } else {
            type = type.replace(/\s+/g, ' ').trim();
        }

        if (type.includes('رصد') && type.includes('مسائي')) {
            return; // Skip evening observation entries
        }

        if (!groupedData[canonicalName]) groupedData[canonicalName] = {};
        if (!groupedData[canonicalName][type]) groupedData[canonicalName][type] = [];

        groupedData[canonicalName][type].push({ date: String(date), value });
    });

    const allNames = new Set([...Object.keys(groupedData), ...employeeRecords.map(e => e.name)]);
    const sortedNames = Array.from(allNames).sort((a, b) => a.localeCompare(b, 'ar'));
    
    const finalSummary: EmployeeSummary[] = [];

    for (const name of sortedNames) {
        const employeeRecord = employeeRecords.find(emp => emp.name === name);
        const initialBalance = employeeRecord?.balance ?? 0;

        const employeeSummary: EmployeeSummary = {
            name,
            leaves: [],
            initialBalance: initialBalance,
            photo: employeeRecord?.photo,
            workplace: employeeRecord?.workplace
        };
        const employeeVacations = groupedData[name] || {};
        
        const priorHourlyBalance = employeeRecord?.priorHourlyBalance || 0;
        
        let workdayHours = employeeRecord?.workdayHours || 7;
        const regularLeavesForHoursCheck = employeeVacations['اجازة اعتيادية'] || [];
        if (regularLeavesForHoursCheck.length > 0) {
            const typicalHours = regularLeavesForHoursCheck[0].value;
            if (typicalHours === 6 || typicalHours === 7) {
                workdayHours = typicalHours;
            }
        }
        
        const leaveTypesToProcess = new Set(Object.keys(employeeVacations));
        if (priorHourlyBalance > 0 && summaryData.length > 0) {
            leaveTypesToProcess.add('اجازة زمنية');
        }

        const sortedTypes = Array.from(leaveTypesToProcess).sort((a, b) => {
            const order = ['اجازة اعتيادية', 'اجازة مرضية', 'ملخص الزمنيات'];
            const indexA = order.indexOf(a);
            const indexB = order.indexOf(b);
            if(indexA > -1 && indexB > -1) return indexA - indexB;
            if(indexA > -1) return -1;
            if(indexB > -1) return 1;
            return a.localeCompare(b, 'ar');
        });

        for (const type of sortedTypes) {
            const entries = employeeVacations[type] || [];

            if (type === 'اجازة اعتيادية') {
                const entriesByMonth: Record<string, { date: string; value: number }[]> = {};
                entries.forEach(entry => {
                    const date = parseDate(String(entry.date));
                    if (!isNaN(date.getTime())) {
                        const monthKey = `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
                        if (!entriesByMonth[monthKey]) {
                            entriesByMonth[monthKey] = [];
                        }
                        entriesByMonth[monthKey].push(entry);
                    }
                });
        
                Object.keys(entriesByMonth).sort().forEach(monthKey => {
                    const monthlyEntries = entriesByMonth[monthKey];
                    const count = monthlyEntries.length;
                    const dates = monthlyEntries.map(e => parseDate(String(e.date)));
                    const dateDetailsStr = formatDatesGroupedByMonth(dates);
        
                    employeeSummary.leaves.push({
                        type,
                        dayCount: count,
                        hourCount: 0,
                        dateDetails: dateDetailsStr,
                        monthYear: monthKey,
                    });
                });
            } else if (type === 'اجازة زمنية') {
                const totalHoursFromSheet = entries.reduce((sum, entry) => sum + entry.value, 0);
                const totalHours = totalHoursFromSheet + (summaryData.length > 0 ? priorHourlyBalance : 0);


                if (totalHours > 0) {
                    const days = Math.floor(totalHours / workdayHours);
                    const remainingHours = totalHours % workdayHours;
                    const formattedHours = parseFloat(remainingHours.toFixed(2));
                    
                    let dateDetails = '';
                    if (entries.length > 0) {
                        dateDetails = `إجمالي ${toArabicNumerals(totalHoursFromSheet)} ساعة عبر ${toArabicNumerals(entries.length)} إدخال`;
                    }
                    if (priorHourlyBalance > 0 && summaryData.length > 0) {
                         if (dateDetails) dateDetails += ' + ';
                         dateDetails += `${toArabicNumerals(priorHourlyBalance)} ساعة رصيد سابق`;
                    }

                    employeeSummary.leaves.push({
                        type: 'ملخص الزمنيات',
                        dayCount: days,
                        hourCount: formattedHours,
                        dateDetails: dateDetails
                    });
                }
            } else if (type.includes('مرضية') || type.includes('طويلة')) {
                const sortedDates = entries
                    .map(e => parseDate(String(e.date)))
                    .filter(d => !isNaN(d.getTime()))
                    .sort((a, b) => a.getTime() - b.getTime());
                
                if (sortedDates.length > 0) {
                   const dateDetailsStr = formatDateRanges(sortedDates);
                   const dayCount = sortedDates.length;

                   employeeSummary.leaves.push({
                        type,
                        dayCount: dayCount,
                        hourCount: 0,
                        dateDetails: dateDetailsStr
                    });
                }
            } else {
                const count = entries.length;
                if (count > 0) {
                    const dateDetailsStr = formatDatesGroupedByMonth(entries.map(e => parseDate(String(e.date))));
                    
                    employeeSummary.leaves.push({
                        type,
                        dayCount: count,
                        hourCount: 0,
                        dateDetails: dateDetailsStr
                    });
                }
            }
        }

        const totalDeducted = getRegularAndTimeBasedTotalDays(employeeSummary);
        employeeSummary.currentBalance = initialBalance - totalDeducted;
        
        if (employeeSummary.leaves.length > 0 || summaryData.length === 0) {
          finalSummary.push(employeeSummary);
        }
    }

    return finalSummary;
  }, [parseDate, getRegularAndTimeBasedTotalDays]);


  // Regenerate summary whenever data or employees change
  useEffect(() => {
    if (employees.length > 0) {
        const newSummary = generateSummary(data, employees);
        setSummary(newSummary);
    }
  }, [data, employees, generateSummary]);

  /**
   * Heuristically parses raw text (e.g., from a PDF) into structured leave data rows.
   * This function is experimental and relies on the text format having the employee's full name,
   * a date, and a leave-type keyword on the same or subsequent lines.
   */
  const parseTextToDataRows = useCallback((text: string): { cleanData: DataRow[], cleanHeaders: string[] } => {
    const cleanData: DataRow[] = [];
    const lines = text.split('\n').filter(line => line.trim() !== '');

    const dateRegex = /(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/;
    const leaveKeywords = ['اعتيادية', 'مرضية', 'زمنية', 'طويلة', 'سفر', 'غائب'];

    let currentName: string | null = null;

    lines.forEach(line => {
      // Find if a known employee name is in the line
      const foundEmployee = employees.find(emp => line.includes(emp.name));
      if (foundEmployee) {
        currentName = foundEmployee.name;
      }

      const dateMatch = line.match(dateRegex);
      if (currentName && dateMatch) {
        const date = dateMatch[0];

        let leaveType = '';
        for (const keyword of leaveKeywords) {
          if (line.includes(keyword)) {
            leaveType = keyword === 'زمنية' ? 'اجازة زمنية' : `اجازة ${keyword}`;
            break;
          }
        }

        if (leaveType) {
          let value: string | number = '';
          if (leaveType.includes('زمنية')) {
            const hourMatch = line.match(/(\d{1,2}(?:\.\d{1,2})?)\s*(ساعة|ساعات)/);
            if (hourMatch) {
              value = parseFloat(hourMatch[1]);
            }
          }

          cleanData.push({
            'الاسم': currentName,
            'التاريخ': date.trim(),
            'يوم العمل': '',
            'نوع الاجازة': leaveType.trim(),
            'القيمة': value,
          });
        }
      }
    });

    const cleanHeaders = ['الاسم', 'التاريخ', 'يوم العمل', 'نوع الاجازة', 'القيمة'];
    return { cleanData, cleanHeaders };
  }, [employees]);


  const handlePdfFile = useCallback(async (file: File) => {
    setIsLoading(true);
    setError(null);
    try {
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
      let fullText = '';
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        // FIX: The type of textContent.items can be 'unknown' or not an array. Added a robust check.
        if (textContent?.items && Array.isArray(textContent.items)) {
          fullText += textContent.items.map((item: any) => (item && item.str) || '').join(' ');
        }
        fullText += '\n\n'; // Add newline between pages
      }
      
      const { cleanData, cleanHeaders } = parseTextToDataRows(fullText);
      
      if (cleanData.length === 0) {
        throw new Error("لم يتم العثور على بيانات إجازات قابلة للمعالجة في ملف PDF. قد يكون تنسيق الملف غير متوافق.");
      }
      
      const uploadId = Date.now();
      const dataWithId = cleanData.map(row => ({ ...row, uploadId }));
      
      setUploadHistory(prev => [...prev, { fileName: file.name, uploadId }]);
      setData(prevData => [...prevData, ...dataWithId]);
      setHeaders(cleanHeaders);
      setProcessedFiles(prev => [...prev, file.name]);
      
      setView('summary');
    } catch (err) {
      console.error("Error processing PDF file:", err);
      const errorMessage = err instanceof Error ? err.message : "حدث خطأ أثناء معالجة ملف PDF.";
      setError(errorMessage);
    } finally {
      setIsLoading(false);
    }
  }, [parseTextToDataRows]);

  const handleExcelFile = useCallback((file: File) => {
    setIsLoading(true);
    setError(null);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const fileData = e.target?.result;
        const workbook = XLSX.read(fileData, { type: 'binary', cellDates: true });
        
        let allRawData: DataRow[] = [];
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const sheetData: DataRow[] = XLSX.utils.sheet_to_json(worksheet, { raw: true, defval: null });
          allRawData = allRawData.concat(sheetData);
        });

        if (allRawData.length > 0) {
          const { cleanData, cleanHeaders } = processAndCleanData(allRawData);
          
          if(cleanData.length === 0){
             throw new Error("لم يتم العثور على بيانات صالحة في الملف بعد المعالجة.");
          }
          
          const uploadId = Date.now();
          const dataWithId = cleanData.map(row => ({ ...row, uploadId }));

          setUploadHistory(prev => [...prev, { fileName: file.name, uploadId }]);
          setData(prevData => [...prevData, ...dataWithId]);
          setHeaders(cleanHeaders);
          setProcessedFiles(prev => [...prev, file.name]);
          
          setView('summary');
        } else {
            setError("الملف فارغ أو لا يحتوي على بيانات.");
        }
      } catch (err) {
        console.error("Error processing Excel file:", err);
        const errorMessage = err instanceof Error ? err.message : "حدث خطأ أثناء معالجة الملف. يرجى التأكد من أنه ملف Excel صالح.";
        setError(errorMessage);
      } finally {
        setIsLoading(false);
      }
    };
    reader.onerror = () => {
        setError("فشل في قراءة الملف.");
        setIsLoading(false);
    }
    reader.readAsBinaryString(file);
  }, []);

  const handleFile = useCallback((file: File) => {
    if (processedFiles.includes(file.name)) {
        setError(`تم تحميل هذا الملف "${file.name}" مسبقًا.`);
        return;
    }

    const fileName = file.name.toLowerCase();

    if (fileName.endsWith('.pdf')) {
        handlePdfFile(file);
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.csv')) {
        handleExcelFile(file);
    } else {
        setError("نوع الملف غير مدعوم. يرجى تحميل ملف Excel أو PDF.");
    }
  }, [processedFiles, handleExcelFile, handlePdfFile]);
  
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file);
    e.target.value = '';
  };
  
  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => e.preventDefault();
  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  };

  const requestSort = (key: string) => {
    let direction: 'ascending' | 'descending' = 'ascending';
    if (sortConfig?.key === key && sortConfig.direction === 'ascending') {
      direction = 'descending';
    }
    setSortConfig({ key, direction });
  };

  const processedData = useMemo(() => {
    let filteredData = [...data];
    if (searchTerm) {
      filteredData = filteredData.filter(row =>
        Object.values(row).some(value =>
          String(value).toLowerCase().includes(searchTerm.toLowerCase())
        )
      );
    }
    if (sortConfig !== null) {
      filteredData.sort((a, b) => {
        const valA = a[sortConfig.key];
        const valB = b[sortConfig.key];
        if (valA == null || valA === '') return 1;
        if (valB == null || valB === '') return -1;
        if (valA < valB) return sortConfig.direction === 'ascending' ? -1 : 1;
        if (valA > valB) return sortConfig.direction === 'ascending' ? 1 : -1;
        return 0;
      });
    }
    return filteredData;
  }, [data, searchTerm, sortConfig]);

  const handleDownloadCSV = useCallback(() => {
    if (processedData.length === 0) return;
    const worksheet = XLSX.utils.json_to_sheet(processedData);
    const csvOutput: string = XLSX.utils.sheet_to_csv(worksheet);
    const blob = new Blob(['\uFEFF' + csvOutput], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    const fileName = `data_export_${new Date().toISOString().split('T')[0]}.csv`;
    link.setAttribute('download', fileName);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }, [processedData]);

  const handleClearAllData = () => {
    if (window.confirm("هل أنت متأكد من حذف جميع البيانات؟ سيتم مسح بيانات الموظفين وبيانات الإجازات المحملة بشكل دائم.")) {
        localStorage.removeItem(EMPLOYEES_STORAGE_KEY);
        localStorage.removeItem(LEAVE_DATA_STORAGE_KEY);
        localStorage.removeItem(PROCESSED_FILES_STORAGE_KEY);
        localStorage.removeItem(LAST_BALANCE_UPDATE_KEY);
        localStorage.removeItem(UPLOAD_HISTORY_KEY);
        window.location.reload();
    }
  };

  const handleDeleteUploadById = (uploadId: number, fileName: string) => {
    setIsLoading(true);
    try {
        setData(prevData => prevData.filter(row => row.uploadId !== uploadId));
        setProcessedFiles(prevFiles => prevFiles.filter(name => name !== fileName));
        setUploadHistory(prevHistory => prevHistory.filter(h => h.uploadId !== uploadId));
    } catch(err) {
        console.error("Error deleting upload:", err);
        setError("حدث خطأ أثناء حذف الملف.");
    } finally {
        setIsLoading(false);
    }
  };

  const handleDeleteLastUpload = () => {
    if (uploadHistory.length === 0) {
        alert("لا توجد عمليات تحميل لحذفها.");
        return;
    }
    const lastUpload = uploadHistory[uploadHistory.length - 1];
    handleDeleteUploadById(lastUpload.uploadId, lastUpload.fileName);
  };
  
  const handleResetBalances = () => {
    if (window.confirm("هل أنت متأكد من إعادة تعيين جميع أرصدة الموظفين إلى قيمها الأولية؟ سيؤدي هذا إلى التراجع عن أي تعديلات يدوية على الأرصدة وقد يحل مشاكل الحسابات الخاطئة.")) {
        const updatedEmployees = employees.map(currentEmp => {
            const initialData = initialEmployeesData.find(initialEmp => normalizeName(initialEmp.name) === normalizeName(currentEmp.name));
            if (initialData) {
                // Return a new object with the original balance, preserving other details
                return {
                    ...currentEmp,
                    balance: initialData.balance,
                };
            }
            // If the employee isn't in the initial hardcoded list, leave them as is
            return currentEmp;
        });

        setEmployees(updatedEmployees);
        alert("تمت إعادة تعيين أرصدة الموظفين إلى قيمها الأولية.");
    }
  };


  const handlePrint = () => {
    window.print();
  };

  const handleExportPDF = async () => {
    if (isLoading) return;
    setIsLoading(true);
    const { jsPDF } = jspdf;
    const outputFileName = `export_${new Date().toISOString().split('T')[0]}.pdf`;

    if (view === 'table') {
       const doc = new jsPDF();
       
       doc.addFileToVFS("Amiri-Regular.ttf", AmiriFont);
       doc.addFont("Amiri-Regular.ttf", "Amiri", "normal");
       doc.setFont("Amiri");

       const head = [headers];
       const body = processedData.map(row => headers.map(header => row[header] !== null && row[header] !== undefined ? String(row[header]) : ''));
       
       const title = 'تقرير بيانات الإكسل';
       const pageWidth = doc.internal.pageSize.getWidth();

       (doc as any).autoTable({
           head: head,
           body: body,
           styles: { font: "Amiri", halign: 'right' },
           headStyles: { halign: 'right', fillColor: [22, 160, 133] },
           didDrawPage: (data: any) => {
               if (data.pageNumber === 1) {
                    doc.setFont("Amiri", "normal");
                    doc.text(title, pageWidth - data.settings.margin.right, 15, { align: 'right' });
               }
           }
       });

       doc.save(outputFileName);

    } else if (view === 'summary' || view === 'rankedSummary') {
        const summaryElement = summaryContentRef.current;
        if (summaryElement) {
            try {
                const canvas = await html2canvas(summaryElement, {
                    scale: 2,
                    backgroundColor: '#ffffff',
                    width: summaryElement.scrollWidth,
                    height: summaryElement.scrollHeight,
                });
                const imgData = canvas.toDataURL('image/png');
                
                const imgWidth = canvas.width;
                const imgHeight = canvas.height;
                
                const pdf = new jsPDF({
                    orientation: imgWidth > imgHeight ? 'l' : 'p',
                    unit: 'px',
                    format: [imgWidth, imgHeight]
                });

                pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);
                pdf.save(outputFileName);
            } catch (err) {
                console.error("Error generating PDF:", err);
                const errorMessage = err instanceof Error ? err.message : "An unknown error occurred during PDF export.";
                setError(`فشل تصدير PDF: ${errorMessage}`);
            }
        }
    }
    setIsLoading(false);
  };

  // --- Auth Handlers ---
  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoginError(null);

    // Admin login
    if (usernameInput === ADMIN_USERNAME && passwordInput === ADMIN_PASSWORD) {
        setCurrentUser('admin');
        localStorage.setItem(CURRENT_USER_STORAGE_KEY, JSON.stringify('admin'));
        setUsernameInput('');
        setPasswordInput('');
        setView(data.length > 0 ? 'summary' : 'upload');
        return;
    }

    // Supervisor login
    if (usernameInput === SUPERVISOR_USERNAME && passwordInput === SUPERVISOR_PASSWORD) {
        setCurrentUser('supervisor');
        localStorage.setItem(CURRENT_USER_STORAGE_KEY, JSON.stringify('supervisor'));
        setUsernameInput('');
        setPasswordInput('');
        setView(data.length > 0 ? 'summary' : 'upload');
        return;
    }

    // Employee login
    const foundEmployee = employees.find(
        emp => emp.username === usernameInput && emp.password === passwordInput
    );

    if (foundEmployee) {
        setIsLoading(true);
        try {
            const firestoreData = await getEmployeeData(foundEmployee.id);
            const userToSet = firestoreData || foundEmployee;

            setCurrentUser(userToSet);
            localStorage.setItem(CURRENT_USER_STORAGE_KEY, JSON.stringify(userToSet));
            
            setUsernameInput('');
            setPasswordInput('');
            setView('summary');
        } catch (err) {
            console.error("Error during employee login fetch:", err);
            setLoginError("حدث خطأ أثناء جلب بياناتك. يرجى المحاولة مرة أخرى.");
        } finally {
            setIsLoading(false);
        }
    } else {
        setLoginError('اسم المستخدم أو كلمة المرور غير صحيحة.');
    }
  };

  const handleLogout = () => {
    setCurrentUser(null);
    localStorage.removeItem(CURRENT_USER_STORAGE_KEY);
    setView('upload');
  };
  
    // --- Employee Management Handlers ---
  const handleAddEmployee = async (e: React.FormEvent) => {
    e.preventDefault();
    const name = newEmployeeName.trim();
    const workplace = newEmployeeWorkplace.trim();
    const username = newEmployeeUsername.trim();
    const password = newEmployeePassword.trim();
    const balance = parseInt(newEmployeeBalance, 10);
    const workdayHours = parseInt(newEmployeeWorkdayHours, 10);

    if (!name || !username || !password || isNaN(balance) || isNaN(workdayHours)) {
      alert("يرجى إدخال جميع الحقول بشكل صحيح.");
      return;
    }

    if (employees.some(emp => emp.username.toLowerCase() === username.toLowerCase())) {
        alert("اسم المستخدم هذا موجود بالفعل. يرجى اختيار اسم آخر.");
        return;
    }
    
    let photoData: string | undefined = undefined;
    if (newEmployeePhotoFile) {
        photoData = await fileToBase64(newEmployeePhotoFile);
    }

    const newEmployee: EmployeeRecord = {
      id: Date.now().toString(),
      name,
      balance,
      username,
      password,
      photo: photoData,
      workdayHours,
      workplace,
    };

    const updatedEmployees = [...employees, newEmployee].sort((a,b) => a.name.localeCompare(b.name, 'ar'));
    setEmployees(updatedEmployees);
    saveEmployeeData(newEmployee.id, newEmployee);

    setNewEmployeeName('');
    setNewEmployeeWorkplace('');
    setNewEmployeeBalance('');
    setNewEmployeeWorkdayHours('7');
    setNewEmployeeUsername(generateRandomUsername());
    setNewEmployeePassword(generateRandomPassword());
    setNewEmployeePhoto(null);
    setNewEmployeePhotoFile(null);
  };

  const handleNewEmployeePhotoChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
        setNewEmployeePhotoFile(file);
        const reader = new FileReader();
        reader.onloadend = () => {
            setNewEmployeePhoto(reader.result as string);
        };
        reader.readAsDataURL(file);
    }
  };

  const handleDeleteEmployee = (id: string) => {
    if (window.confirm("هل أنت متأكد من حذف هذا الموظف؟")) {
      setEmployees(prev => prev.filter(emp => emp.id !== id));
    }
  };

  const handleStartEdit = (employee: EmployeeRecord) => {
    setEditingEmployee({ ...employee });
  };

  const handleCancelEdit = () => {
    setEditingEmployee(null);
  };
  
  const handleSaveEdit = async () => {
    if (!editingEmployee) return;

    const name = editingEmployee.name.trim();
    const username = editingEmployee.username.trim();
    const password = editingEmployee.password.trim();
    const balance = Number(editingEmployee.balance);
    const workdayHours = Number(editingEmployee.workdayHours);

    if (!name || !username || !password || isNaN(balance) || isNaN(workdayHours)) {
      alert("يرجى إدخال جميع الحقول بشكل صحيح.");
      return;
    }

    if (employees.some(emp => emp.id !== editingEmployee.id && emp.username.toLowerCase() === username.toLowerCase())) {
        alert("اسم المستخدم هذا موجود بالفعل. يرجى اختيار اسم آخر.");
        return;
    }
    
    const updatedEmployees = employees.map(emp => (emp.id === editingEmployee.id ? editingEmployee : emp))
      .sort((a,b) => a.name.localeCompare(b.name, 'ar'));
    setEmployees(updatedEmployees);
    saveEmployeeData(editingEmployee.id, editingEmployee);
    setEditingEmployee(null);
  };

  const handleEditPhotoChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0] && editingEmployee) {
        const base64 = await fileToBase64(e.target.files[0]);
        setEditingEmployee({ ...editingEmployee, photo: base64 });
    }
  };
  
  const filteredEmployees = useMemo(() => {
    if (!employeeSearchTerm) return employees;
    return employees.filter(emp => 
        emp.name.toLowerCase().includes(employeeSearchTerm.toLowerCase()) ||
        emp.username.toLowerCase().includes(employeeSearchTerm.toLowerCase()) ||
        emp.workplace?.toLowerCase().includes(employeeSearchTerm.toLowerCase())
    );
  }, [employees, employeeSearchTerm]);

  const handleExportUserList = useCallback(() => {
    if (employees.length === 0) return;
    const headers = ['الاسم', 'اسم المستخدم', 'كلمة المرور'];
    const csvRows = [
        headers.join(','),
        ...employees.map(emp => [
            `"${emp.name.replace(/"/g, '""')}"`,
            `"${emp.username}"`,
            `"${emp.password}"`,
        ].join(','))
    ];
    const csvContent = csvRows.join('\n');
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    const fileName = `user_credentials_${new Date().toISOString().split('T')[0]}.csv`;
    link.setAttribute('download', fileName);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }, [employees]);
  
  const handleCorrectBalances = () => {
    if (window.confirm("هل أنت متأكد من أنك تريد خصم 5 أيام من رصيد كل موظف؟ لا يمكن التراجع عن هذا الإجراء.")) {
        const updatedEmployees = employees.map(emp => ({ ...emp, balance: emp.balance - 5 }));
        setEmployees(updatedEmployees);
        updatedEmployees.forEach(emp => saveEmployeeData(emp.id, { balance: emp.balance }));
        alert("تم خصم 5 أيام من رصيد جميع الموظفين.");
    }
  };

  // --- Data for Summary View ---

  const availableYears = useMemo(() => {
    if (data.length === 0) return [];
    const years = new Set<string>();
    data.forEach(row => {
        const date = parseDate(String(row['التاريخ']));
        if (!isNaN(date.getTime())) {
            years.add(String(date.getUTCFullYear()));
        }
    });
    return Array.from(years).sort((a, b) => parseInt(b) - parseInt(a));
  }, [data, parseDate]);

  useEffect(() => {
    if (selectedYear === 'all') {
        setSelectedMonth('all');
    }
  }, [selectedYear]);

  const displayedSummary = useMemo(() => {
      const dateFilteredData = data.filter(row => {
          if (selectedYear === 'all') return true;
          const date = parseDate(String(row['التاريخ']));
          if (isNaN(date.getTime())) return false;

          const year = date.getUTCFullYear();
          const month = date.getUTCMonth() + 1;

          const yearMatches = String(year) === selectedYear;
          if (!yearMatches) return false;

          const monthMatches = selectedMonth === 'all' || String(month) === selectedMonth;
          return monthMatches;
      });
      
      let finalSummary = generateSummary(dateFilteredData, employees);
      
      // If the current user is an employee (object), filter for their data only. Admins and supervisors see all.
      if (currentUser && typeof currentUser === 'object') {
          return finalSummary.filter(emp => normalizeName(emp.name) === normalizeName(currentUser.name));
      }
      
      // For admin/supervisor, apply search term
      if (summarySearchTerm) {
          finalSummary = finalSummary.filter(emp => emp.name.toLowerCase().includes(summarySearchTerm.toLowerCase()));
      }

      return finalSummary;
  }, [data, summarySearchTerm, selectedYear, selectedMonth, generateSummary, employees, parseDate, currentUser]);

  const displayedUniqueLeaveTypes = useMemo(() => {
    const allTypes = new Set<string>();
    displayedSummary.forEach(employee => {
        employee.leaves.forEach(leave => {
            allTypes.add(leave.type);
        });
    });
    return Array.from(allTypes).sort((a, b) => {
        const order = ['اجازة اعتيادية', 'اجازة مرضية', 'ملخص الزمنيات'];
        const indexA = order.indexOf(a);
        const indexB = order.indexOf(b);
        if(indexA > -1 && indexB > -1) return indexA - indexB;
        if(indexA > -1) return -1;
        if(indexB > -1) return 1;
        return a.localeCompare(b, 'ar');
    });
  }, [displayedSummary]);
  
  const getSummaryTotalLeaveDays = useCallback((employee: EmployeeSummary): number => {
    return employee.leaves.reduce((total, leave) => total + leave.dayCount, 0);
  }, []);

  const rankedFilteredSummary = useMemo(() => {
    // Use the `displayedSummary` which is already filtered by date
    const sorted = [...displayedSummary].sort((a, b) => {
        const totalA = getRegularAndTimeBasedTotalDays(a);
        const totalB = getRegularAndTimeBasedTotalDays(b);
        return totalA - totalB;
    });
    return sorted.filter(emp => getRegularAndTimeBasedTotalDays(emp) >= 1);
  }, [displayedSummary, getRegularAndTimeBasedTotalDays]);
  
  const rankedSummaryLeaveTypes = useMemo(() => 
    displayedUniqueLeaveTypes.filter(type => !type.includes('طويلة') && !type.includes('مرضية')), 
    [displayedUniqueLeaveTypes]
  );

  const rankedSummaryTableHeaders = useMemo(() => 
    ['ت', 'الاسم', 'عدد ايام الاجازات', ...rankedSummaryLeaveTypes.map(type => type === 'ملخص الزمنيات' ? 'عن ساعات زمنية' : 'أجلاة اعتيادية')], 
    [rankedSummaryLeaveTypes]
  );
  
  const alphabeticallySortedSummary = useMemo(() => {
    return [...rankedFilteredSummary].sort((a, b) => a.name.localeCompare(b.name, 'ar'));
  }, [rankedFilteredSummary]);

  const groupedRankedSummary = useMemo(() => {
    const groups: Record<number, EmployeeSummary[]> = {};
    
    rankedFilteredSummary.forEach(employee => {
        const totalDays = getRegularAndTimeBasedTotalDays(employee);
        if (totalDays > 0) {
            if (!groups[totalDays]) {
                groups[totalDays] = [];
            }
            groups[totalDays].push(employee);
        }
    });

    const getTimeSummaryDays = (employee: EmployeeSummary): number => {
        const timeSummary = employee.leaves.find(l => l.type === 'ملخص الزمنيات');
        return timeSummary ? timeSummary.dayCount : 0;
    };

    for (const dayCount in groups) {
        groups[dayCount].sort((a, b) => {
            const timeDaysA = getTimeSummaryDays(a);
            const timeDaysB = getTimeSummaryDays(b);

            if (timeDaysA === 0 && timeDaysB > 0) return -1;
            if (timeDaysB === 0 && timeDaysA > 0) return 1;
            
            return a.name.localeCompare(b.name, 'ar');
        });
    }

    return groups;
  }, [rankedFilteredSummary, getRegularAndTimeBasedTotalDays]);

  const employeesWithSickLeave = useMemo(() => {
    return summary
        .map(employee => {
            const sickLeaves = employee.leaves.filter(l => l.type.includes('مرضية'));
            if (sickLeaves.length > 0) {
                const totalSickDays = sickLeaves.reduce((sum, l) => sum + l.dayCount, 0);
                if (totalSickDays <= 5) {
                    return {
                        name: employee.name,
                        sickLeaves: sickLeaves.map(l => ({
                            dateDetails: l.dateDetails,
                            dayCount: l.dayCount
                        }))
                    };
                }
            }
            return null;
        })
        .filter((e): e is { name: string; sickLeaves: { dateDetails: string; dayCount: number }[] } => e !== null)
        .sort((a, b) => a.name.localeCompare(b.name, 'ar'));
  }, [summary]);

  const workplaces = useMemo(() => ['all', ...Array.from(new Set(employees.map(e => e.workplace).filter(Boolean))) as string[]], [employees]);

  const handleCopyToClipboard = useCallback(() => {
    if (displayedSummary.length === 0) return;

    const headers = ['الاسم', ...displayedUniqueLeaveTypes, 'مجموع الاجازات الاعتيادية', 'المجموع الكلي'];
    let textToCopy = headers.join('\t') + '\n';
    
    displayedSummary.forEach(employee => {
        const leaveMap = new Map<string, LeaveSummary[]>();
        employee.leaves.forEach(l => {
            if (!leaveMap.has(l.type)) leaveMap.set(l.type, []);
            leaveMap.get(l.type)!.push(l);
        });
        
        const rowData = [employee.name];

        displayedUniqueLeaveTypes.forEach(type => {
            const leaves = leaveMap.get(type);
            let content = '';
            if (leaves && leaves.length > 0) {
                const totalDays = leaves.reduce((sum, l) => sum + l.dayCount, 0);
                const totalHours = leaves.reduce((sum, l) => sum + l.hourCount, 0);
                
                const parts: string[] = [];
                if (totalDays > 0) parts.push(`${toArabicNumerals(totalDays)} يوم`);
                if (totalHours > 0) parts.push(`${toArabicNumerals(totalHours)} ساعة`);
                if (parts.length > 0) content = parts.join(' و ');
            }
            rowData.push(content);
        });
        
        const regularAndTimeBasedTotal = getRegularAndTimeBasedTotalDays(employee);
        rowData.push(regularAndTimeBasedTotal > 0 ? formatDaysArabic(regularAndTimeBasedTotal) : '');

        let totalDays = 0;
        let totalHours = 0;
        employee.leaves.forEach(l => {
            totalDays += l.dayCount;
            totalHours += l.hourCount;
        });
        
        const employeeRecord = employees.find(e => e.name === employee.name);
        const workdayHours = employeeRecord?.workdayHours || 7;

        if (totalHours >= workdayHours) {
            totalDays += Math.floor(totalHours / workdayHours);
            totalHours %= workdayHours;
        }
        const totalParts: string[] = [];
        if (totalDays > 0) totalParts.push(`${toArabicNumerals(totalDays)} يوم`);
        const totalHoursFormatted = parseFloat(totalHours.toFixed(2));
        if (totalHoursFormatted > 0) totalParts.push(`${toArabicNumerals(totalHoursFormatted)} ساعة`);
        const totalContent = totalParts.join(' و ') || '٠';
        rowData.push(totalContent);

        textToCopy += rowData.join('\t') + '\n';
    });


    navigator.clipboard.writeText(textToCopy.trim()).then(() => {
        setShowCopyNotification(true);
        setTimeout(() => setShowCopyNotification(false), 2000); 
    });
  }, [displayedSummary, displayedUniqueLeaveTypes, getRegularAndTimeBasedTotalDays, employees]);

    const handleExportSummaryToWord = useCallback(() => {
        if (displayedSummary.length === 0) return;

        const headers = ['الاسم', ...displayedUniqueLeaveTypes, 'مجموع الاجازات الاعتيادية', 'المجموع الكلي'];
        let tableHTML = `<table border="1" style="border-collapse: collapse; width: 100%; text-align: center; font-family: 'Amiri', Arial, sans-serif;">
            <thead style="background-color: #f2f2f2;">
                <tr>${headers.map(h => `<th style="padding: 8px;">${h}</th>`).join('')}</tr>
            </thead>
            <tbody>`;

        displayedSummary.forEach(employee => {
            const leaveMap = new Map<string, LeaveSummary[]>();
            employee.leaves.forEach(l => {
                if (!leaveMap.has(l.type)) leaveMap.set(l.type, []);
                leaveMap.get(l.type)!.push(l);
            });
            
            tableHTML += '<tr>';
            tableHTML += `<td style="padding: 8px; text-align: right; padding-right: 10px;">${employee.name}</td>`;

            displayedUniqueLeaveTypes.forEach(type => {
                const leaves = leaveMap.get(type);
                let cellContent = '';
                if (leaves && leaves.length > 0) {
                    const totalDays = leaves.reduce((sum, l) => sum + l.dayCount, 0);
                    const totalHours = leaves.reduce((sum, l) => sum + l.hourCount, 0);
                    const parts: string[] = [];
                    if (totalDays > 0) parts.push(`${toArabicNumerals(totalDays)} يوم`);
                    if (totalHours > 0) parts.push(`${toArabicNumerals(totalHours)} ساعة`);
                    if (parts.length > 0) cellContent = parts.join(' و ');
                }
                tableHTML += `<td style="padding: 8px;">${cellContent}</td>`;
            });

            const regularAndTimeBasedTotal = getRegularAndTimeBasedTotalDays(employee);
            tableHTML += `<td style="padding: 8px; font-weight: bold; color: #4338ca;">${regularAndTimeBasedTotal > 0 ? formatDaysArabic(regularAndTimeBasedTotal) : ''}</td>`;
            
            let totalDays = 0;
            let totalHours = 0;
            employee.leaves.forEach(l => {
                totalDays += l.dayCount;
                totalHours += l.hourCount;
            });
            
            const employeeRecord = employees.find(e => e.name === employee.name);
            const workdayHours = employeeRecord?.workdayHours || 7;

            if (totalHours >= workdayHours) {
                totalDays += Math.floor(totalHours / workdayHours);
                totalHours %= workdayHours;
            }
            const totalParts: string[] = [];
            if (totalDays > 0) totalParts.push(`${toArabicNumerals(totalDays)} يوم`);
            const totalHoursFormatted = parseFloat(totalHours.toFixed(2));
            if (totalHoursFormatted > 0) totalParts.push(`${toArabicNumerals(totalHoursFormatted)} ساعة`);
            const totalContent = totalParts.join(' و ') || '٠';

            tableHTML += `<td style="padding: 8px; font-weight: bold;">${totalContent}</td>`;
            tableHTML += '</tr>';
        });


        tableHTML += `</tbody></table>`;

        const htmlContent = `
            <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
            <head><meta charset='utf-8'><title>Export HTML To Doc</title></head>
            <body dir="rtl">
                <h1 style="text-align: center; font-family: 'Amiri', Arial, sans-serif;">ملخص إجازات الموظفين</h1>
                ${tableHTML}
            </body></html>
        `;

        const blob = new Blob(['\uFEFF' + htmlContent], { type: 'application/msword' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `ملخص_إجازات_الموظفين.doc`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }, [displayedSummary, displayedUniqueLeaveTypes, getRegularAndTimeBasedTotalDays, employees]);
  
  const handleExportWord = useCallback(() => {
    let allTablesHTML = '';
    const getLeaveMap = (employee: EmployeeSummary) => {
        const map = new Map<string, LeaveSummary[]>();
        employee.leaves.forEach(l => {
            if (!map.has(l.type)) map.set(l.type, []);
            map.get(l.type)!.push(l);
        });
        return map;
    };


    if (rankedSortOrder === 'alphabetical') {
        allTablesHTML += `<h2 style="text-align: center; font-family: 'Times New Roman'; font-size: 16pt; margin-top: 20px;">اسماء الموظفين مرتبة أبجديًا</h2>`;
        
        let tableHTML = `<table border="1" style="border-collapse: collapse; width: 100%; text-align: center; font-family: 'Times New Roman'; font-size: 14pt;">
            <thead style="background-color: #f2f2f2;">
                <tr>${rankedSummaryTableHeaders.map(h => `<th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">${h}</th>`).join('')}</tr>
            </thead>
            <tbody>`;

        alphabeticallySortedSummary.forEach((employee, index) => {
            const totalDays = getRegularAndTimeBasedTotalDays(employee);
            const leaveMap = getLeaveMap(employee);

            tableHTML += '<tr>';
            tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${toArabicNumerals(index + 1)}</td>`;
            tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman'; text-align: right;">${employee.name}</td>`;
            tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${formatDaysArabic(totalDays)}</td>`;
            
            rankedSummaryLeaveTypes.forEach(type => {
                const leaves = leaveMap.get(type) || [];
                let content = ' ';
                if (leaves.length > 0) {
                    if (type === 'ملخص الزمنيات') {
                        const dayCount = leaves.reduce((sum, l) => sum + l.dayCount, 0);
                        if (dayCount > 0) content = `(${toArabicNumerals(dayCount)}) يوم`;
                    } else if (type === 'اجازة اعتيادية') {
                        const dateDetails = leaves.map(l => l.dateDetails).join(' | ');
                        if (dateDetails) {
                           const datePart = dateDetails.split('|').map(part => {
                                const segments = part.trim().split('/');
                                if (segments.length === 3) {
                                    const days = segments[0].split('،').map(d => parseInt(d.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d).toString()))).sort((a,b) => a-b).map(d=> toArabicNumerals(d)).join('،');
                                    return `${days}/${segments[1]}/${segments[2]}`;
                                }
                                return part;
                            }).join(' | ');
                            content = `<span dir="ltr" style="unicode-bidi: embed;">${datePart}</span>`;
                        } else {
                            content = ' ';
                        }
                    }
                }
                tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${content}</td>`;
            });
            
            tableHTML += '</tr>';
        });
        tableHTML += '</tbody></table>';
        allTablesHTML += tableHTML;

    } else {
        const sortedDayCounts = Object.keys(groupedRankedSummary).sort((a, b) => parseInt(a) - parseInt(b));
        sortedDayCounts.forEach(dayCount => {
            const dayCountNum = parseInt(dayCount, 10);
            allTablesHTML += `<h2 style="text-align: center; font-family: 'Times New Roman'; font-size: 16pt; margin-top: 20px;">اسماء الموظفين الذين تم منحهم ${formatDaysArabic(dayCountNum)}</h2>`;
            
            let tableHTML = `<table border="1" style="border-collapse: collapse; width: 100%; text-align: center; font-family: 'Times New Roman'; font-size: 14pt;">
                <thead style="background-color: #f2f2f2;">
                    <tr>${rankedSummaryTableHeaders.map(h => `<th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">${h}</th>`).join('')}</tr>
                </thead>
                <tbody>`;

            const employeesInGroup = groupedRankedSummary[dayCountNum];
            employeesInGroup.forEach((employee, index) => {
                const leaveMap = getLeaveMap(employee);
                
                tableHTML += '<tr>';
                tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${toArabicNumerals(index + 1)}</td>`;
                tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman'; text-align: right;">${employee.name}</td>`;
                tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${formatDaysArabic(dayCountNum)}</td>`;
                
                rankedSummaryLeaveTypes.forEach(type => {
                    const leaves = leaveMap.get(type) || [];
                    let content = ' ';
                    if (leaves.length > 0) {
                       if (type === 'ملخص الزمنيات') {
                            const dayCount = leaves.reduce((sum, l) => sum + l.dayCount, 0);
                            if (dayCount > 0) content = `(${toArabicNumerals(dayCount)}) يوم`;
                        } else if (type === 'اجازة اعتيادية') {
                           const dateDetails = leaves.map(l => l.dateDetails).join(' | ');
                           if (dateDetails) {
                                const datePart = dateDetails.split('|').map(part => {
                                    const segments = part.trim().split('/');
                                    if (segments.length === 3) {
                                        const days = segments[0].split('،').map(d => parseInt(d.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d).toString()))).sort((a,b) => a-b).map(d=> toArabicNumerals(d)).join('،');
                                        return `${days}/${segments[1]}/${segments[2]}`;
                                    }
                                    return part;
                                }).join(' | ');
                                content = `<span dir="ltr" style="unicode-bidi: embed;">${datePart}</span>`;
                            } else {
                                content = ' ';
                            }
                        }
                    }
                    tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${content}</td>`;
                });

                tableHTML += '</tr>';
            });
            tableHTML += '</tbody></table>';
            allTablesHTML += tableHTML;
        });
    }

    if (employeesWithSickLeave.length > 0) {
        allTablesHTML += `<h2 style="text-align: center; font-family: 'Times New Roman'; font-size: 16pt; margin-top: 20px;">الموظفون الحاصلون على اجازة مرضية</h2>`;
        
        let sickLeaveTableHTML = `<table border="1" style="border-collapse: collapse; width: 100%; text-align: center; font-family: 'Times New Roman'; font-size: 14pt;">
            <thead style="background-color: #f2f2f2;">
                <tr>
                    <th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">ت</th>
                    <th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman'; text-align: right;">الاسم</th>
                    <th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">المدة</th>
                    <th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">التاريخ</th>
                </tr>
            </thead>
            <tbody>`;
        
        employeesWithSickLeave.forEach((employee, index) => {
            const totalSickDays = employee.sickLeaves.reduce((sum, l) => sum + l.dayCount, 0);
            const sickLeaveDates = employee.sickLeaves.map(l => l.dateDetails).join(' | ');
    
            sickLeaveTableHTML += '<tr>';
            sickLeaveTableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${toArabicNumerals(index + 1)}</td>`;
            sickLeaveTableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman'; text-align: right;">${employee.name}</td>`;
            sickLeaveTableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${formatDaysArabic(totalSickDays)}</td>`;
            sickLeaveTableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';"><span dir="ltr" style="unicode-bidi: embed;">${sickLeaveDates}</span></td>`;
            sickLeaveTableHTML += '</tr>';
        });
        
        sickLeaveTableHTML += '</tbody></table>';
        allTablesHTML += sickLeaveTableHTML;
    }


    const htmlContent = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head><meta charset='utf-8'><title>Export HTML To Doc</title></head>
        <body dir="rtl">
            ${allTablesHTML}
        </body></html>
    `;

    const blob = new Blob(['\uFEFF' + htmlContent], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `ملخص_مرتب.doc`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}, [groupedRankedSummary, rankedSummaryTableHeaders, rankedSortOrder, alphabeticallySortedSummary, getRegularAndTimeBasedTotalDays, rankedSummaryLeaveTypes, employeesWithSickLeave]);

    // --- Monthly Report Handlers ---
    const getLeavePeriodDetails = useCallback((monthlyLeaves: DataRow[], leaveTypeKeyword: string): { count: number; dateRange: string } => {
        const leaves = monthlyLeaves.filter(l => String(l['نوع الاجازة']).includes(leaveTypeKeyword));
        if (leaves.length === 0) return { count: 0, dateRange: '' };

        const dates = leaves.map(l => parseDate(String(l['التاريخ']))).sort((a, b) => a.getTime() - b.getTime());
        
        const count = leaves.length;
        const minDate = dates[0];
        const maxDate = dates[dates.length - 1];

        const startDay = minDate.getUTCDate();
        const startMonth = minDate.getUTCMonth() + 1;
        const startYear = minDate.getUTCFullYear();
        
        let dateRangeStr = '';
        if (minDate.getTime() === maxDate.getTime()) {
            dateRangeStr = `${toArabicNumerals(startDay)}/${toArabicNumerals(startMonth)}/${toArabicNumerals(startYear)}`;
        } else {
            const endDay = maxDate.getUTCDate();
            const month = minDate.getUTCMonth() + 1; 
            const year = minDate.getUTCFullYear();
            dateRangeStr = `${toArabicNumerals(startDay)}-${toArabicNumerals(endDay)}/${toArabicNumerals(month)}/${toArabicNumerals(year)}`;
        }

        return { count, dateRange: dateRangeStr };
    }, [parseDate]);

    const handleGenerateMonthlyReport = useCallback(() => {
        setIsLoading(true);
        const reportStartDate = new Date(Date.UTC(reportYear, reportMonth - 1, 1));
        const reportEndDate = new Date(Date.UTC(reportYear, reportMonth, 0));
        
        const normalizedEmployeeMap = new Map<string, EmployeeRecord>();
        employees.forEach(emp => {
            normalizedEmployeeMap.set(normalizeName(emp.name), emp);
        });
    
        const nameResolutionCache = new Map<string, string | null>();
    
        const resolveCanonicalName = (sheetName: string): string | null => {
            if (!sheetName) return null;
            const cached = nameResolutionCache.get(sheetName);
            if (cached !== undefined) {
                return cached;
            }
    
            const normalizedSheetName = normalizeName(sheetName);
    
            if (normalizedEmployeeMap.has(normalizedSheetName)) {
                const canonicalName = normalizedEmployeeMap.get(normalizedSheetName)!.name;
                nameResolutionCache.set(sheetName, canonicalName);
                return canonicalName;
            }
    
            const potentialMatches = employees.filter(emp => {
                const normalizedEmpName = normalizeName(emp.name);
                return normalizedEmpName.includes(normalizedSheetName) || normalizedSheetName.includes(normalizedEmpName);
            });
    
            if (potentialMatches.length === 1) {
                const canonicalName = potentialMatches[0].name;
                nameResolutionCache.set(sheetName, canonicalName);
                return canonicalName;
            } else if (potentialMatches.length > 1) {
                potentialMatches.sort((a, b) =>
                    Math.abs(normalizeName(a.name).length - normalizedSheetName.length) -
                    Math.abs(normalizeName(b.name).length - normalizedSheetName.length)
                );
                const canonicalName = potentialMatches[0].name;
                nameResolutionCache.set(sheetName, canonicalName);
                return canonicalName;
            }
    
            nameResolutionCache.set(sheetName, null);
            return null;
        };
        
        const newReportData: MonthlyReportRow[] = employees.map(employee => {
            const leavesForEmployee = data.filter(d => d['الاسم'] && resolveCanonicalName(d['الاسم']) === employee.name);
            
            let workdayHours = employee.workdayHours || 7; // Start with stored or default
            const regularLeaves = leavesForEmployee.filter(l => String(l['نوع الاجازة']).includes('اعتيادية'));
            if (regularLeaves.length > 0) {
                const typicalHours = parseFloat(String(regularLeaves[0]['القيمة']));
                if (typicalHours === 6 || typicalHours === 7) {
                    workdayHours = typicalHours;
                }
            }


            // Leaves within the current month
            const monthlyLeaves = leavesForEmployee.filter(row => {
                const date = parseDate(String(row['التاريخ']));
                return date >= reportStartDate && date <= reportEndDate;
            });

            const regularLeavesMonth = monthlyLeaves.filter(l => String(l['نوع الاجازة']).includes('اعتيادية'));
            const regularDates = regularLeavesMonth.map(l => parseDate(String(l['التاريخ'])).getUTCDate()).sort((a,b)=>a-b).map(toArabicNumerals).join('، ');

            const sickLeaveDetails = getLeavePeriodDetails(monthlyLeaves, 'مرضية');
            const longLeaveDetails = getLeavePeriodDetails(monthlyLeaves, 'طويلة');
            const travelLeaveDetails = getLeavePeriodDetails(monthlyLeaves, 'سفر');
            const absentLeaveDetails = getLeavePeriodDetails(monthlyLeaves, 'غائب');
            
            // --- Cumulative Hourly Calculation for monthly display ---
            const leavesBeforeMonth = leavesForEmployee.filter(row => parseDate(String(row['التاريخ'])) < reportStartDate);
            
            const hoursBeforeMonthSheet = leavesBeforeMonth
                .filter(l => String(l['نوع الاجازة']).startsWith('زمنية'))
                .reduce((sum, l) => sum + (parseFloat(String(l['القيمة'])) || 0), 0);
            const totalHoursBeforeMonth = hoursBeforeMonthSheet + (employee.priorHourlyBalance || 0);

            const daysBeforeMonth = Math.floor(totalHoursBeforeMonth / workdayHours);

            const hourlyLeavesMonth = monthlyLeaves.filter(l => String(l['نوع الاجازة']).startsWith('زمنية'));
            const hoursThisMonth = hourlyLeavesMonth.reduce((sum, l) => sum + (parseFloat(String(l['القيمة'])) || 0), 0);
            
            const totalHoursAtMonthEnd = totalHoursBeforeMonth + hoursThisMonth;
            const daysAtMonthEnd = Math.floor(totalHoursAtMonthEnd / workdayHours);
            const remainingHoursAtMonthEnd = totalHoursAtMonthEnd % workdayHours;

            const daysToShowForMonth = daysAtMonthEnd - daysBeforeMonth;
            const hoursToShowForMonth = remainingHoursAtMonthEnd;

            // --- Calculation for final remaining balance ---
            const leavesUntilMonthEnd = leavesForEmployee.filter(row => {
                const date = parseDate(String(row['التاريخ']));
                return date <= reportEndDate;
            });

            const totalRegularDays = leavesUntilMonthEnd.filter(l => String(l['نوع الاجازة']).includes('اعتيادية')).length;
            const totalLongLeaveDays = leavesUntilMonthEnd.filter(l => String(l['نوع الاجازة']).includes('طويلة')).length;
            
            const hoursUntilMonthEndSheet = leavesUntilMonthEnd
                .filter(l => String(l['نوع الاجازة']).startsWith('زمنية'))
                .reduce((sum, l) => sum + (parseFloat(String(l['القيمة'])) || 0), 0);
            const totalHoursUntilMonthEnd = hoursUntilMonthEndSheet + (employee.priorHourlyBalance || 0);
            const totalDaysFromHours = Math.floor(totalHoursUntilMonthEnd / workdayHours);
            
            const finalBalance = (employee.balance || 0) - totalRegularDays - totalLongLeaveDays - totalDaysFromHours;

            return {
                name: employee.name,
                workplace: employee.workplace || 'غير محدد',
                initialBalance: employee.balance || 0,
                regularLeaves: { count: regularLeavesMonth.length, dates: regularDates },
                hourlyLeaves: { days: daysToShowForMonth, hours: hoursToShowForMonth },
                sickLeave: sickLeaveDetails,
                longLeave: longLeaveDetails,
                travelLeave: travelLeaveDetails,
                absentLeave: absentLeaveDetails,
                finalBalance: finalBalance,
            };
        });

        setMonthlyReportData(newReportData);
        setIsLoading(false);
    }, [employees, data, reportYear, reportMonth, parseDate, getLeavePeriodDetails]);

    const handleGenerateAlphabeticalReport = useCallback(() => {
        setIsLoading(true);
        
        const normalizedEmployeeMap = new Map<string, EmployeeRecord>();
        employees.forEach(emp => {
            normalizedEmployeeMap.set(normalizeName(emp.name), emp);
        });
        const nameResolutionCache = new Map<string, string | null>();
        const resolveCanonicalName = (sheetName: string): string | null => {
            if (!sheetName) return null;
            const cached = nameResolutionCache.get(sheetName);
            if (cached !== undefined) return cached;
    
            const normalizedSheetName = normalizeName(sheetName);
            if (normalizedEmployeeMap.has(normalizedSheetName)) {
                const canonicalName = normalizedEmployeeMap.get(normalizedSheetName)!.name;
                nameResolutionCache.set(sheetName, canonicalName);
                return canonicalName;
            }
    
            const potentialMatches = employees.filter(emp => {
                const normalizedEmpName = normalizeName(emp.name);
                return normalizedEmpName.includes(normalizedSheetName) || normalizedSheetName.includes(normalizedEmpName);
            });
    
            if (potentialMatches.length === 1) {
                const canonicalName = potentialMatches[0].name;
                nameResolutionCache.set(sheetName, canonicalName);
                return canonicalName;
            }
            nameResolutionCache.set(sheetName, null);
            return null;
        };

        const sortedEmployees = [...employees].sort((a, b) => a.name.localeCompare(b.name, 'ar'));

        const reportData = sortedEmployees.map(employee => {
            const leavesForEmployee = data.filter(d => d['الاسم'] && resolveCanonicalName(d['الاسم']) === employee.name);
            
            let workdayHours = employee.workdayHours || 7;
            const regularLeavesForCheck = leavesForEmployee.filter(l => String(l['نوع الاجازة']).includes('اعتيادية'));
             if (regularLeavesForCheck.length > 0) {
                const typicalHours = parseFloat(String(regularLeavesForCheck[0]['القيمة']));
                if (typicalHours === 6 || typicalHours === 7) {
                    workdayHours = typicalHours;
                }
            }
            
            // Regular leaves
            const totalRegularDays = regularLeavesForCheck.length;
            const regularLeavesDates = regularLeavesForCheck.map(l => parseDate(String(l['التاريخ'])));
            const regularDatesFormatted = regularLeavesDates.length > 0
                ? regularLeavesDates.map(d => d.getUTCDate()).sort((a, b) => a - b).map(toArabicNumerals).join('، ')
                : '';

            // Other leaves (for display)
            const sickLeavesRaw = leavesForEmployee.filter(l => String(l['نوع الاجازة']).includes('مرضية'));
            const totalSickDays = sickLeavesRaw.length;
            const sickDatesFormatted = formatDateRanges(sickLeavesRaw.map(l => parseDate(String(l['التاريخ']))));
            
            const longLeavesRaw = leavesForEmployee.filter(l => String(l['نوع الاجازة']).includes('طويلة'));
            const totalLongLeaveDays = longLeavesRaw.length;
            const longDatesFormatted = formatDateRanges(longLeavesRaw.map(l => parseDate(String(l['التاريخ']))));

            const travelLeavesRaw = leavesForEmployee.filter(l => String(l['نوع الاجازة']).includes('سفر'));
            const totalTravelDays = travelLeavesRaw.length;
            const travelDatesFormatted = formatDateRanges(travelLeavesRaw.map(l => parseDate(String(l['التاريخ']))));

            const absentLeavesRaw = leavesForEmployee.filter(l => String(l['نوع الاجازة']).includes('غائب'));
            const totalAbsentDays = absentLeavesRaw.length;
            const absentDatesFormatted = formatDateRanges(absentLeavesRaw.map(l => parseDate(String(l['التاريخ']))));

            // Hourly leaves calculation
            const totalHoursFromSheet = leavesForEmployee
                .filter(l => String(l['نوع الاجازة']).startsWith('زمنية'))
                .reduce((sum, l) => sum + (parseFloat(String(l['القيمة'])) || 0), 0);
                
            const totalHoursOverall = totalHoursFromSheet + (employee.priorHourlyBalance || 0);
            const totalDaysFromHours = Math.floor(totalHoursOverall / workdayHours);
            const remainingHours = totalHoursOverall % workdayHours;
            
            // Corrected final balance calculation
            const totalDeductedDays = totalRegularDays + totalLongLeaveDays + totalDaysFromHours;
            const finalBalance = (employee.balance || 0) - totalDeductedDays;

            return {
                name: employee.name,
                workplace: employee.workplace || 'غير محدد',
                initialBalance: employee.balance || 0,
                regularLeaves: { count: totalRegularDays, dates: regularDatesFormatted },
                hourlyLeaves: { days: totalDaysFromHours, hours: remainingHours },
                sickLeave: { count: totalSickDays, dateRange: sickDatesFormatted },
                longLeave: { count: totalLongLeaveDays, dateRange: longDatesFormatted },
                travelLeave: { count: totalTravelDays, dateRange: travelDatesFormatted },
                absentLeave: { count: totalAbsentDays, dateRange: absentDatesFormatted },
                finalBalance: finalBalance,
            };
        });

        setAlphabeticalReportData(reportData);
        setShowAlphabeticalReport(true);
        setIsLoading(false);
    }, [employees, data, parseDate]);
    
    const handleExportMonthlyReportToWord = useCallback(() => {
        const dataToExport = showAlphabeticalReport ? alphabeticalReportData : monthlyReportData;
        const reportTitle = showAlphabeticalReport
            ? 'التقرير الأبجدي الشامل'
            : `التقرير الشهري لـ ${toArabicNumerals(reportMonth)}/${toArabicNumerals(reportYear)}`;
            
        if (dataToExport.length === 0) {
            alert("لا توجد بيانات لتصديرها. يرجى إنشاء تقرير أولاً.");
            return;
        }

        const filteredData = dataToExport.filter(row => reportWorkplaceFilter === 'all' || row.workplace === reportWorkplaceFilter);

        const headers: string[] = ['الاسم', 'مكان العمل', 'الرصيد الأولي'];
        if (reportColumnVisibility.regular) headers.push(showAlphabeticalReport ? 'اعتيادية (أيام)' : 'اعتيادية');
        if (reportColumnVisibility.hourly) headers.push('زمنية (أيام/ساعات)');
        if (reportColumnVisibility.sick) headers.push('مرضية');
        if (reportColumnVisibility.long) headers.push('طويلة');
        if (reportColumnVisibility.travel) headers.push('سفر');
        if (reportColumnVisibility.absent) headers.push('غائب');
        headers.push('الرصيد المتبقي');

        let tableHTML = `<table border="1" style="border-collapse: collapse; width: 100%; text-align: center; font-family: 'Times New Roman', serif;">
            <thead style="background-color: #f2f2f2;">
                <tr>${headers.map(h => `<th style="padding: 8px;">${h}</th>`).join('')}</tr>
            </thead>
            <tbody>`;
            
        filteredData.forEach(row => {
            tableHTML += `<tr>
                <td style="padding: 8px; text-align: right;">${row.name}</td>
                <td style="padding: 8px;">${row.workplace}</td>
                <td style="padding: 8px;">${toArabicNumerals(row.initialBalance)}</td>`;

            if (reportColumnVisibility.regular) {
                const content = showAlphabeticalReport ? row.regularLeaves.dates : (row.regularLeaves.dates);
                tableHTML += `<td style="padding: 8px;">${content || ''}</td>`;
            }
            if (reportColumnVisibility.hourly) {
                tableHTML += `<td style="padding: 8px;">${formatLeaveCount(row.hourlyLeaves.days, row.hourlyLeaves.hours)}</td>`;
            }
            if (reportColumnVisibility.sick) {
                let content = '';
                if (showAlphabeticalReport) {
                    content = row.sickLeave.dateRange || '';
                } else {
                    content = row.sickLeave.count > 0 ? `${toArabicNumerals(row.sickLeave.count)} يوم (${row.sickLeave.dateRange})` : '';
                }
                tableHTML += `<td style="padding: 8px;">${content}</td>`;
            }
             if (reportColumnVisibility.long) {
                let content = '';
                if (showAlphabeticalReport) {
                    content = row.longLeave.dateRange || '';
                } else {
                    content = row.longLeave.count > 0 ? `${toArabicNumerals(row.longLeave.count)} يوم (${row.longLeave.dateRange})` : '';
                }
                tableHTML += `<td style="padding: 8px;">${content}</td>`;
            }
             if (reportColumnVisibility.travel) {
                let content = '';
                if (showAlphabeticalReport) {
                    content = row.travelLeave.dateRange || '';
                } else {
                    content = row.travelLeave.count > 0 ? `${toArabicNumerals(row.travelLeave.count)} يوم (${row.travelLeave.dateRange})` : '';
                }
                tableHTML += `<td style="padding: 8px;">${content}</td>`;
            }
            if (reportColumnVisibility.absent) {
                let content = '';
                if (showAlphabeticalReport) {
                    content = row.absentLeave.dateRange || '';
                } else {
                    content = row.absentLeave.count > 0 ? `${toArabicNumerals(row.absentLeave.count)} يوم (${row.absentLeave.dateRange})` : '';
                }
                tableHTML += `<td style="padding: 8px;">${content}</td>`;
            }

            tableHTML += `<td style="padding: 8px; font-weight: bold;">${toArabicNumerals(row.finalBalance)}</td></tr>`;
        });
        
        tableHTML += `</tbody></table>`;

        const htmlContent = `
            <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
            <head><meta charset='utf-8'><title>Monthly Report</title></head>
            <body dir="rtl">
                <h1 style="text-align: center; font-family: 'Times New Roman', serif;">${reportTitle}</h1>
                ${tableHTML}
            </body></html>`;

        const blob = new Blob(['\uFEFF' + htmlContent], { type: 'application/msword' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `تقرير_${showAlphabeticalReport ? 'ابجدي' : 'شهري'}.doc`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);

    }, [monthlyReportData, alphabeticalReportData, showAlphabeticalReport, reportColumnVisibility, reportWorkplaceFilter, reportMonth, reportYear]);

    // When generating a new monthly report, hide the alphabetical one to avoid confusion.
    useEffect(() => {
        if (monthlyReportData.length > 0) {
            setShowAlphabeticalReport(false);
        }
    }, [monthlyReportData]);

    const groupedMonthlyReport = useMemo(() => {
        if (monthlyReportData.length === 0) {
            return {};
        }
        return monthlyReportData.reduce((acc, row) => {
            const key = row.workplace || 'غير محدد';
            if (!acc[key]) {
                acc[key] = [];
            }
            acc[key].push(row);
            return acc;
        }, {} as Record<string, MonthlyReportRow[]>);
    }, [monthlyReportData]);

    // --- RENDER LOGIC ---

    if (isLoading && !currentUser) {
        return (
            <div className="flex items-center justify-center h-screen bg-slate-100 dark:bg-gray-900">
                <div className="text-center">
                    <div className="w-16 h-16 border-4 border-dashed rounded-full animate-spin border-indigo-600"></div>
                    <p className="mt-4 text-lg font-semibold dark:text-white">...جاري تهيئة التطبيق</p>
                </div>
            </div>
        );
    }


    if (!currentUser) {
        return (
            <div className="flex items-center justify-center min-h-screen bg-slate-100 dark:bg-gray-900">
                <div className="w-full max-w-md p-8 space-y-8 bg-white rounded-2xl shadow-lg dark:bg-slate-800">
                    <div>
                        <h2 className="text-3xl font-extrabold text-center text-gray-900 dark:text-white">
                            مستخرج الإجازات
                        </h2>
                        <p className="mt-2 text-center text-sm text-gray-600 dark:text-gray-400">
                            مرحبًا بك، يرجى تسجيل الدخول
                        </p>
                    </div>
                    <form className="mt-8 space-y-6" onSubmit={handleLogin}>
                        <div className="rounded-md shadow-sm -space-y-px">
                            <div>
                                <input
                                    id="username"
                                    name="username"
                                    type="text"
                                    required
                                    className="relative block w-full px-3 py-2 text-gray-900 placeholder-gray-500 bg-gray-50 border border-gray-300 rounded-t-md focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 focus:z-10 sm:text-sm dark:bg-slate-700 dark:border-slate-600 dark:placeholder-gray-400 dark:text-white"
                                    placeholder="اسم المستخدم"
                                    value={usernameInput}
                                    onChange={(e) => setUsernameInput(e.target.value)}
                                />
                            </div>
                            <div>
                                <input
                                    id="password"
                                    name="password"
                                    type="password"
                                    required
                                    className="relative block w-full px-3 py-2 text-gray-900 placeholder-gray-500 bg-gray-50 border border-gray-300 rounded-b-md focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 focus:z-10 sm:text-sm dark:bg-slate-700 dark:border-slate-600 dark:placeholder-gray-400 dark:text-white"
                                    placeholder="كلمة المرور"
                                    value={passwordInput}
                                    onChange={(e) => setPasswordInput(e.target.value)}
                                />
                            </div>
                        </div>

                        {loginError && (
                            <p className="text-sm text-center text-red-500">{loginError}</p>
                        )}

                        <div>
                            <button
                                type="submit"
                                disabled={isLoading}
                                className="relative flex justify-center w-full px-4 py-2 text-sm font-medium text-white bg-indigo-600 border border-transparent rounded-md group hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 disabled:opacity-50"
                            >
                                {isLoading ? 'جاري الدخول...' : 'تسجيل الدخول'}
                            </button>
                        </div>
                    </form>
                </div>
            </div>
        );
    }
    
    // FIX: Replaced `JSX.Element` with `React.ReactElement` to resolve "Cannot find namespace 'JSX'".
    const navItems: { view: ViewType; label: string; icon: React.ReactElement; roles: ('admin' | 'supervisor' | 'employee')[] }[] = [
        { view: 'upload', label: 'تحميل ملف جديد', icon: <UploadIcon className="w-5 h-5" />, roles: ['admin', 'supervisor'] },
        { view: 'table', label: 'عرض البيانات العام', icon: <TableCellsIcon className="w-5 h-5" />, roles: ['admin', 'supervisor'] },
        { view: 'summary', label: 'ملخص الإجازات', icon: <DocumentTextIcon className="w-5 h-5" />, roles: ['admin', 'supervisor', 'employee'] },
        { view: 'rankedSummary', label: 'الملخص المرتب', icon: <TrophyIcon className="w-5 h-5" />, roles: ['admin', 'supervisor'] },
        { view: 'monthlyReport', label: 'التقرير الشهري', icon: <CalendarDaysIcon className="w-5 h-5" />, roles: ['admin', 'supervisor'] },
        { view: 'employeeManagement', label: 'إدارة الموظفين', icon: <UsersIcon className="w-5 h-5" />, roles: ['admin'] },
    ];

    const availableNavItems = navItems.filter(item => {
        if (item.roles.includes('employee') && typeof currentUser === 'object') return true;
        if (item.roles.includes('admin') && currentUser === 'admin') return true;
        if (item.roles.includes('supervisor') && currentUser === 'supervisor') return true;
        return false;
    });

    const renderHeader = () => {
        const viewTitles: Record<ViewType, string> = {
            upload: 'تحميل ملف جديد',
            table: 'عرض البيانات',
            summary: 'ملخص الإجازات',
            rankedSummary: 'الملخص المرتب',
            monthlyReport: 'التقرير الشهري',
            employeeManagement: 'إدارة الموظفين',
        };

        let currentUserName = 'مستخدم';
        if (currentUser === 'admin') currentUserName = 'مسؤول النظام';
        else if (currentUser === 'supervisor') currentUserName = 'مشرف';
        else if (currentUser && typeof currentUser === 'object') currentUserName = currentUser.name;

        return (
            <header className="sticky top-0 z-30 flex items-center justify-between px-4 py-3 bg-white border-b shadow-sm dark:bg-slate-800 dark:border-slate-700 sm:px-6 lg:px-8">
                <div className="flex items-center gap-4">
                    <button onClick={() => setIsSidebarOpen(!isSidebarOpen)} className="p-2 text-gray-500 rounded-md lg:hidden hover:bg-gray-100 dark:text-gray-400 dark:hover:bg-slate-700">
                        <Bars3Icon className="w-6 h-6" />
                    </button>
                    <h1 className="text-xl font-semibold text-gray-800 dark:text-white">{viewTitles[view]}</h1>
                </div>
                <div className="flex items-center gap-4">
                    <button onClick={toggleTheme} className="p-2 text-gray-500 rounded-full hover:bg-gray-100 dark:text-gray-400 dark:hover:bg-slate-700">
                        {theme === 'light' ? <MoonIcon className="w-6 h-6" /> : <SunIcon className="w-6 h-6" />}
                    </button>
                    <div className="text-sm text-right">
                        <div className="font-medium text-gray-800 dark:text-gray-200">{currentUserName}</div>
                    </div>
                    <button onClick={handleLogout} className="flex items-center gap-2 px-3 py-2 text-sm text-red-600 rounded-md hover:bg-red-50 dark:text-red-400 dark:hover:bg-red-900/20">
                        <LogoutIcon className="w-5 h-5" />
                        <span className="hidden sm:inline">خروج</span>
                    </button>
                </div>
            </header>
        );
    };

    const renderSidebar = () => (
        <aside className={`fixed inset-y-0 right-0 z-40 flex-shrink-0 w-64 bg-slate-900 text-white transform transition-transform duration-300 lg:relative lg:translate-x-0 ${isSidebarOpen ? 'translate-x-0' : 'translate-x-full'}`}>
            <div className="flex flex-col h-full">
                <div className="flex items-center justify-center h-20 border-b border-slate-700">
                    <h1 className="text-2xl font-bold text-center">مستخرج الإجازات</h1>
                </div>
                <nav className="flex-1 px-4 py-6 space-y-2">
                    {availableNavItems.map((item) => (
                        <button
                            key={item.view}
                            onClick={() => { setView(item.view); setIsSidebarOpen(false); }}
                            className={`flex items-center w-full gap-3 px-4 py-2.5 rounded-lg text-base font-medium transition-colors ${view === item.view ? 'bg-indigo-600' : 'hover:bg-slate-700'}`}
                        >
                            {item.icon}
                            <span>{item.label}</span>
                        </button>
                    ))}
                </nav>
                <div className="px-4 py-4 mt-auto border-t border-slate-700">
                    <div className="relative" ref={deleteMenuRef}>
                        <button onClick={() => setIsDeleteMenuOpen(prev => !prev)} className="flex items-center w-full gap-3 px-4 py-2.5 text-sm text-red-400 rounded-lg hover:bg-red-900/20">
                            <TrashIcon className="w-5 h-5"/>
                            <span>إدارة البيانات</span>
                        </button>
                        {isDeleteMenuOpen && (
                            <div className="absolute bottom-full right-0 mb-2 w-56 bg-white dark:bg-slate-700 rounded-md shadow-lg ring-1 ring-black ring-opacity-5">
                                <div className="py-1">
                                    {currentUser === 'admin' && (
                                        <button onClick={handleResetBalances} className="block w-full px-4 py-2 text-sm text-right text-gray-700 dark:text-gray-200 hover:bg-gray-100 dark:hover:bg-slate-600">
                                            إعادة تعيين أرصدة الموظفين
                                        </button>
                                    )}
                                    <button onClick={handleDeleteLastUpload} className="block w-full px-4 py-2 text-sm text-right text-gray-700 dark:text-gray-200 hover:bg-gray-100 dark:hover:bg-slate-600">
                                        حذف آخر ملف تم تحميله
                                    </button>
                                    {currentUser === 'admin' && (
                                    <button onClick={handleClearAllData} className="block w-full px-4 py-2 text-sm text-right text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20">
                                        حذف جميع البيانات نهائياً
                                    </button>
                                    )}
                                </div>
                            </div>
                        )}
                    </div>
                </div>
            </div>
        </aside>
    );

    const renderContent = () => {
        // Each view is rendered here based on the `view` state
        switch (view) {
            case 'upload':
                return (
                    <div className="max-w-4xl p-4 mx-auto sm:p-6 lg:p-8">
                        <div
                            onDragOver={handleDragOver}
                            onDrop={handleDrop}
                            className="flex flex-col items-center justify-center p-12 text-center border-4 border-dashed rounded-lg border-slate-300 dark:border-slate-600 hover:border-indigo-500 dark:hover:border-indigo-400 transition-colors"
                        >
                            <UploadIcon className="w-16 h-16 mx-auto text-slate-400 dark:text-slate-500" />
                            <h3 className="mt-4 text-xl font-semibold text-gray-900 dark:text-white">اسحب وأفلت الملف هنا</h3>
                            <p className="mt-1 text-slate-500 dark:text-slate-400">أو</p>
                            <button
                                type="button"
                                onClick={() => fileInputRef.current?.click()}
                                className="mt-4 px-6 py-2 text-sm font-medium text-white bg-indigo-600 rounded-md hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
                            >
                                اختر ملفًا
                            </button>
                            <input
                                type="file"
                                ref={fileInputRef}
                                onChange={handleFileChange}
                                className="hidden"
                                accept=".xlsx, .xls, .csv, .pdf"
                            />
                            <p className="mt-4 text-xs text-slate-500 dark:text-slate-400">
                                أنواع الملفات المدعومة: Excel (XLSX, XLS), CSV, PDF
                            </p>
                        </div>
                        {error && <div className="p-4 mt-4 text-red-700 bg-red-100 border-l-4 border-red-500" role="alert"><p className="font-bold">خطأ</p><p>{error}</p></div>}
                        
                        <div className="mt-8">
                            <h3 className="text-lg font-semibold text-gray-800 dark:text-white">الملفات التي تم تحميلها</h3>
                            {uploadHistory.length > 0 ? (
                                <ul className="mt-4 space-y-2">
                                {uploadHistory.map((upload) => (
                                    <li key={upload.uploadId} className="flex items-center justify-between p-3 bg-white rounded-lg shadow-sm dark:bg-slate-800">
                                    <span className="text-gray-700 dark:text-gray-300">{upload.fileName}</span>
                                    <button
                                        onClick={() => handleDeleteUploadById(upload.uploadId, upload.fileName)}
                                        className="p-1 text-red-500 rounded-full hover:bg-red-100 dark:hover:bg-red-900/20"
                                        aria-label={`حذف ملف ${upload.fileName}`}
                                    >
                                        <TrashIcon className="w-5 h-5" />
                                    </button>
                                    </li>
                                ))}
                                </ul>
                            ) : (
                                <p className="mt-4 text-center text-gray-500 dark:text-gray-400">لم يتم تحميل أي ملفات بعد.</p>
                            )}
                        </div>
                    </div>
                );
            case 'table':
                return (
                     <div className="p-4 sm:p-6 lg:p-8">
                        <div className="p-4 mb-4 bg-white rounded-lg shadow dark:bg-slate-800">
                             <div className="flex flex-wrap items-center justify-between gap-4">
                                <div className="relative flex-grow">
                                    <input
                                        type="text"
                                        placeholder="بحث في الجدول..."
                                        value={searchTerm}
                                        onChange={(e) => setSearchTerm(e.target.value)}
                                        className="w-full p-2 pl-10 border rounded-md dark:bg-slate-700 dark:border-slate-600"
                                    />
                                    <SearchIcon className="absolute w-5 h-5 text-gray-400 transform -translate-y-1/2 top-1/2 right-3" />
                                </div>
                                <button onClick={handleDownloadCSV} className="flex items-center gap-2 px-4 py-2 text-white bg-green-600 rounded-md hover:bg-green-700">
                                    <DownloadIcon className="w-5 h-5" />
                                    <span>تحميل CSV</span>
                                </button>
                            </div>
                        </div>
                        <div className="overflow-x-auto bg-white rounded-lg shadow dark:bg-slate-800">
                            <table className="w-full text-sm text-right text-gray-500 dark:text-gray-400">
                                <thead className="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-slate-700 dark:text-gray-400">
                                    <tr>
                                        {headers.map((header) => (
                                            <th key={header} scope="col" className="px-6 py-3">
                                                <div className="flex items-center justify-end gap-2 cursor-pointer" onClick={() => requestSort(header)}>
                                                    {header}
                                                    <SortIcon direction={sortConfig?.key === header ? sortConfig.direction : undefined} />
                                                </div>
                                            </th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {processedData.map((row, index) => (
                                        <tr key={index} className="bg-white border-b dark:bg-slate-800 dark:border-slate-700 hover:bg-gray-50 dark:hover:bg-slate-900/50">
                                            {headers.map((header) => (
                                                <td key={header} className="px-6 py-4">{String(row[header] ?? '')}</td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                );
            case 'summary': {
                const EmployeeCard = ({ employee }: { employee: EmployeeSummary }) => {
                    const { monthlyLeaves, otherLeaves } = useMemo(() => {
                        const monthly: Record<string, LeaveSummary[]> = {};
                        const other: LeaveSummary[] = [];
                        employee.leaves.forEach(leave => {
                            if (leave.monthYear) {
                                if (!monthly[leave.monthYear]) {
                                    monthly[leave.monthYear] = [];
                                }
                                monthly[leave.monthYear].push(leave);
                            } else {
                                other.push(leave);
                            }
                        });
                        return { monthlyLeaves: monthly, otherLeaves: other };
                    }, [employee.leaves]);
        
                    const renderLeaveItem = (leave: LeaveSummary, key: string | number) => (
                        <div key={key}>
                            <div className="flex justify-between font-semibold">
                                <span className="text-gray-800 dark:text-gray-200">{leave.type}</span>
                                <span className="text-indigo-600 dark:text-indigo-400">{formatLeaveCount(leave.dayCount, leave.hourCount)}</span>
                            </div>
                            <p className="text-xs text-gray-500 dark:text-gray-400 ltr-text">{leave.dateDetails}</p>
                        </div>
                    );
        
                    const renderCardContent = () => {
                        if (selectedMonth !== 'all' && selectedYear !== 'all') {
                            // Flat view for single month selection
                            return employee.leaves.map((leave, index) => renderLeaveItem(leave, index));
                        }
        
                        // Monthly breakdown for "All Months" or "All Time" view
                        const sortedMonths = Object.keys(monthlyLeaves).sort().reverse();
                        return (
                            <>
                                {sortedMonths.map(monthKey => {
                                    const [year, month] = monthKey.split('-');
                                    return (
                                    <div key={monthKey} className="pb-4 mb-4 border-b border-gray-200 dark:border-gray-700 last:border-b-0 last:mb-0 last:pb-0">
                                        <h4 className="font-semibold text-gray-700 dark:text-gray-300">
                                            {new Date(parseInt(year), parseInt(month) - 1).toLocaleString('ar-EG', { month: 'long', year: 'numeric' })}
                                        </h4>
                                        <div className="mt-2 space-y-3">
                                            {monthlyLeaves[monthKey].map((leave, index) => renderLeaveItem(leave, `month-${index}`))}
                                        </div>
                                    </div>
                                    );
                                })}
                                {otherLeaves.length > 0 && (
                                    <div className="mt-4 pt-4 border-t border-gray-200 dark:border-gray-700">
                                        <div className="space-y-3">
                                            {otherLeaves.map((leave, index) => renderLeaveItem(leave, `other-${index}`))}
                                        </div>
                                    </div>
                                )}
                            </>
                        );
                    };
        
                    const totalDeducted = getRegularAndTimeBasedTotalDays(employee);
                    const finalBalance = employee.currentBalance ?? 0;
        
                    return (
                        <div className="flex flex-col bg-white rounded-lg shadow-md dark:bg-slate-800 transition-all duration-300 hover:shadow-xl">
                            <div className="flex items-center p-4 border-b border-gray-200 dark:border-gray-700">
                                {employee.photo ? (
                                    <img src={employee.photo} alt={employee.name} className="object-cover w-16 h-16 rounded-full" />
                                ) : (
                                    <UserCircleIcon className="w-16 h-16 text-gray-300 dark:text-gray-500" />
                                )}
                                <div className="mr-4">
                                    <h3 className="text-lg font-bold text-gray-900 dark:text-white">{employee.name}</h3>
                                    <p className="text-sm text-gray-500 dark:text-gray-400">{employee.workplace}</p>
                                </div>
                            </div>
                            
                            <div className="flex-grow p-4 space-y-4">
                                {employee.leaves.length > 0 ? renderCardContent() : <p className="text-sm text-center text-gray-500 dark:text-gray-400">لا توجد إجازات مسجلة</p>}
                            </div>
                            
                            {employee.leaves.length > 0 && (
                                <div className="p-4 mt-auto space-y-2 bg-gray-50 rounded-b-lg dark:bg-slate-900/50">
                                    <div className="flex items-center justify-between">
                                        <span className="text-sm font-semibold text-gray-600 dark:text-gray-300">الرصيد الأولي:</span>
                                        <span className="text-lg font-bold text-blue-600 dark:text-blue-400">{toArabicNumerals(employee.initialBalance ?? 0)} يوم</span>
                                    </div>
                                    <div className="flex items-center justify-between">
                                        <span className="text-sm font-semibold text-gray-600 dark:text-gray-300">مجموع الاعتيادي والزمني:</span>
                                        <span className="text-lg font-bold text-red-500">{formatDaysArabic(totalDeducted)}</span>
                                    </div>
                                    <div className="flex items-center justify-between pt-2 mt-2 border-t border-gray-200 dark:border-gray-700">
                                        <span className="text-sm font-semibold text-gray-600 dark:text-gray-300">الرصيد المتبقي:</span>
                                        <span className={`text-lg font-bold ${finalBalance < 0 ? 'text-red-500' : 'text-green-600'}`}>{toArabicNumerals(finalBalance)} يوم</span>
                                    </div>
                                </div>
                            )}
                        </div>
                    );
                };

                return (
                    <div className="p-4 sm:p-6 lg:p-8">
                        {(currentUser === 'admin' || currentUser === 'supervisor') && (
                        <div className="p-4 mb-6 bg-white rounded-lg shadow dark:bg-slate-800">
                             <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
                                <div className="relative flex-grow">
                                    <input type="text" placeholder="...بحث عن موظف" value={summarySearchTerm} onChange={(e) => setSummarySearchTerm(e.target.value)} className="w-full p-2 pl-10 border rounded-md dark:bg-slate-700 dark:border-slate-600" />
                                    <SearchIcon className="absolute w-5 h-5 text-gray-400 transform -translate-y-1/2 top-1/2 right-3" />
                                </div>
                                <div>
                                    <select value={selectedYear} onChange={(e) => setSelectedYear(e.target.value)} className="w-full p-2 border rounded-md dark:bg-slate-700 dark:border-slate-600">
                                        <option value="all">كل السنوات</option>
                                        {availableYears.map(year => <option key={year} value={year}>{toArabicNumerals(year)}</option>)}
                                    </select>
                                </div>
                                <div>
                                    <select value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} disabled={selectedYear === 'all'} className="w-full p-2 border rounded-md disabled:opacity-50 dark:bg-slate-700 dark:border-slate-600">
                                        <option value="all">كل الشهور</option>
                                        {Array.from({ length: 12 }, (_, i) => i + 1).map(month => <option key={month} value={month}>{new Date(2000, month-1).toLocaleString('ar-EG', { month: 'long' })}</option>)}
                                    </select>
                                </div>
                                <div className="flex items-center gap-2">
                                     <button onClick={handleCopyToClipboard} className="flex items-center justify-center flex-1 gap-2 px-4 py-2 text-white bg-blue-600 rounded-md hover:bg-blue-700">
                                        <ClipboardDocumentIcon className="w-5 h-5" /> <span>نسخ</span>
                                    </button>
                                     <button onClick={handleExportSummaryToWord} className="flex items-center justify-center flex-1 gap-2 px-4 py-2 text-white bg-green-600 rounded-md hover:bg-green-700">
                                        <DocumentTextIcon className="w-5 h-5" /> <span>Word</span>
                                    </button>
                                </div>
                            </div>
                        </div>
                        )}

                        {showCopyNotification && <div className="fixed top-20 right-5 p-3 text-white bg-green-500 rounded-lg shadow-lg z-50">تم نسخ البيانات بنجاح!</div>}
                        
                        <div ref={summaryContentRef} className="printable-area">
                            <div className="grid grid-cols-1 gap-6 md:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-4 printable-content">
                                {displayedSummary.map((employee) => (
                                    <EmployeeCard key={employee.name} employee={employee} />
                                ))}
                            </div>
                        </div>
                    </div>
                );
            }
            case 'rankedSummary':
                // ... (rest of the file remains the same)
                return (
                    <div className="p-4 sm:p-6 lg:p-8">
                       <div className="p-4 mb-6 bg-white rounded-lg shadow dark:bg-slate-800">
                         <div className="flex items-center justify-between">
                            <div className="flex items-center gap-4">
                               <span className="font-semibold">ترتيب حسب:</span>
                               <button onClick={() => setRankedSortOrder('byDays')} className={`px-4 py-2 rounded-md ${rankedSortOrder === 'byDays' ? 'bg-indigo-600 text-white' : 'bg-gray-200 dark:bg-slate-700'}`}>
                                <BarsArrowUpIcon className="inline w-5 h-5 ml-2"/>
                                عدد الأيام
                               </button>
                               <button onClick={() => setRankedSortOrder('alphabetical')} className={`px-4 py-2 rounded-md ${rankedSortOrder === 'alphabetical' ? 'bg-indigo-600 text-white' : 'bg-gray-200 dark:bg-slate-700'}`}>
                                <AlphabeticalSortIcon className="inline w-5 h-5 ml-2"/>
                                أبجدي
                               </button>
                            </div>
                            <button onClick={handleExportWord} className="flex items-center gap-2 px-4 py-2 text-white bg-blue-600 rounded-md hover:bg-blue-700">
                                <DownloadIcon className="w-5 h-5"/>
                                تصدير Word
                            </button>
                         </div>
                       </div>
                       
                       <div className="space-y-8">
                           {rankedSortOrder === 'byDays' ? (
                               Object.keys(groupedRankedSummary).sort((a,b) => parseInt(a) - parseInt(b)).map(dayCount => (
                                <div key={dayCount} className="p-4 bg-white rounded-lg shadow dark:bg-slate-800">
                                    <h2 className="mb-4 text-xl font-bold text-center text-indigo-700 dark:text-indigo-400">
                                        الموظفون الحاصلون على {formatDaysArabic(parseInt(dayCount))}
                                    </h2>
                                    <div className="overflow-x-auto">
                                        <table className="w-full text-sm text-right text-gray-500 dark:text-gray-400">
                                            <thead className="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-slate-700 dark:text-gray-400">
                                                <tr>
                                                   {rankedSummaryTableHeaders.map(h => <th key={h} className="px-6 py-3">{h}</th>)}
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {groupedRankedSummary[parseInt(dayCount)].map((employee, index) => {
                                                    const totalDays = getRegularAndTimeBasedTotalDays(employee);
                                                    const leaveMap = new Map<string, LeaveSummary>();
                                                    employee.leaves.forEach(l => leaveMap.set(l.type, l));
                                                    
                                                    return (
                                                        <tr key={employee.name} className="bg-white border-b dark:bg-slate-800 dark:border-slate-700">
                                                            <td className="px-6 py-4">{toArabicNumerals(index + 1)}</td>
                                                            <td className="px-6 py-4 font-medium text-gray-900 whitespace-nowrap dark:text-white">{employee.name}</td>
                                                            <td className="px-6 py-4">{formatDaysArabic(totalDays)}</td>
                                                            {rankedSummaryLeaveTypes.map(type => {
                                                                const leave = leaveMap.get(type);
                                                                let content = ' ';
                                                                if (leave) {
                                                                    if (type === 'ملخص الزمنيات') {
                                                                        content = leave.dayCount > 0 ? `(${toArabicNumerals(leave.dayCount)}) يوم` : ' ';
                                                                    } else {
                                                                        content = leave.dateDetails;
                                                                    }
                                                                }
                                                                return <td key={type} className="px-6 py-4 ltr-text">{content}</td>;
                                                            })}
                                                        </tr>
                                                    );
                                                })}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                               ))
                           ) : (
                             <div className="p-4 bg-white rounded-lg shadow dark:bg-slate-800">
                                <h2 className="mb-4 text-xl font-bold text-center text-indigo-700 dark:text-indigo-400">
                                    ملخص أبجدي
                                </h2>
                                <div className="overflow-x-auto">
                                        <table className="w-full text-sm text-right text-gray-500 dark:text-gray-400">
                                            <thead className="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-slate-700 dark:text-gray-400">
                                                <tr>
                                                   {rankedSummaryTableHeaders.map(h => <th key={h} className="px-6 py-3">{h}</th>)}
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {alphabeticallySortedSummary.map((employee, index) => {
                                                    const totalDays = getRegularAndTimeBasedTotalDays(employee);
                                                    const leaveMap = new Map<string, LeaveSummary>();
                                                    employee.leaves.forEach(l => leaveMap.set(l.type, l));
                                                    
                                                    return (
                                                        <tr key={employee.name} className="bg-white border-b dark:bg-slate-800 dark:border-slate-700">
                                                            <td className="px-6 py-4">{toArabicNumerals(index + 1)}</td>
                                                            <td className="px-6 py-4 font-medium text-gray-900 whitespace-nowrap dark:text-white">{employee.name}</td>
                                                            <td className="px-6 py-4">{formatDaysArabic(totalDays)}</td>
                                                            {rankedSummaryLeaveTypes.map(type => {
                                                                const leave = leaveMap.get(type);
                                                                let content = ' ';
                                                                if (leave) {
                                                                     if (type === 'ملخص الزمنيات') {
                                                                        content = leave.dayCount > 0 ? `(${toArabicNumerals(leave.dayCount)}) يوم` : ' ';
                                                                    } else {
                                                                        content = leave.dateDetails;
                                                                    }
                                                                }
                                                                return <td key={type} className="px-6 py-4 ltr-text">{content}</td>;
                                                            })}
                                                        </tr>
                                                    );
                                                })}
                                            </tbody>
                                        </table>
                                    </div>
                             </div>
                           )}
                           {employeesWithSickLeave.length > 0 && (
                                <div className="p-4 bg-white rounded-lg shadow dark:bg-slate-800">
                                    <h2 className="mb-4 text-xl font-bold text-center text-indigo-700 dark:text-indigo-400">
                                        الموظفون الحاصلون على اجازة مرضية (5 أيام أو أقل)
                                    </h2>
                                    <table className="w-full text-sm text-right text-gray-500 dark:text-gray-400">
                                       <thead className="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-slate-700 dark:text-gray-400">
                                            <tr>
                                                <th className="px-6 py-3">ت</th>
                                                <th className="px-6 py-3">الاسم</th>
                                                <th className="px-6 py-3">المدة</th>
                                                <th className="px-6 py-3">التاريخ</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {employeesWithSickLeave.map((emp, index) => {
                                                const totalDays = emp.sickLeaves.reduce((sum, l) => sum + l.dayCount, 0);
                                                const dateDetails = emp.sickLeaves.map(l => l.dateDetails).join(' | ');
                                                return (
                                                     <tr key={emp.name} className="bg-white border-b dark:bg-slate-800 dark:border-slate-700">
                                                        <td className="px-6 py-4">{toArabicNumerals(index + 1)}</td>
                                                        <td className="px-6 py-4 font-medium text-gray-900 whitespace-nowrap dark:text-white">{emp.name}</td>
                                                        <td className="px-6 py-4">{formatDaysArabic(totalDays)}</td>
                                                        <td className="px-6 py-4 ltr-text">{dateDetails}</td>
                                                    </tr>
                                                );
                                            })}
                                        </tbody>
                                    </table>
                                </div>
                           )}
                       </div>
                    </div>
                );
             case 'employeeManagement':
                return (
                     <div className="p-4 sm:p-6 lg:p-8">
                        {/* Add new employee form */}
                        <div className="p-4 mb-6 bg-white rounded-lg shadow dark:bg-slate-800">
                            <h2 className="text-xl font-semibold">إضافة موظف جديد</h2>
                            <form onSubmit={handleAddEmployee} className="grid grid-cols-1 gap-4 mt-4 sm:grid-cols-2 lg:grid-cols-4">
                                <input type="text" placeholder="الاسم الكامل" value={newEmployeeName} onChange={e => setNewEmployeeName(e.target.value)} required className="p-2 border rounded-md dark:bg-slate-700 dark:border-slate-600"/>
                                <input type="text" placeholder="مكان العمل" value={newEmployeeWorkplace} onChange={e => setNewEmployeeWorkplace(e.target.value)} className="p-2 border rounded-md dark:bg-slate-700 dark:border-slate-600"/>
                                <input type="number" placeholder="الرصيد الأولي" value={newEmployeeBalance} onChange={e => setNewEmployeeBalance(e.target.value)} required className="p-2 border rounded-md dark:bg-slate-700 dark:border-slate-600"/>
                                <input type="number" placeholder="ساعات العمل اليومية" value={newEmployeeWorkdayHours} onChange={e => setNewEmployeeWorkdayHours(e.target.value)} required className="p-2 border rounded-md dark:bg-slate-700 dark:border-slate-600"/>
                                <input type="text" value={newEmployeeUsername} onChange={e => setNewEmployeeUsername(e.target.value)} required className="p-2 border rounded-md dark:bg-slate-700 dark:border-slate-600"/>
                                <input type="text" value={newEmployeePassword} onChange={e => setNewEmployeePassword(e.target.value)} required className="p-2 border rounded-md dark:bg-slate-700 dark:border-slate-600"/>
                                <div className="flex items-center col-span-1 sm:col-span-2 lg:col-span-1">
                                    <label className="px-4 py-2 text-white bg-gray-500 rounded-md cursor-pointer hover:bg-gray-600">
                                        <span>اختر صورة</span>
                                        <input type="file" className="hidden" accept="image/*" onChange={handleNewEmployeePhotoChange} />
                                    </label>
                                    {newEmployeePhoto && <img src={newEmployeePhoto} alt="Preview" className="w-12 h-12 mr-4 rounded-full"/>}
                                </div>
                                <button type="submit" className="px-4 py-2 text-white bg-indigo-600 rounded-md sm:col-span-2 lg:col-span-1 hover:bg-indigo-700">إضافة موظف</button>
                            </form>
                        </div>
                         {/* Employee list */}
                        <div className="p-4 bg-white rounded-lg shadow dark:bg-slate-800">
                            <div className="flex flex-wrap items-center justify-between gap-4 mb-4">
                                 <h2 className="text-xl font-semibold">قائمة الموظفين</h2>
                                 <div className="relative flex-grow max-w-sm">
                                    <input type="text" placeholder="...بحث" value={employeeSearchTerm} onChange={(e) => setEmployeeSearchTerm(e.target.value)} className="w-full p-2 pl-10 border rounded-md dark:bg-slate-700 dark:border-slate-600"/>
                                    <SearchIcon className="absolute w-5 h-5 text-gray-400 transform -translate-y-1/2 top-1/2 right-3" />
                                </div>
                                 <button onClick={handleExportUserList} className="flex items-center gap-2 px-4 py-2 text-white bg-green-600 rounded-md hover:bg-green-700">
                                    <DownloadIcon className="w-5 h-5" />
                                    <span>تصدير قائمة المستخدمين</span>
                                </button>
                            </div>
                            <div className="overflow-x-auto">
                                <table className="w-full text-sm text-right text-gray-500 dark:text-gray-400">
                                     <thead className="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-slate-700 dark:text-gray-400">
                                        <tr>
                                            <th className="px-6 py-3">الصورة</th>
                                            <th className="px-6 py-3">الاسم</th>
                                            <th className="px-6 py-3">مكان العمل</th>
                                            <th className="px-6 py-3">الرصيد</th>
                                            <th className="px-6 py-3">اسم المستخدم</th>
                                            <th className="px-6 py-3">كلمة المرور</th>
                                            <th className="px-6 py-3">الإجراءات</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {filteredEmployees.map(emp => (
                                            <tr key={emp.id} className="bg-white border-b dark:bg-slate-800 dark:border-slate-700">
                                                <td className="px-6 py-4">
                                                    {editingEmployee?.id === emp.id ? (
                                                        <>
                                                         <input type="file" accept="image/*" onChange={handleEditPhotoChange} className="text-xs"/>
                                                         {editingEmployee.photo && <img src={editingEmployee.photo} className="w-10 h-10 mt-2 rounded-full"/>}
                                                        </>
                                                    ) : (
                                                        emp.photo ? <img src={emp.photo} alt={emp.name} className="w-10 h-10 rounded-full"/> : <UserCircleIcon className="w-10 h-10 text-gray-300"/>
                                                    )}
                                                </td>
                                                <td className="px-6 py-4">{editingEmployee?.id === emp.id ? <input className="p-1 border rounded-md dark:bg-slate-700" value={editingEmployee.name} onChange={e => setEditingEmployee({...editingEmployee, name: e.target.value})}/> : emp.name}</td>
                                                <td className="px-6 py-4">{editingEmployee?.id === emp.id ? <input className="p-1 border rounded-md dark:bg-slate-700" value={editingEmployee.workplace || ''} onChange={e => setEditingEmployee({...editingEmployee, workplace: e.target.value})}/> : emp.workplace}</td>
                                                <td className="px-6 py-4">{editingEmployee?.id === emp.id ? <input type="number" className="w-20 p-1 border rounded-md dark:bg-slate-700" value={editingEmployee.balance} onChange={e => setEditingEmployee({...editingEmployee, balance: Number(e.target.value)})}/> : emp.balance}</td>
                                                <td className="px-6 py-4">{editingEmployee?.id === emp.id ? <input className="p-1 border rounded-md dark:bg-slate-700" value={editingEmployee.username} onChange={e => setEditingEmployee({...editingEmployee, username: e.target.value})}/> : emp.username}</td>
                                                <td className="px-6 py-4">{editingEmployee?.id === emp.id ? <input className="p-1 border rounded-md dark:bg-slate-700" value={editingEmployee.password} onChange={e => setEditingEmployee({...editingEmployee, password: e.target.value})}/> : emp.password}</td>
                                                <td className="px-6 py-4">
                                                    {editingEmployee?.id === emp.id ? (
                                                        <div className="flex gap-2">
                                                            <button onClick={handleSaveEdit} className="text-green-500">حفظ</button>
                                                            <button onClick={handleCancelEdit} className="text-gray-500">إلغاء</button>
                                                        </div>
                                                    ) : (
                                                        <div className="flex gap-2">
                                                            <button onClick={() => handleStartEdit(emp)}><PencilIcon className="w-5 h-5 text-blue-500"/></button>
                                                            <button onClick={() => handleDeleteEmployee(emp.id)}><TrashIcon className="w-5 h-5 text-red-500"/></button>
                                                        </div>
                                                    )}
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                     </div>
                );
            case 'monthlyReport':
                return (
                     <div className="p-4 sm:p-6 lg:p-8">
                        {/* Report Generation Form */}
                        <div className="p-4 mb-6 bg-white rounded-lg shadow dark:bg-slate-800">
                            <h2 className="text-xl font-semibold">إنشاء تقرير</h2>
                            <div className="grid grid-cols-1 gap-4 mt-4 sm:grid-cols-2 lg:grid-cols-4">
                                <div>
                                    <label className="block text-sm">السنة</label>
                                    <select value={reportYear} onChange={e => setReportYear(parseInt(e.target.value))} className="w-full p-2 border rounded-md dark:bg-slate-700 dark:border-slate-600">
                                        {Array.from({ length: 10 }, (_, i) => new Date().getFullYear() - i).map(year => <option key={year} value={year}>{toArabicNumerals(year)}</option>)}
                                    </select>
                                </div>
                                <div>
                                    <label className="block text-sm">الشهر</label>
                                    <select value={reportMonth} onChange={e => setReportMonth(parseInt(e.target.value))} className="w-full p-2 border rounded-md dark:bg-slate-700 dark:border-slate-600">
                                        {Array.from({ length: 12 }, (_, i) => i + 1).map(month => <option key={month} value={month}>{new Date(2000, month-1).toLocaleString('ar-EG', { month: 'long' })}</option>)}
                                    </select>
                                </div>
                                <div className="self-end">
                                    <button onClick={handleGenerateMonthlyReport} className="w-full px-4 py-2 text-white bg-blue-600 rounded-md hover:bg-blue-700">
                                        إنشاء تقرير شهري
                                    </button>
                                </div>
                                 <div className="self-end">
                                    <button onClick={handleGenerateAlphabeticalReport} className="w-full px-4 py-2 text-white bg-purple-600 rounded-md hover:bg-purple-700">
                                        إنشاء تقرير أبجدي
                                    </button>
                                </div>
                            </div>
                        </div>
                        
                        {/* Report Display */}
                        {(monthlyReportData.length > 0 || alphabeticalReportData.length > 0) && (
                            <div className="p-4 bg-white rounded-lg shadow dark:bg-slate-800">
                                <div className="flex flex-wrap items-center justify-between gap-4 mb-4">
                                     <h2 className="text-xl font-semibold">
                                        {showAlphabeticalReport ? 'التقرير الأبجدي الشامل' : `التقرير الشهري لـ ${toArabicNumerals(reportMonth)}/${toArabicNumerals(reportYear)}`}
                                     </h2>
                                     <div>
                                        <select value={reportWorkplaceFilter} onChange={e => setReportWorkplaceFilter(e.target.value)} className="p-2 border rounded-md dark:bg-slate-700 dark:border-slate-600">
                                            {workplaces.map(wp => <option key={wp} value={wp}>{wp === 'all' ? 'كل أماكن العمل' : wp}</option>)}
                                        </select>
                                    </div>
                                     <button onClick={handleExportMonthlyReportToWord} className="flex items-center gap-2 px-4 py-2 text-white bg-green-600 rounded-md hover:bg-green-700">
                                        <DownloadIcon className="w-5 h-5"/> تصدير Word
                                    </button>
                                </div>
                                <div className="flex flex-wrap gap-4 mb-4">
                                    {Object.keys(leaveTypeLabels).map(key => (
                                        <label key={key} className="flex items-center gap-2">
                                            <input type="checkbox" checked={reportColumnVisibility[key as keyof typeof reportColumnVisibility]} onChange={() => setReportColumnVisibility(prev => ({...prev, [key]: !prev[key as keyof typeof reportColumnVisibility]}))} />
                                            {leaveTypeLabels[key]}
                                        </label>
                                    ))}
                                </div>
                                <div className="overflow-x-auto">
                                   <table className="w-full text-sm text-right text-gray-500 dark:text-gray-400">
                                       <thead className="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-slate-700 dark:text-gray-400">
                                           <tr>
                                               <th className="px-2 py-3">الاسم</th>
                                               <th className="px-2 py-3">مكان العمل</th>
                                               <th className="px-2 py-3">الرصيد الأولي</th>
                                               {reportColumnVisibility.regular && <th className="px-2 py-3">{showAlphabeticalReport ? 'اعتيادية (أيام)' : 'اعتيادية'}</th>}
                                               {reportColumnVisibility.hourly && <th className="px-2 py-3">زمنية (أيام/ساعات)</th>}
                                               {reportColumnVisibility.sick && <th className="px-2 py-3">مرضية</th>}
                                               {reportColumnVisibility.long && <th className="px-2 py-3">طويلة</th>}
                                               {reportColumnVisibility.travel && <th className="px-2 py-3">سفر</th>}
                                               {reportColumnVisibility.absent && <th className="px-2 py-3">غائب</th>}
                                               <th className="px-2 py-3">الرصيد المتبقي</th>
                                           </tr>
                                       </thead>
                                       <tbody>
                                            {(showAlphabeticalReport ? alphabeticalReportData : monthlyReportData).filter(row => reportWorkplaceFilter === 'all' || row.workplace === reportWorkplaceFilter).map(row => (
                                                 <tr key={row.name} className="bg-white border-b dark:bg-slate-800 dark:border-slate-700">
                                                    <td className="px-2 py-4 font-medium">{row.name}</td>
                                                    <td className="px-2 py-4">{row.workplace}</td>
                                                    <td className="px-2 py-4">{toArabicNumerals(row.initialBalance)}</td>
                                                    {reportColumnVisibility.regular && <td className="px-2 py-4">{row.regularLeaves.dates || ''}</td>}
                                                    {reportColumnVisibility.hourly && <td className="px-2 py-4">{formatLeaveCount(row.hourlyLeaves.days, row.hourlyLeaves.hours)}</td>}
                                                    {reportColumnVisibility.sick && <td className="px-2 py-4">{row.sickLeave.dateRange || ''}</td>}
                                                    {reportColumnVisibility.long && <td className="px-2 py-4">{row.longLeave.dateRange || ''}</td>}
                                                    {reportColumnVisibility.travel && <td className="px-2 py-4">{row.travelLeave.dateRange || ''}</td>}
                                                    {reportColumnVisibility.absent && <td className="px-2 py-4">{row.absentLeave.dateRange || ''}</td>}
                                                    <td className={`px-2 py-4 font-bold ${row.finalBalance < 0 ? 'text-red-500' : 'text-green-600'}`}>{toArabicNumerals(row.finalBalance)}</td>
                                                 </tr>
                                            ))}
                                       </tbody>
                                   </table>
                                </div>
                            </div>
                        )}
                     </div>
                );

            default: return <div>عرض غير معروف</div>;
        }
    };
    

    return (
        <div className="flex h-screen bg-slate-100 dark:bg-gray-950">
            {isSidebarOpen && <div onClick={() => setIsSidebarOpen(false)} className="fixed inset-0 z-30 bg-black/50 lg:hidden"></div>}
            {renderSidebar()}
            <div className="flex flex-col flex-1 w-full overflow-y-auto">
                {renderHeader()}
                <main className="flex-1">
                    {renderContent()}
                </main>
            </div>
        </div>
    );
};
