import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Users,
  DollarSign,
  Calendar,
  Clock,
  AlertTriangle,
  Upload,
  Download,
  Plus,
  Trash2,
  Edit3,
  RotateCcw,
  CheckCircle2,
  Search,
  Filter,
  Eye,
  EyeOff,
  Sun,
  Award,
  ChevronRight,
  Sparkles,
  FileSpreadsheet,
  X,
  Sliders,
  ShieldCheck,
  Building2,
  TrendingUp,
  Activity,
  Layers,
  HelpCircle
} from "lucide-react";

// =========================================================================
// 1. HARD-CODED PRELOADED EMPLOYEE ROSTER (ZERO SETUP REQUIRED)
// =========================================================================
const FACTORY_EMPLOYEE_MASTER = [
  { empId: "1", name: "DALPAT SINGH", dutyHours: 9.5, salary: 18000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "10", name: "SATYA NARYAN", dutyHours: 0, salary: 16000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "11", name: "KANHAIYA LAL MAHAWAR", dutyHours: 11, salary: 13000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "12", name: "AKASH NIHAL", dutyHours: 0, salary: 14000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "13", name: "GAJENDRA GURJAR", dutyHours: 9.5, salary: 16500, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "14", name: "SUNIL MAHAWAR", dutyHours: 10, salary: 18000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "15", name: "MUKESH SEN", dutyHours: 9.5, salary: 11500, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "16", name: "MUKESH MAHAWAR", dutyHours: 9.5, salary: 20500, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "17", name: "PRADEEP SAINI", dutyHours: 9.5, salary: 21000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "18", name: "GULSHAN VERMA", dutyHours: 9.5, salary: 13500, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "19", name: "MAHAVEER SHAKYAWAL", dutyHours: 10, salary: 16500, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "2", name: "LALIT MAHAWAR", dutyHours: 9.5, salary: 21000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "20", name: "CHANCHAL SEN", dutyHours: 8, salary: 16000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "21", name: "RAVINDRA MAHAWAR", dutyHours: 9.5, salary: 16500, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "22", name: "DILKHUSH MEENA", dutyHours: 0, salary: 0, ownTiffin: "No", bonus: "Yes", active: "No" },
  { empId: "23", name: "MAHAVEER GURJAR", dutyHours: 9.5, salary: 14500, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "24", name: "DEEPAK SAINI", dutyHours: 9.5, salary: 18000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "25", name: "GAJENDRA SAINI", dutyHours: 9.5, salary: 15200, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "26", name: "VINOD MEGHWAL", dutyHours: 9.5, salary: 18500, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "27", name: "AJAY KUSHWAHA", dutyHours: 0, salary: 0, ownTiffin: "No", bonus: "Yes", active: "No" },
  { empId: "28", name: "RAJESH REGAR", dutyHours: 9.5, salary: 15000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "29", name: "PRADEEP SINGH", dutyHours: 0, salary: 16000, ownTiffin: "No", bonus: "Yes", active: "No" },
  { empId: "3", name: "SONU SUMAN", dutyHours: 9.5, salary: 22000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "30", name: "VINAYAK GAUTAM", dutyHours: 9.5, salary: 17000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "31", name: "SANTOSH", dutyHours: 9.5, salary: 22000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "32", name: "DAYARAM SUMAN", dutyHours: 0, salary: 0, ownTiffin: "No", bonus: "Yes", active: "No" },
  { empId: "33", name: "AMIT MEHRA", dutyHours: 10, salary: 10000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "34", name: "MANISH VERMA", dutyHours: 0, salary: 0, ownTiffin: "No", bonus: "Yes", active: "No" },
  { empId: "35", name: "ROHIT DHARWAL", dutyHours: 9.5, salary: 25000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "36", name: "VINOD GURJAR", dutyHours: 0, salary: 12000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "37", name: "LOVE KUSH", dutyHours: 0, salary: 17000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "38", name: "ANOOP SHARMA", dutyHours: 0, salary: 0, ownTiffin: "No", bonus: "Yes", active: "No" },
  { empId: "39", name: "RAJESH GOCHER", dutyHours: 0, salary: 12500, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "4", name: "DEVESH PRAJAPATI", dutyHours: 9.5, salary: 15000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "40", name: "RAMJEET NAGAR", dutyHours: 9.5, salary: 17000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "41", name: "VISHNU GOCHER", dutyHours: 9.5, salary: 11500, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "42", name: "BRIJ MOHAN", dutyHours: 9.5, salary: 18000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "43", name: "SUNIL GOYAL", dutyHours: 9.5, salary: 17000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "44", name: "BALRAM MEGHWAL", dutyHours: 9.5, salary: 17000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "46", name: "VISHNU SHAKYAWAL", dutyHours: 0, salary: 17000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "47", name: "ANIL MAHAWAR", dutyHours: 0, salary: 21000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "48", name: "MUKESH REGAR", dutyHours: 10, salary: 24000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "49", name: "VISHAL SINGH", dutyHours: 9.5, salary: 16000, ownTiffin: "No", bonus: "Yes", active: "Yes" },
  { empId: "5", name: "ROHIT MEGHWAL", dutyHours: 10, salary: 13000, ownTiffin: "No", bonus: "Yes", active: "Yes" }
];

const LOCAL_STORAGE_KEY = "PAYROLL_MASTER";

// Currency Formatter
const formatINR = (val) => {
  const num = Math.round(Number(val) || 0);
  return "₹" + num.toLocaleString("en-IN");
};

// Decimal Formatter
const formatDec = (val, dec = 2) => {
  const num = Number(val) || 0;
  return num.toFixed(dec);
};

// Normalize ID strictly (strips accidental Excel trailing .0 and trims)
const cleanId = (val) => String(val || "").trim().replace(/\.0$/, "");

export default function App() {
  const [activeTab, setActiveTab] = useState("dashboard");
  const [hideInactive, setHideInactive] = useState(true);

  // Employee Master State initialized from localStorage or Factory Defaults
  const [master, setMaster] = useState(() => {
    try {
      const stored = localStorage.getItem(LOCAL_STORAGE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch {
      // fallback
    }
    return FACTORY_EMPLOYEE_MASTER;
  });

  // Sync Master to localStorage
  useEffect(() => {
    try {
      localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(master));
    } catch (err) {
      console.error("Failed to persist master to localStorage:", err);
    }
  }, [master]);

  // Attendance & Payroll State
  const [attendanceRecords, setAttendanceRecords] = useState([]);
  const [payrollBreakdown, setPayrollBreakdown] = useState([]);
  const [metaInfo, setMetaInfo] = useState({
    fileName: "DEMO_BIO_APRIL_2026.xlsx",
    fromDate: "01/04/2026",
    toDate: "30/04/2026",
    totalDays: 30,
    sundaysCount: 4,
    processedAt: new Date().toLocaleTimeString()
  });

  const [parseLogs, setParseLogs] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [attendanceFilter, setAttendanceFilter] = useState("ALL");

  // Modal State for Add / Edit Employee
  const [editingEmp, setEditingEmp] = useState(null);
  const [isModalOpen, setIsModalOpen] = useState(false);

  // =========================================================================
  // 2. PARSING HELPERS (TIME & DURATIONS)
  // =========================================================================
  const parseTimeToHours = (timeVal) => {
    if (timeVal === undefined || timeVal === null || timeVal === "") return null;
    if (typeof timeVal === "number") {
      // Excel fractional day (0.5 = 12:00)
      return Number((timeVal * 24).toFixed(2));
    }
    const str = String(timeVal).trim();
    if (str === "00:00" || str === "0:00" || str === "--" || str === "-") return null;
    const parts = str.split(":");
    if (parts.length >= 2) {
      const hh = parseInt(parts[0], 10) || 0;
      const mm = parseInt(parts[1], 10) || 0;
      const ss = parts[2] ? parseInt(parts[2], 10) || 0 : 0;
      if (hh === 0 && mm === 0 && ss === 0) return null;
      return Number((hh + mm / 60 + ss / 3600).toFixed(2));
    }
    return null;
  };

  const parseDurationHours = (durVal) => {
    if (!durVal || durVal === "00:00" || durVal === "--") return 0;
    if (typeof durVal === "number") {
      return Number((durVal * 24).toFixed(2));
    }
    const str = String(durVal).trim();
    const parts = str.split(":");
    if (parts.length >= 2) {
      const hh = parseInt(parts[0], 10) || 0;
      const mm = parseInt(parts[1], 10) || 0;
      return Number((hh + mm / 60).toFixed(2));
    }
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  };

  // Convert Excel Serial Date or String to valid date object
  const parseReportDate = (raw) => {
    if (!raw) return null;
    if (typeof raw === "number" && raw > 40000) {
      return new Date(Math.round((raw - 25569) * 86400 * 1000));
    }
    const s = String(raw).trim();
    const parts = s.split(/[\/\-\.]/);
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
      }
      return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
    }
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  };

  // =========================================================================
  // 3. INDUSTRIAL PAYROLL COMPUTATION ENGINE
  // =========================================================================
  const runPayrollEngine = (attendanceList, currentMaster, meta) => {
    const masterMap = new Map();
    currentMaster.forEach((m) => masterMap.set(cleanId(m.empId), m));

    const results = attendanceList.map((rec) => {
      const empId = cleanId(rec.empId);
      const profile = masterMap.get(empId) || {
        empId,
        name: rec.name || `EMPLOYEE #${empId}`,
        dutyHours: 9.5,
        salary: 15000,
        ownTiffin: "No",
        bonus: "Yes",
        active: "Yes"
      };

      const isActive = profile.active === "Yes";
      const isDayWorker = Number(profile.dutyHours) <= 0;
      const dutyHours = Number(profile.dutyHours) || 0;
      const monthlySalary = Number(profile.salary) || 0;
      const ownTiffin = profile.ownTiffin === "Yes";
      const bonusEligible = profile.bonus === "Yes";

      // If inactive, completely zero out metrics
      if (!isActive) {
        return {
          empId,
          name: profile.name,
          isActive: false,
          isDayWorker,
          dutyHours,
          effectiveDuty: 0,
          monthlySalary,
          dailySalary: 0,
          hourlyRate: 0,
          normalPresent: 0,
          absentDays: 0,
          actualTotalHours: 0,
          sundayWorkedDays: 0,
          sundayExtraDays: 0,
          sundayDeductionHours: 0,
          normalWorkingDays: 0,
          requiredHours: 0,
          adjustedActualHours: 0,
          shortHours: 0,
          extraHours: 0,
          absentDeduction: 0,
          shortHourDeduction: 0,
          extraHourSalary: 0,
          sundayExtraSalary: 0,
          sundayDeductionSalary: 0,
          bonusDaysCount: 0,
          bonusDaySalary: 0,
          totalAdjustment: 0,
          netSalary: 0,
          remarks: "INACTIVE / EXCLUDED"
        };
      }

      // Rates (Rule: Always divide monthly salary by 30)
      const dailySalary = monthlySalary / 30;
      const hourlyRate = isDayWorker ? 0 : dailySalary / (dutyHours || 1);
      const effectiveDuty = isDayWorker ? 0 : ownTiffin ? Math.max(0, dutyHours - 0.5) : dutyHours;

      let normalPresent = 0;
      let absentDays = 0;
      let actualTotalHours = 0;
      let sundayWorkedDays = 0;
      let sundayExtraDays = 0;
      let sundayDeductionHours = 0;
      let hasSunday14HoursFlag = false;

      // Evaluate days
      rec.days.forEach((d) => {
        const { isSunday, evaluatedStatus, actualHours } = d;
        actualTotalHours += actualHours || 0;

        if (isSunday) {
          if (evaluatedStatus === "PW" || actualHours > 0) {
            sundayWorkedDays += 1;
            if (actualHours > 0 && actualHours < 8) {
              sundayExtraDays += 1;
              if (!isDayWorker) sundayDeductionHours += 6.5;
            } else if (actualHours >= 8 && actualHours < 14) {
              sundayExtraDays += 2;
              if (!isDayWorker) sundayDeductionHours += 12.0;
            } else if (actualHours >= 14) {
              sundayExtraDays += 2;
              if (!isDayWorker) sundayDeductionHours += 12.0;
              hasSunday14HoursFlag = true;
            } else {
              // Status was PW without duration
              sundayExtraDays += 1;
            }
          }
        } else {
          // Normal Working Day
          if (evaluatedStatus === "P") {
            normalPresent += 1.0;
          } else if (evaluatedStatus === "HD") {
            normalPresent += 0.5;
            absentDays += 0.5;
          } else if (evaluatedStatus === "A") {
            absentDays += 1.0;
          } else if (evaluatedStatus === "W") {
            // Midweek off
          } else {
            absentDays += 1.0;
          }
        }
      });

      // Working Days & Hours requirements
      const normalWorkingDays = Math.max(0, meta.totalDays - meta.sundaysCount - absentDays);
      const requiredHours = isDayWorker ? 0 : normalWorkingDays * effectiveDuty;

      // Adjusted actual punch hours
      const adjustedActualHours = isDayWorker
        ? actualTotalHours
        : Math.max(0, actualTotalHours - sundayDeductionHours);

      let shortHours = 0;
      let extraHours = 0;
      if (!isDayWorker) {
        if (adjustedActualHours < requiredHours) {
          shortHours = Number((requiredHours - adjustedActualHours).toFixed(2));
        } else {
          extraHours = Number((adjustedActualHours - requiredHours).toFixed(2));
        }
      }

      // Deductions and Earnings
      const absentDeduction = isDayWorker
        ? absentDays * dailySalary
        : absentDays * dutyHours * hourlyRate;

      const shortHourDeduction = shortHours * hourlyRate;
      const extraHourSalary = extraHours * hourlyRate;
      const sundayExtraSalary = isDayWorker
        ? sundayExtraDays * dailySalary
        : sundayExtraDays * dutyHours * hourlyRate;

      const sundayDeductionSalary = sundayDeductionHours * hourlyRate;

      // Attendance Bonus Calculation
      let bonusDaysCount = 0;
      if (bonusEligible) {
        if (normalPresent >= 26) {
          bonusDaysCount = 3;
        } else if (normalPresent >= 24) {
          bonusDaysCount = 2;
        } else if (normalPresent >= 23) {
          bonusDaysCount = 1;
        } else {
          bonusDaysCount = 0;
        }
      }
      const bonusDaySalary = bonusDaysCount * dailySalary;

      // Final Net Salary
      let netSalary = 0;
      if (isDayWorker) {
        netSalary = normalPresent * dailySalary + sundayExtraSalary + bonusDaySalary;
      } else {
        netSalary =
          monthlySalary -
          absentDeduction -
          shortHourDeduction +
          extraHourSalary +
          sundayExtraSalary -
          sundayDeductionSalary +
          bonusDaySalary;
      }

      // Adjustment = bonusDaySalary + sundayExtraSalary + extraHourSalary - shortHourDeduction
      const totalAdjustment =
        bonusDaySalary + sundayExtraSalary + extraHourSalary - shortHourDeduction;

      const remarksArr = [];
      if (hasSunday14HoursFlag) remarksArr.push("CHECK 14+ HRS SUNDAY");
      if (isDayWorker) remarksArr.push("Day Worker");
      if (shortHours > 12) remarksArr.push("Excess Short Hours");
      if (remarksArr.length === 0) remarksArr.push("Verified OK");

      return {
        empId,
        name: profile.name,
        isActive: true,
        isDayWorker,
        dutyHours,
        effectiveDuty,
        monthlySalary,
        dailySalary,
        hourlyRate,
        normalPresent,
        absentDays,
        actualTotalHours,
        sundayWorkedDays,
        sundayExtraDays,
        sundayDeductionHours,
        normalWorkingDays,
        requiredHours,
        adjustedActualHours,
        shortHours,
        extraHours,
        absentDeduction,
        shortHourDeduction,
        extraHourSalary,
        sundayExtraSalary,
        sundayDeductionSalary,
        bonusDaysCount,
        bonusDaySalary,
        totalAdjustment,
        netSalary: Math.max(0, Math.round(netSalary)),
        remarks: remarksArr.join(" | ")
      };
    });

    // Natural integer order
    results.sort((a, b) => parseInt(a.empId, 10) - parseInt(b.empId, 10));
    return results;
  };

  // Re-run payroll if master changes
  useEffect(() => {
    if (attendanceRecords.length > 0) {
      const recalculated = runPayrollEngine(attendanceRecords, master, metaInfo);
      setPayrollBreakdown(recalculated);
    }
  }, [master]);

  // =========================================================================
  // 4. BIOMETRIC EXCEL PARSER (LEGACY MERGED FORMAT COMPATIBILITY)
  // =========================================================================
  const parseBiometricWorkbook = async (file) => {
    setIsProcessing(true);
    const logs = [];

    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rawMatrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

      if (!rawMatrix || rawMatrix.length < 5) {
        throw new Error("Uploaded Excel sheet does not contain enough rows to be a biometric report.");
      }

      // Step 1: Detect From & To date in rows 1 to 15
      let fromDateStr = "";
      let toDateStr = "";
      let detectedTotalDays = 30;
      let fromDateObj = null;

      for (let r = 0; r < Math.min(15, rawMatrix.length); r++) {
        const row = rawMatrix[r];
        for (let c = 0; c < row.length; c++) {
          const val = String(row[c] || "").trim();
          if (/from\s*:/i.test(val)) {
            const mFrom = val.match(/from\s*:\s*([^\s]+)/i);
            const mTo = val.match(/to\s*:\s*([^\s]+)/i);
            if (mFrom) fromDateStr = mFrom[1];
            if (mTo) toDateStr = mTo[1];

            // If "To:" is in subsequent cells
            if (!toDateStr) {
              for (let c2 = c + 1; c2 < Math.min(c + 6, row.length); c2++) {
                const val2 = String(row[c2] || "").trim();
                const mTo2 = val2.match(/(?:to\s*:|^to\b)\s*([^\s]+)/i);
                if (mTo2) {
                  toDateStr = mTo2[1];
                  break;
                }
              }
            }
          }
        }
      }

      if (fromDateStr) {
        fromDateObj = parseReportDate(fromDateStr);
      }
      if (!fromDateObj) {
        fromDateObj = new Date(2026, 3, 1); // default April 2026
        fromDateStr = "01/04/2026";
        toDateStr = "30/04/2026";
      }

      // Step 2: Detect Sequence 1, 2, 3...N
      let dayHeaderRowIdx = -1;
      let dayDataStartColIdx = -1;

      for (let r = 0; r < Math.min(30, rawMatrix.length); r++) {
        const row = rawMatrix[r];
        for (let c = 0; c < row.length - 2; c++) {
          const n1 = parseInt(row[c], 10);
          const n2 = parseInt(row[c + 1], 10);
          const n3 = parseInt(row[c + 2], 10);

          if (n1 === 1 && n2 === 2 && n3 === 3) {
            dayHeaderRowIdx = r;
            dayDataStartColIdx = c;
            let count = 0;
            while (parseInt(row[c + count], 10) === count + 1) {
              count++;
            }
            detectedTotalDays = count;
            break;
          }
        }
        if (dayHeaderRowIdx !== -1) break;
      }

      if (dayHeaderRowIdx === -1) {
        logs.push("Could not strictly locate sequence 1,2,3. Used default column index 3 with 30 days.");
        dayDataStartColIdx = 3;
        detectedTotalDays = 30;
      }

      // Identify Calendar Sundays
      const sundayColIndices = [];
      for (let d = 0; d < detectedTotalDays; d++) {
        const dObj = new Date(fromDateObj);
        dObj.setDate(fromDateObj.getDate() + d);
        if (dObj.getDay() === 0) {
          sundayColIndices.push(d);
        }
      }

      // Step 3: Scan Employees & Punch Rows
      const parsedRecords = [];
      const masterById = new Map();
      master.forEach((m) => masterById.set(cleanId(m.empId), m));

      for (let r = 0; r < rawMatrix.length; r++) {
        const row = rawMatrix[r];
        const rowStr = row.map((x) => String(x || "")).join(" ");

        if (/emp\s*id\s*:/i.test(rowStr)) {
          let rawEmpId = "";
          let rawEmpName = "";

          // Extract ID and Name
          for (let c = 0; c < row.length; c++) {
            const cell = String(row[c] || "").trim();
            if (/emp\s*id\s*:/i.test(cell)) {
              const match = cell.match(/emp\s*id\s*:\s*([^\s]+)/i);
              if (match && match[1]) {
                rawEmpId = match[1];
              } else {
                // Find next non-empty cell
                for (let nextC = c + 1; nextC < Math.min(c + 5, row.length); nextC++) {
                  const nxt = String(row[nextC] || "").trim();
                  if (nxt) {
                    rawEmpId = nxt;
                    break;
                  }
                }
              }
            }

            if (/employee\s*name\s*:/i.test(cell) || /emp\s*name\s*:/i.test(cell)) {
              const nameMatch = cell.match(/(?:employee|emp)\s*name\s*:\s*(.*)/i);
              if (nameMatch && nameMatch[1].trim()) {
                rawEmpName = nameMatch[1].trim();
              } else {
                for (let nextC = c + 1; nextC < Math.min(c + 8, row.length); nextC++) {
                  const nxt = String(row[nextC] || "").trim();
                  if (nxt && !/type|dept/i.test(nxt)) {
                    rawEmpName = nxt;
                    break;
                  }
                }
              }
            }
          }

          const empId = cleanId(rawEmpId);
          if (!empId) continue;

          // Lookup in master to ensure strict naming integrity
          const profile = masterById.get(empId);
          const finalName = profile ? profile.name : rawEmpName || `EMPLOYEE #${empId}`;

          // Find Punch Rows in this employee block
          let statusRow = null;
          let firstInRow = null;
          let lastOutRow = null;
          let totalHrsRow = null;

          for (let subR = r + 1; subR < Math.min(r + 14, rawMatrix.length); subR++) {
            const subRow = rawMatrix[subR];
            const labelZone = subRow
              .slice(0, Math.max(4, dayDataStartColIdx))
              .map((x) => String(x || "").toLowerCase().trim())
              .join(" ");

            if (/emp\s*id\s*:/i.test(labelZone)) break; // next employee arrived

            if (!statusRow && labelZone.includes("status")) {
              statusRow = subRow;
            } else if (!firstInRow && (labelZone.includes("first in") || labelZone.includes("in time"))) {
              firstInRow = subRow;
            } else if (!lastOutRow && (labelZone.includes("last out") || labelZone.includes("out time"))) {
              lastOutRow = subRow;
            } else if (!totalHrsRow && (labelZone.includes("total hrs") || labelZone.includes("work hrs") || labelZone.includes("total time"))) {
              totalHrsRow = subRow;
            }
          }

          const isDayWorker = profile ? Number(profile.dutyHours) <= 0 : false;
          const daysData = [];

          for (let d = 0; d < detectedTotalDays; d++) {
            const col = dayDataStartColIdx + d;
            const rawStatus = statusRow ? String(statusRow[col] || "").trim() : "";
            const rawIn = firstInRow ? String(firstInRow[col] || "").trim() : "";
            const rawOut = lastOutRow ? String(lastOutRow[col] || "").trim() : "";
            const rawHrs = totalHrsRow ? String(totalHrsRow[col] || "").trim() : "";

            const isSunday = sundayColIndices.includes(d);
            const inDec = parseTimeToHours(rawIn);
            const outDec = parseTimeToHours(rawOut);
            let totDec = parseDurationHours(rawHrs);

            if (totDec === 0 && inDec !== null && outDec !== null && outDec > inDec) {
              totDec = Number((outDec - inDec).toFixed(2));
            }

            // Normalization of status
            let finalStatus = "A";
            const sUpper = rawStatus.toUpperCase();

            if (isDayWorker && (sUpper === "E" || sUpper === "P") && inDec !== null) {
              finalStatus = "P";
            } else if (sUpper === "PW" || (isSunday && (sUpper === "P" || sUpper === "PW" || totDec > 0))) {
              finalStatus = isSunday && totDec === 0 && sUpper !== "PW" ? "W" : "PW";
            } else if (sUpper === "HD" || sUpper === "HP") {
              finalStatus = "HD";
            } else if (sUpper === "P") {
              finalStatus = "P";
            } else if (sUpper === "W" || sUpper === "WO") {
              finalStatus = "W";
            } else if (isSunday) {
              finalStatus = "W";
            } else {
              finalStatus = "A";
            }

            daysData.push({
              dayIndex: d + 1,
              isSunday,
              rawStatus,
              firstIn: inDec !== null ? rawIn : "--",
              lastOut: outDec !== null ? rawOut : "--",
              firstInDec: inDec,
              lastOutDec: outDec,
              actualHours: totDec,
              evaluatedStatus: finalStatus
            });
          }

          parsedRecords.push({
            empId,
            name: finalName,
            days: daysData
          });
        }
      }

      if (parsedRecords.length === 0) {
        throw new Error("No employee punch blocks with 'Emp ID :' were identified in the file.");
      }

      // Sort naturally
      parsedRecords.sort((a, b) => parseInt(a.empId, 10) - parseInt(b.empId, 10));

      const newMeta = {
        fileName: file.name,
        fromDate: fromDateStr,
        toDate: toDateStr || `${detectedTotalDays}/Month`,
        totalDays: detectedTotalDays,
        sundaysCount: sundayColIndices.length,
        processedAt: new Date().toLocaleTimeString()
      };

      setMetaInfo(newMeta);
      setAttendanceRecords(parsedRecords);

      const payroll = runPayrollEngine(parsedRecords, master, newMeta);
      setPayrollBreakdown(payroll);

      logs.push(`Successfully loaded ${parsedRecords.length} employees for ${detectedTotalDays} days.`);
      setParseLogs(logs);
      setActiveTab("dashboard");
    } catch (err) {
      logs.push(`Error: ${err.message}`);
      setParseLogs(logs);
    } finally {
      setIsProcessing(false);
    }
  };

  // =========================================================================
  // 5. TEST GENERATOR (INSTANT 1-CLICK DEMO LOADER)
  // =========================================================================
  const loadInstantFactoryDemo = () => {
    setIsProcessing(true);
    const totalDays = 30;
    const sundays = [5, 12, 19, 26]; // 1-indexed days

    const sampleAttendance = master.map((emp) => {
      const isDayWorker = emp.dutyHours <= 0;
      const days = [];

      for (let d = 1; d <= totalDays; d++) {
        const isSunday = sundays.includes(d);
        let status = "P";
        let inTime = "09:15";
        let outTime = "19:00";
        let hrs = isDayWorker ? 9.5 : emp.dutyHours;

        if (isSunday) {
          // Some worked Sunday OT, some off
          if (d === 12 && parseInt(emp.empId, 10) % 3 === 0) {
            status = "PW";
            hrs = 8.5;
            inTime = "09:00";
            outTime = "17:30";
          } else {
            status = "W";
            hrs = 0;
            inTime = "--";
            outTime = "--";
          }
        } else {
          // Standard month punches
          if (d === 7 && parseInt(emp.empId, 10) % 4 === 0) {
            status = "A";
            hrs = 0;
            inTime = "--";
            outTime = "--";
          } else if (d === 18 && parseInt(emp.empId, 10) % 5 === 0) {
            status = "HD";
            hrs = Math.max(0, hrs / 2);
            outTime = "14:00";
          }
        }

        days.push({
          dayIndex: d,
          isSunday,
          rawStatus: status,
          firstIn: inTime,
          lastOut: outTime,
          firstInDec: inTime === "--" ? null : 9.25,
          lastOutDec: outTime === "--" ? null : 19.0,
          actualHours: hrs,
          evaluatedStatus: status
        });
      }

      return {
        empId: emp.empId,
        name: emp.name,
        days
      };
    });

    const meta = {
      fileName: "FACTORY_SYNTHETIC_APRIL_2026.xlsx",
      fromDate: "01/04/2026",
      toDate: "30/04/2026",
      totalDays: 30,
      sundaysCount: 4,
      processedAt: new Date().toLocaleTimeString()
    };

    setAttendanceRecords(sampleAttendance);
    setMetaInfo(meta);
    const payroll = runPayrollEngine(sampleAttendance, master, meta);
    setPayrollBreakdown(payroll);
    setParseLogs(["Instant factory demonstration dataset loaded successfully!"]);
    setIsProcessing(false);
  };

  // Initialize with demo if not loaded yet
  useEffect(() => {
    if (attendanceRecords.length === 0) {
      loadInstantFactoryDemo();
    }
  }, []);

  // =========================================================================
  // 6. EXPORTING COMPLETE PAYROLL TO EXCEL VIA SHEETJS
  // =========================================================================
  const exportPayrollToExcel = () => {
    if (!payrollBreakdown.length) return;

    const exportRows = payrollBreakdown
      .filter((row) => (hideInactive ? row.isActive : true))
      .map((r) => ({
        "Emp ID": r.empId,
        "Employee Name": r.name,
        "Status": r.isActive ? "ACTIVE" : "INACTIVE",
        "Category": r.isDayWorker ? "Day Worker (Daily Rate)" : "Duty Worker",
        "Duty Hrs": r.dutyHours,
        "Effective Duty": r.effectiveDuty,
        "Monthly Salary (₹)": r.monthlySalary,
        "Daily Rate (₹)": Math.round(r.dailySalary),
        "Hourly Rate (₹)": Number(r.hourlyRate.toFixed(2)),
        "Normal Present Days": r.normalPresent,
        "Absent Days": r.absentDays,
        "Sunday OT Days": r.sundayExtraDays,
        "Punch Actual Hrs": Number(r.actualTotalHours.toFixed(2)),
        "Required Hrs": Number(r.requiredHours.toFixed(2)),
        "Short Hrs": r.shortHours,
        "Extra Hrs": r.extraHours,
        "Bonus Days": r.bonusDaysCount,
        "Absent Ded (₹)": Math.round(r.absentDeduction),
        "Short Hrs Ded (₹)": Math.round(r.shortHourDeduction),
        "Extra Hrs Pay (₹)": Math.round(r.extraHourSalary),
        "Sunday Pay (₹)": Math.round(r.sundayExtraSalary),
        "Bonus Pay (₹)": Math.round(r.bonusDaySalary),
        "Adjustment (AD) (₹)": Math.round(r.totalAdjustment),
        "NET PAYABLE (₹)": r.netSalary,
        "Remarks / Audit": r.remarks
      }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Payroll_Calculation");
    XLSX.writeFile(wb, `Industrial_Payroll_${metaInfo.totalDays}Days_${metaInfo.fileName.replace(/\.xlsx?/i, "")}.xlsx`);
  };

  // =========================================================================
  // 7. KPI AGGREGATIONS (RESPECTING INACTIVE FILTERING STRICTLY)
  // =========================================================================
  const kpiStats = useMemo(() => {
    // Only ACTIVE employees count towards financials & metrics
    const activeRows = payrollBreakdown.filter((p) => p.isActive);

    const totalActiveEmp = activeRows.length;
    const totalNet = activeRows.reduce((acc, r) => acc + r.netSalary, 0);
    const totalBase = activeRows.reduce((acc, r) => acc + r.monthlySalary, 0);
    const totalAdjustment = activeRows.reduce((acc, r) => acc + r.totalAdjustment, 0);
    const totalAbsent = activeRows.reduce((acc, r) => acc + r.absentDays, 0);
    const totalShortHours = activeRows.reduce((acc, r) => acc + r.shortHours, 0);
    const totalExtraHours = activeRows.reduce((acc, r) => acc + r.extraHours, 0);
    const totalSundayOTDays = activeRows.reduce((acc, r) => acc + r.sundayExtraDays, 0);
    const totalBonusPay = activeRows.reduce((acc, r) => acc + r.bonusDaySalary, 0);

    const totalDayWorkers = activeRows.filter((r) => r.isDayWorker).length;
    const totalMonthlyWorkers = totalActiveEmp - totalDayWorkers;

    return {
      totalActiveEmp,
      totalNet,
      totalBase,
      totalAdjustment,
      totalAbsent,
      totalShortHours,
      totalExtraHours,
      totalSundayOTDays,
      totalBonusPay,
      totalDayWorkers,
      totalMonthlyWorkers
    };
  }, [payrollBreakdown]);

  // Filtered Payroll View
  const displayedPayroll = useMemo(() => {
    return payrollBreakdown.filter((r) => {
      if (hideInactive && !r.isActive) return false;
      if (!searchQuery) return true;
      const q = searchQuery.toLowerCase();
      return (
        r.empId.toLowerCase().includes(q) ||
        r.name.toLowerCase().includes(q) ||
        r.remarks.toLowerCase().includes(q)
      );
    });
  }, [payrollBreakdown, hideInactive, searchQuery]);

  // Filtered Attendance Daily Punches
  const flattenedAttendance = useMemo(() => {
    const list = [];
    attendanceRecords.forEach((emp) => {
      const profile = master.find((m) => cleanId(m.empId) === cleanId(emp.empId));
      const isActive = profile ? profile.active === "Yes" : true;
      if (hideInactive && !isActive) return;

      emp.days.forEach((d) => {
        list.push({
          empId: emp.empId,
          name: emp.name,
          isActive,
          ...d
        });
      });
    });
    return list;
  }, [attendanceRecords, master, hideInactive]);

  const displayedAttendance = useMemo(() => {
    return flattenedAttendance.filter((rec) => {
      const matchSearch =
        !searchQuery ||
        rec.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
        rec.empId.toLowerCase().includes(searchQuery.toLowerCase()) ||
        `Day ${rec.dayIndex}`.toLowerCase().includes(searchQuery.toLowerCase());

      const matchStatus =
        attendanceFilter === "ALL" ? true : rec.evaluatedStatus === attendanceFilter;

      return matchSearch && matchStatus;
    });
  }, [flattenedAttendance, searchQuery, attendanceFilter]);

  // Master Table Filter
  const displayedMaster = useMemo(() => {
    return [...master]
      .filter((m) => {
        if (hideInactive && m.active === "No") return false;
        if (!searchQuery) return true;
        const q = searchQuery.toLowerCase();
        return m.empId.toLowerCase().includes(q) || m.name.toLowerCase().includes(q);
      })
      .sort((a, b) => parseInt(a.empId, 10) - parseInt(b.empId, 10));
  }, [master, hideInactive, searchQuery]);

  // =========================================================================
  // 8. CRUD ACTIONS FOR EMPLOYEE MASTER
  // =========================================================================
  const handleSaveEmployee = (emp) => {
    const targetId = cleanId(emp.empId);
    if (!targetId || !emp.name) {
      alert("Employee ID and Name are required.");
      return;
    }

    if (editingEmp) {
      setMaster((prev) =>
        prev.map((item) => (cleanId(item.empId) === cleanId(editingEmp.empId) ? emp : item))
      );
    } else {
      const exists = master.some((item) => cleanId(item.empId) === targetId);
      if (exists) {
        alert(`Employee with ID "${targetId}" already exists!`);
        return;
      }
      setMaster((prev) => [...prev, emp]);
    }
    setIsModalOpen(false);
    setEditingEmp(null);
  };

  const handleDeleteEmployee = (id) => {
    if (window.confirm(`Delete Employee ID: ${id} from Master?`)) {
      setMaster((prev) => prev.filter((m) => cleanId(m.empId) !== cleanId(id)));
    }
  };

  const handleResetToFac