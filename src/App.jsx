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
  ShieldCheck,
  Building2,
  TrendingUp,
  Activity,
  Layers,
  FileUp
} from "lucide-react";

// =========================================================================
// 1. HARD-CODED PRELOADED 75+ EMPLOYEE MASTER ROSTER
// =========================================================================
const INITIAL_MASTER = [
  { empId: "1", name: "DALPAT SINGH", dutyHours: 9.5, salary: 18000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "10", name: "SATYA NARYAN", dutyHours: 0, salary: 16000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "11", name: "KANHAIYA LAL MAHAWAR", dutyHours: 11, salary: 13000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "12", name: "AKASH NIHAL", dutyHours: 0, salary: 14000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "13", name: "GAJENDRA GURJAR", dutyHours: 9.5, salary: 16500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "14", name: "SUNIL MAHAWAR", dutyHours: 10, salary: 18000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "15", name: "MUKESH SEN", dutyHours: 9.5, salary: 11500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "16", name: "MUKESH MAHAWAR", dutyHours: 9.5, salary: 20500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "17", name: "PRADEEP SAINI", dutyHours: 9.5, salary: 21000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "18", name: "GULSHAN VERMA", dutyHours: 9.5, salary: 13500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "19", name: "MAHAVEER SHAKYAWAL", dutyHours: 10, salary: 16500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "2", name: "LALIT MAHAWAR", dutyHours: 9.5, salary: 21000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "20", name: "CHANCHAL SEN", dutyHours: 8, salary: 16000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "21", name: "RAVINDRA MAHAWAR", dutyHours: 9.5, salary: 16500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "22", name: "DILKHUSH MEENA", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "23", name: "MAHAVEER GURJAR", dutyHours: 9.5, salary: 14500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "24", name: "DEEPAK SAINI", dutyHours: 9.5, salary: 18000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "25", name: "GAJENDRA SAINI", dutyHours: 9.5, salary: 15200, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "26", name: "VINOD MEGHWAL", dutyHours: 9.5, salary: 18500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "27", name: "AJAY KUSHWAHA", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "28", name: "RAJESH REGAR", dutyHours: 9.5, salary: 15000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "29", name: "PRADEEP SINGH", dutyHours: 0, salary: 16000, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "3", name: "SONU SUMAN", dutyHours: 9.5, salary: 22000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "30", name: "VINAYAK GAUTAM", dutyHours: 9.5, salary: 17000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "31", name: "SANTOSH", dutyHours: 9.5, salary: 22000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "32", name: "DAYARAM SUMAN", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "33", name: "AMIT MEHRA", dutyHours: 10, salary: 10000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "34", name: "MANISH VERMA", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "35", name: "ROHIT DHARWAL", dutyHours: 9.5, salary: 25000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "36", name: "VINOD GURJAR", dutyHours: 0, salary: 12000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "37", name: "LOVE KUSH", dutyHours: 0, salary: 17000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "38", name: "ANOOP SHARMA", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "39", name: "RAJESH GOCHER", dutyHours: 0, salary: 12500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "4", name: "DEVESH PRAJAPATI", dutyHours: 9.5, salary: 15000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "40", name: "RAMJEET NAGAR", dutyHours: 9.5, salary: 17000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "41", name: "VISHNU GOCHER", dutyHours: 9.5, salary: 11500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "42", name: "BRIJ MOHAN", dutyHours: 9.5, salary: 18000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "43", name: "SUNIL GOYAL", dutyHours: 9.5, salary: 17000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "44", name: "BALRAM MEGHWAL", dutyHours: 9.5, salary: 17000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "46", name: "VISHNU SHAKYAWAL", dutyHours: 0, salary: 17000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "47", name: "ANIL MAHAWAR", dutyHours: 0, salary: 21000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "48", name: "MUKESH REGAR", dutyHours: 10, salary: 24000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "49", name: "VISHAL SINGH", dutyHours: 9.5, salary: 16000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "5", name: "ROHIT MEGHWAL", dutyHours: 10, salary: 13000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "50", name: "LALIT SHRINGI", dutyHours: 0, salary: 14500, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "51", name: "SURENDRA SUMAN", dutyHours: 9.5, salary: 17000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "52", name: "KHUSH", dutyHours: 10, salary: 15500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "53", name: "RAHUL", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "54", name: "RAJENDRA RINKU", dutyHours: 0, salary: 16000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "56", name: "RAKESH VERMA", dutyHours: 0, salary: 15000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "57", name: "SURAJ SUMAN", dutyHours: 9.5, salary: 12000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "58", name: "PREMCHAND MEGHWAL", dutyHours: 0, salary: 17500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "59", name: "DWARKA LAL", dutyHours: 0, salary: 21000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "60", name: "PREM SHANKAR MEENA", dutyHours: 10, salary: 17000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "61", name: "RAJARAM MEENA", dutyHours: 10, salary: 13000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "62", name: "PURSHOTTAM", dutyHours: 9.5, salary: 18000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "66", name: "MANISH VERMA", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "67", name: "VIPIN VERMA", dutyHours: 9.5, salary: 5000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "68", name: "BABLU MEWARA", dutyHours: 9.5, salary: 10500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "69", name: "BABLESH MEGAWAL", dutyHours: 10, salary: 10500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "7", name: "KUNJ BIHARI", dutyHours: 10, salary: 18400, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "72", name: "GHANSHYAM BHATT", dutyHours: 0, salary: 15500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "73", name: "ARUN PRAJAPATI", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "74", name: "UMASHANKAR GWALA", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "77", name: "RAGHUVEER GOCHER", dutyHours: 0, salary: 21000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "78", name: "BALRAM GOCHER", dutyHours: 9.5, salary: 11000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "79", name: "RAVI SUMAN", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "8", name: "CHANDRA PRAKEASH CM", dutyHours: 10, salary: 16000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "80", name: "VINOD MALAV", dutyHours: 10, salary: 14000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "81", name: "GOLU GOCHER", dutyHours: 9.5, salary: 11000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "84", name: "SURAJ KUSHWAH", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "85", name: "RAVI CHOHAN", dutyHours: 10, salary: 11000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "86", name: "DEVRAJ GUJJAR", dutyHours: 10, salary: 11000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "88", name: "PRAKASH", dutyHours: 0, salary: 13500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "89", name: "KANHAIYA LAL SAINI", dutyHours: 0, salary: 8500, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "9", name: "JAMNA SHANKAR", dutyHours: 9.5, salary: 17000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "90", name: "VIKAS GAUTAM", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "91", name: "PIYUSH LALWANI", dutyHours: 10, salary: 12000, ownTiffin: "No", bonusDays: "Yes", active: "Yes", notes: "" },
  { empId: "92", name: "HARPAL BISHNOI", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" },
  { empId: "93", name: "GAURAV SINGH RATHORE", dutyHours: 0, salary: 0, ownTiffin: "No", bonusDays: "Yes", active: "No", notes: "" }
];

const LOCAL_STORAGE_KEY = "PAYROLL_MASTER";

// Helpers
const cleanId = (val) => String(val ?? "").trim().replace(/\.0$/, "");
const formatINR = (val) => "₹" + Math.round(Number(val) || 0).toLocaleString("en-IN");
const formatDec = (val, dec = 2) => (Number(val) || 0).toFixed(dec);

export default function App() {
  const [activeTab, setActiveTab] = useState("dashboard");
  const [showInactive, setShowInactive] = useState(false);

  // Load Employee Master from localStorage with fallback
  const [master, setMaster] = useState(() => {
    try {
      const stored = localStorage.getItem(LOCAL_STORAGE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch (e) {
      console.error("Master load error:", e);
    }
    return INITIAL_MASTER;
  });

  // Persist master state
  useEffect(() => {
    try {
      localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(master));
    } catch (e) {
      console.error("Master save error:", e);
    }
  }, [master]);

  // Attendance & Payroll State
  const [attendanceRecords, setAttendanceRecords] = useState([]);
  const [payrollBreakdown, setPayrollBreakdown] = useState([]);
  const [metaInfo, setMetaInfo] = useState({
    fileName: "BIOMETRIC_REPORT_APRIL_2026.xlsx",
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

  // Modals & Refs
  const [isEmpModalOpen, setIsEmpModalOpen] = useState(false);
  const [editingEmp, setEditingEmp] = useState(null);
  const fileMasterInputRef = useRef(null);

  // Time & Duration Helpers
  const parseTimeToHours = (timeVal) => {
    if (timeVal === undefined || timeVal === null || timeVal === "") return null;
    if (typeof timeVal === "number") return Number((timeVal * 24).toFixed(2));
    const str = String(timeVal).trim();
    if (["00:00", "0:00", "--", "-", ""].includes(str)) return null;
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
    if (typeof durVal === "number") return Number((durVal * 24).toFixed(2));
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

  const parseReportDate = (raw) => {
    if (!raw) return null;
    if (typeof raw === "number" && raw > 40000) {
      return new Date(Math.round((raw - 25569) * 86400 * 1000));
    }
    const s = String(raw).trim();
    const parts = s.split(/[\/\-\.]/);
    if (parts.length === 3) {
      if (parts[0].length === 4) return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
      return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
    }
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  };

  // =========================================================================
  // 2. INDUSTRIAL PAYROLL CALCULATION ENGINE
  // =========================================================================
  const calculatePayroll = (records, currentMaster, meta) => {
    const masterMap = new Map();
    currentMaster.forEach((m) => masterMap.set(cleanId(m.empId), m));

    const results = records.map((rec) => {
      const empId = cleanId(rec.empId);
      const profile = masterMap.get(empId) || {
        empId,
        name: rec.name || `EMPLOYEE #${empId}`,
        dutyHours: 9.5,
        salary: 15000,
        ownTiffin: "No",
        bonusDays: "Yes",
        active: "Yes",
        notes: ""
      };

      const isActive = String(profile.active || "").trim().toLowerCase() === "yes";
      const isDayWorker = Number(profile.dutyHours) <= 0;
      const dutyHours = Number(profile.dutyHours) || 0;
      const monthlySalary = Number(profile.salary) || 0;
      const ownTiffin = String(profile.ownTiffin || "").trim().toLowerCase() === "yes";
      const isBonusEligible = String(profile.bonusDays || "").trim().toLowerCase() === "yes";

      // If Inactive: Strictly zero out financials
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
          normalPresentDays: 0,
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
          bonusDays: 0,
          bonusDaySalary: 0,
          totalAdjustment: 0,
          netSalary: 0,
          remarks: "INACTIVE / EXCLUDED"
        };
      }

      // Rates (30-day divisor standard)
      const dailySalary = monthlySalary / 30;
      const hourlyRate = isDayWorker ? 0 : dailySalary / (dutyHours || 1);
      const effectiveDuty = isDayWorker ? 0 : ownTiffin ? Math.max(0, dutyHours - 0.5) : dutyHours;

      let normalPresentDays = 0;
      let absentDays = 0;
      let actualTotalHours = 0;
      let sundayWorkedDays = 0;
      let sundayExtraDays = 0;
      let sundayDeductionHours = 0;
      let hasSunday14HoursFlag = false;

      // Evaluate punches day-by-day
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
              sundayExtraDays += 1;
            }
          }
        } else {
          // Calendar Non-Sunday
          if (evaluatedStatus === "P") {
            normalPresentDays += 1.0;
          } else if (evaluatedStatus === "HD") {
            normalPresentDays += 0.5;
            absentDays += 0.5;
          } else if (evaluatedStatus === "A") {
            absentDays += 1.0;
          } else if (evaluatedStatus === "W") {
            // Scheduled weekly off
          } else {
            absentDays += 1.0;
          }
        }
      });

      // Normal Working Days requirement
      const normalWorkingDays = Math.max(0, meta.totalDays - meta.sundaysCount - absentDays);
      const requiredHours = isDayWorker ? 0 : normalWorkingDays * effectiveDuty;

      // Adjusted Punch Hours
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

      // CRITICAL BONUS DAYS EVALUATION
      let bonusDays = 0;
      if (isBonusEligible) {
        if (normalPresentDays >= 26) {
          bonusDays = 3;
        } else if (normalPresentDays === 25 || normalPresentDays === 24) {
          bonusDays = 2;
        } else if (normalPresentDays === 23) {
          bonusDays = 1;
        } else {
          bonusDays = 0;
        }
      }
      const bonusDaySalary = bonusDays * dailySalary;

      // Net Salary
      let netSalary = 0;
      if (isDayWorker) {
        netSalary = normalPresentDays * dailySalary + sundayExtraSalary + bonusDaySalary;
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

      // Adjustment
      const totalAdjustment =
        bonusDaySalary + sundayExtraSalary + extraHourSalary - shortHourDeduction;

      const remarksArr = [];
      if (hasSunday14HoursFlag) remarksArr.push("CHECK 14+ HRS SUNDAY");
      if (isDayWorker) remarksArr.push("Day Worker");
      if (shortHours > 12) remarksArr.push("High Short Hours");
      if (remarksArr.length === 0) remarksArr.push("Verified");

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
        normalPresentDays,
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
        bonusDays,
        bonusDaySalary,
        totalAdjustment,
        netSalary: Math.max(0, Math.round(netSalary)),
        remarks: remarksArr.join(" | ")
      };
    });

    // Safe natural numeric sorting
    results.sort((a, b) => {
      const numA = parseInt(a.empId, 10);
      const numB = parseInt(b.empId, 10);
      if (isNaN(numA) || isNaN(numB)) return a.empId.localeCompare(b.empId);
      return numA - numB;
    });

    return results;
  };

  // Re-run payroll if master records change
  useEffect(() => {
    if (attendanceRecords.length > 0) {
      const recalculated = calculatePayroll(attendanceRecords, master, metaInfo);
      setPayrollBreakdown(recalculated);
    }
  }, [master]);

  // =========================================================================
  // 3. EXACT ATTENDANCE PUNCH & DAILY PARSING LOGIC (ZERO COLUMN SHIFT)
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
        throw new Error("Uploaded workbook contains insufficient rows.");
      }

      // Step 1: Detect Date Range
      let fromDateStr = "";
      let toDateStr = "";
      let detectedTotalDays = 30;
      let fromDateObj = null;

      for (let r = 0; r < Math.min(15, rawMatrix.length); r++) {
        const row = rawMatrix[r];
        if (!row || !Array.isArray(row)) continue;
        for (let c = 0; c < row.length; c++) {
          const val = String(row[c] || "").trim();
          if (/from\s*:/i.test(val)) {
            const mFrom = val.match(/from\s*:\s*([^\s]+)/i);
            const mTo = val.match(/to\s*:\s*([^\s]+)/i);
            if (mFrom) fromDateStr = mFrom[1];
            if (mTo) toDateStr = mTo[1];

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

      if (fromDateStr) fromDateObj = parseReportDate(fromDateStr);
      if (!fromDateObj) {
        fromDateObj = new Date(2026, 3, 1);
        fromDateStr = "01/04/2026";
        toDateStr = "30/04/2026";
      }

      // Step 2: Detect sequence header 1..N
      let detectedDaysCount = 0;
      for (let r = 0; r < Math.min(30, rawMatrix.length); r++) {
        const row = rawMatrix[r];
        if (!row || !Array.isArray(row)) continue;
        for (let c = 0; c < row.length - 2; c++) {
          const n1 = parseInt(row[c], 10);
          const n2 = parseInt(row[c + 1], 10);
          const n3 = parseInt(row[c + 2], 10);
          if (n1 === 1 && n2 === 2 && n3 === 3) {
            let count = 0;
            while (parseInt(row[c + count], 10) === count + 1) count++;
            detectedDaysCount = count;
            break;
          }
        }
        if (detectedDaysCount > 0) break;
      }
      if (detectedDaysCount > 0) detectedTotalDays = detectedDaysCount;

      const sundayIndices = [];
      for (let d = 0; d < detectedTotalDays; d++) {
        const dObj = new Date(fromDateObj);
        dObj.setDate(fromDateObj.getDate() + d);
        if (dObj.getDay() === 0) sundayIndices.push(d);
      }

      // Step 3: Scan Employees & Sub-Rows
      // Physical Column 2 is DAY 1
      const parsedRecords = [];
      const masterById = new Map();
      master.forEach((m) => masterById.set(cleanId(m.empId), m));

      for (let r = 0; r < rawMatrix.length; r++) {
        const row = rawMatrix[r];
        if (!row || !Array.isArray(row)) continue;
        const rowStr = row.map((x) => String(x || "")).join(" ");

        if (/emp\s*id\s*:/i.test(rowStr)) {
          let rawEmpId = "";
          let rawEmpName = "";

          for (let c = 0; c < row.length; c++) {
            const cell = String(row[c] || "").trim();
            if (/emp\s*id\s*:/i.test(cell)) {
              const match = cell.match(/emp\s*id\s*:\s*([^\s]+)/i);
              if (match && match[1]) {
                rawEmpId = match[1];
              } else {
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

          const profile = masterById.get(empId);
          const finalName = profile ? profile.name : rawEmpName || `EMPLOYEE #${empId}`;

          let statusRow = null;
          let firstInRow = null;
          let lastOutRow = null;
          let totalHrsRow = null;

          for (let subR = r + 1; subR < Math.min(r + 18, rawMatrix.length); subR++) {
            const subRow = rawMatrix[subR];
            if (!subRow || !Array.isArray(subRow)) continue;
            const labelCell = String(subRow[0] || "").toLowerCase().trim();

            if (/emp\s*id\s*:/i.test(subRow.map((x) => String(x || "")).join(" "))) break;

            if (!statusRow && labelCell.includes("status")) {
              statusRow = subRow;
            } else if (!firstInRow && (labelCell.includes("first in") || labelCell.includes("in time"))) {
              firstInRow = subRow;
            } else if (!lastOutRow && (labelCell.includes("last out") || labelCell.includes("out time"))) {
              lastOutRow = subRow;
            } else if (!totalHrsRow && (labelCell.includes("total hrs") || labelCell.includes("work hrs") || labelCell.includes("total time"))) {
              totalHrsRow = subRow;
            }
          }

          const isDayWorker = profile ? Number(profile.dutyHours) <= 0 : false;
          const daysData = [];

          for (let d = 0; d < detectedTotalDays; d++) {
            const col = 2 + d;
            const rawStatus = statusRow ? String(statusRow[col] || "").trim() : "";
            const rawIn = firstInRow ? String(firstInRow[col] || "").trim() : "";
            const rawOut = lastOutRow ? String(lastOutRow[col] || "").trim() : "";
            const rawHrs = totalHrsRow ? String(totalHrsRow[col] || "").trim() : "";

            const isSunday = sundayIndices.includes(d);
            const inDec = parseTimeToHours(rawIn);
            const outDec = parseTimeToHours(rawOut);
            let totDec = parseDurationHours(rawHrs);

            if (totDec === 0 && inDec !== null && outDec !== null && outDec > inDec) {
              totDec = Number((outDec - inDec).toFixed(2));
            }

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

      parsedRecords.sort((a, b) => {
        const numA = parseInt(a.empId, 10);
        const numB = parseInt(b.empId, 10);
        if (isNaN(numA) || isNaN(numB)) return a.empId.localeCompare(b.empId);
        return numA - numB;
      });

      const newMeta = {
        fileName: file.name,
        fromDate: fromDateStr,
        toDate: toDateStr || `${detectedTotalDays}/Month`,
        totalDays: detectedTotalDays,
        sundaysCount: sundayIndices.length,
        processedAt: new Date().toLocaleTimeString()
      };

      setMetaInfo(newMeta);
      setAttendanceRecords(parsedRecords);

      const payroll = calculatePayroll(parsedRecords, master, newMeta);
      setPayrollBreakdown(payroll);

      logs.push(`Biometric sheet loaded with zero column shift. Extracted ${parsedRecords.length} staff records (${detectedTotalDays} days).`);
      setParseLogs(logs);
      setActiveTab("dashboard");
    } catch (err) {
      logs.push(`Parsing Diagnostic: ${err.message}`);
      setParseLogs(logs);
    } finally {
      setIsProcessing(false);
    }
  };

  // =========================================================================
  // 4. PRELOADED DEMO GENERATOR (DALPAT SINGH = 252.30 HRS & 26 PRESENT)
  // =========================================================================
  const loadInstantDemo = () => {
    setIsProcessing(true);
    const totalDays = 30;
    const sundays = [5, 12, 19, 26];

    const sampleAttendance = master.map((emp) => {
      const isDalpat = cleanId(emp.empId) === "1";
      const isDayWorker = Number(emp.dutyHours) <= 0;
      const days = [];

      for (let d = 1; d <= totalDays; d++) {
        const isSunday = sundays.includes(d);
        let status = "P";
        let inTime = "09:00";
        let outTime = "19:00";
        let hrs = isDayWorker ? 9.5 : emp.dutyHours;

        if (isDalpat) {
          if (isSunday) {
            status = "W";
            hrs = 0;
            inTime = "--";
            outTime = "--";
          } else {
            status = "P";
            hrs = d <= 8 ? 9.65 : d <= 18 ? 9.75 : 9.70;
            inTime = "09:12";
            outTime = "18:54";
          }
        } else {
          if (isSunday) {
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
            if (d === 8 && parseInt(emp.empId, 10) % 5 === 0) {
              status = "A";
              hrs = 0;
              inTime = "--";
              outTime = "--";
            } else if (d === 18 && parseInt(emp.empId, 10) % 7 === 0) {
              status = "HD";
              hrs = Math.max(0, hrs / 2);
              outTime = "14:00";
            }
          }
        }

        days.push({
          dayIndex: d,
          isSunday,
          rawStatus: status,
          firstIn: inTime,
          lastOut: outTime,
          firstInDec: inTime === "--" ? null : 9.2,
          lastOutDec: outTime === "--" ? null : 18.9,
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
      fileName: "BIOMETRIC_REPORT_APRIL_2026.xlsx",
      fromDate: "01/04/2026",
      toDate: "30/04/2026",
      totalDays: 30,
      sundaysCount: 4,
      processedAt: new Date().toLocaleTimeString()
    };

    setAttendanceRecords(sampleAttendance);
    setMetaInfo(meta);
    const payroll = calculatePayroll(sampleAttendance, master, meta);
    setPayrollBreakdown(payroll);
    setParseLogs(["Industrial reference dataset loaded. Dalpat Singh: 252.30 Total Hrs, 26 Present Days, 3 Bonus Days."]);
    setIsProcessing(false);
  };

  useEffect(() => {
    if (attendanceRecords.length === 0) {
      loadInstantDemo();
    }
  }, []);

  // =========================================================================
  // 5. BULK MASTER UPLOAD & EXCEL EXPORT
  // =========================================================================
  const handleBulkMasterUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws);

        if (!json.length) {
          alert("The uploaded master sheet contains no data.");
          return;
        }

        const cleanedRows = json.map((r) => {
          const rawId = cleanId(r["Emp ID"] || r.empId || r.ID || r["Employee ID"]);
          const rawName = String(r["Employee Name"] || r.name || r.Name || "").trim();
          const rawDuty = Number(r["Duty Hours"] ?? r.dutyHours ?? 9.5);
          const rawSal = Number(String(r["Monthly Salary"] ?? r.salary ?? 15000).replace(/[₹,]/g, ""));
          const rawTiffin = String(r["Own Tiffin"] ?? r.ownTiffin ?? "No").trim();
          const rawBonus = String(r["Bonus Days"] ?? r.bonusDays ?? "Yes").trim();
          const rawActive = String(r["Active"] ?? r.active ?? "Yes").trim();
          const rawNotes = String(r["Notes"] ?? r.notes ?? "").trim();

          return {
            empId: rawId,
            name: rawName.toUpperCase(),
            dutyHours: isNaN(rawDuty) ? 9.5 : rawDuty,
            salary: isNaN(rawSal) ? 0 : rawSal,
            ownTiffin: rawTiffin.toLowerCase() === "yes" ? "Yes" : "No",
            bonusDays: rawBonus.toLowerCase() === "no" ? "No" : "Yes",
            active: rawActive.toLowerCase() === "no" ? "No" : "Yes",
            notes: rawNotes
          };
        }).filter((x) => x.empId && x.name);

        if (!cleanedRows.length) {
          alert("No valid employee entries found. Verify column headers.");
          return;
        }

        setMaster((prev) => {
          const map = new Map(prev.map((item) => [cleanId(item.empId), item]));
          cleanedRows.forEach((row) => map.set(cleanId(row.empId), row));
          const merged = Array.from(map.values()).sort((a, b) => {
            const numA = parseInt(a.empId, 10);
            const numB = parseInt(b.empId, 10);
            if (isNaN(numA) || isNaN(numB)) return a.empId.localeCompare(b.empId);
            return numA - numB;
          });
          localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(merged));
          return merged;
        });

        alert(`Successfully imported / merged ${cleanedRows.length} employees into Master!`);
      } catch (err) {
        alert("Failed to parse Master file: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const handleDownloadMasterTemplate = () => {
    const templateData = master.slice(0, 10).map((m) => ({
      "Emp ID": m.empId,
      "Employee Name": m.name,
      "Duty Hours": m.dutyHours,
      "Monthly Salary": m.salary,
      "Own Tiffin": m.ownTiffin,
      "Bonus Days": m.bonusDays,
      "Active": m.active,
      "Notes": m.notes || "Standard"
    }));

    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Master_Template");
    XLSX.writeFile(wb, "Employee_Master_Template.xlsx");
  };

  const handleResetMaster = () => {
    if (window.confirm("Restore factory default roster (75+ employees)? Any unsaved additions will be overwritten.")) {
      setMaster(INITIAL_MASTER);
      localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(INITIAL_MASTER));
    }
  };

  const handleSaveSingleEmp = (emp) => {
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
      if (master.some((item) => cleanId(item.empId) === targetId)) {
        alert(`Employee with ID "${targetId}" already exists!`);
        return;
      }
      setMaster((prev) => [...prev, emp]);
    }
    setIsEmpModalOpen(false);
    setEditingEmp(null);
  };

  const handleDeleteEmp = (id) => {
    if (window.confirm(`Delete Employee ID: ${id} from Master?`)) {
      setMaster((prev) => prev.filter((m) => cleanId(m.empId) !== cleanId(id)));
    }
  };

  const exportPayrollToExcel = () => {
    if (!payrollBreakdown.length) return;

    const exportRows = payrollBreakdown
      .filter((row) => (showInactive ? true : row.isActive))
      .map((r) => ({
        "Emp ID": r.empId,
        "Employee Name": r.name,
        "Status": r.isActive ? "ACTIVE" : "INACTIVE",
        "Category": r.isDayWorker ? "Day Worker" : "Duty Worker",
        "Duty Hrs": r.dutyHours,
        "Effective Duty": r.effectiveDuty,
        "Monthly Salary (₹)": r.monthlySalary,
        "Daily Rate (₹)": Math.round(r.dailySalary),
        "Hourly Rate (₹)": Number(r.hourlyRate.toFixed(2)),
        "Normal Present Days": r.normalPresentDays,
        "Absent Days": r.absentDays,
        "Sunday OT Days": r.sundayExtraDays,
        "Punch Actual Hrs": Number(r.actualTotalHours.toFixed(2)),
        "Required Hrs": Number(r.requiredHours.toFixed(2)),
        "Short Hrs": r.shortHours,
        "Extra Hrs": r.extraHours,
        "Bonus Days": r.bonusDays,
        "Absent Ded (₹)": Math.round(r.absentDeduction),
        "Short Hrs Ded (₹)": Math.round(r.shortHourDeduction),
        "Extra Hrs Pay (₹)": Math.round(r.extraHourSalary),
        "Sunday Pay (₹)": Math.round(r.sundayExtraSalary),
        "Bonus Pay (₹)": Math.round(r.bonusDaySalary),
        "Adjustment (AD) (₹)": Math.round(r.totalAdjustment),
        "NET PAYABLE (₹)": r.netSalary,
        "Remarks": r.remarks
      }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Payroll_Report");
    XLSX.writeFile(wb, `Payroll_Summary_${metaInfo.totalDays}Days_${metaInfo.fileName.replace(/\.xlsx?/i, "")}.xlsx`);
  };

  // =========================================================================
  // 6. COMPUTED AGGREGATIONS (KPIs STRICTLY RESPECT INACTIVE EXCLUSION)
  // =========================================================================
  const kpiStats = useMemo(() => {
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

  // Filters for tables
  const displayedPayroll = useMemo(() => {
    return payrollBreakdown.filter((r) => {
      if (!showInactive && !r.isActive) return false;
      if (!searchQuery) return true;
      const q = searchQuery.toLowerCase();
      return r.empId.toLowerCase().includes(q) || r.name.toLowerCase().includes(q) || r.remarks.toLowerCase().includes(q);
    });
  }, [payrollBreakdown, showInactive, searchQuery]);

  const displayedMaster = useMemo(() => {
    return [...master]
      .filter((m) => {
        if (!showInactive && String(m.active).toLowerCase() === "no") return false;
        if (!searchQuery) return true;
        const q = searchQuery.toLowerCase();
        return m.empId.toLowerCase().includes(q) || m.name.toLowerCase().includes(q);
      })
      .sort((a, b) => {
        const numA = parseInt(a.empId, 10);
        const numB = parseInt(b.empId, 10);
        if (isNaN(numA) || isNaN(numB)) return a.empId.localeCompare(b.empId);
        return numA - numB;
      });
  }, [master, showInactive, searchQuery]);

  const flattenedAttendance = useMemo(() => {
    const list = [];
    attendanceRecords.forEach((emp) => {
      const profile = master.find((m) => cleanId(m.empId) === cleanId(emp.empId));
      const isActive = profile ? String(profile.active).toLowerCase() === "yes" : true;
      if (!showInactive && !isActive) return;

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
  }, [attendanceRecords, master, showInactive]);

  const displayedAttendance = useMemo(() => {
    return flattenedAttendance.filter((rec) => {
      const matchSearch =
        !searchQuery ||
        rec.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
        rec.empId.toLowerCase().includes(searchQuery.toLowerCase()) ||
        `Day ${rec.dayIndex}`.toLowerCase().includes(searchQuery.toLowerCase());

      const matchStatus = attendanceFilter === "ALL" ? true : rec.evaluatedStatus === attendanceFilter;
      return matchSearch && matchStatus;
    });
  }, [flattenedAttendance, searchQuery, attendanceFilter]);

  return (
    <div className="min-h-screen bg-[#0B0F19] text-slate-100 font-sans selection:bg-indigo-500 selection:text-white flex flex-col relative overflow-x-hidden">
      {/* Background Ambient Orbs */}
      <div className="fixed top-0 left-1/4 w-[650px] h-[400px] bg-indigo-600/10 rounded-full blur-[140px] pointer-events-none -z-10" />
      <div className="fixed bottom-10 right-10 w-[550px] h-[450px] bg-emerald-600/10 rounded-full blur-[150px] pointer-events-none -z-10" />
      <div className="fixed top-1/2 right-1/4 w-[450px] h-[350px] bg-amber-600/10 rounded-full blur-[130px] pointer-events-none -z-10" />

      {/* Top Header */}
      <header className="sticky top-0 z-40 bg-[#0B0F19]/85 backdrop-blur-xl border-b border-slate-800/80 px-4 sm:px-8 py-3.5 flex flex-wrap items-center justify-between gap-4 shadow-2xl">
        <div className="flex items-center space-x-3">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-indigo-600 via-violet-600 to-cyan-400 p-[1px] shadow-lg shadow-indigo-500/20">
            <div className="w-full h-full bg-[#0B0F19] rounded-[11px] flex items-center justify-center">
              <Building2 className="w-5 h-5 text-indigo-400" />
            </div>
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="font-extrabold tracking-tight text-white text-base sm:text-lg">
                PUNCH<span className="text-transparent bg-clip-text bg-gradient-to-r from-indigo-400 to-cyan-400">PAY</span> INDUSTRIAL
              </h1>
              <span className="text-[10px] font-mono uppercase bg-indigo-500/10 text-indigo-300 border border-indigo-500/30 px-2 py-0.5 rounded-full font-bold flex items-center gap-1">
                <ShieldCheck className="w-3 h-3 text-emerald-400" /> Biometric Engine
              </span>
            </div>
            <p className="text-[11px] text-slate-400 flex items-center gap-1.5">
              <span>{metaInfo.fileName}</span>
              <span className="w-1 h-1 rounded-full bg-slate-600" />
              <span>{metaInfo.fromDate} - {metaInfo.toDate}</span>
              <span className="w-1 h-1 rounded-full bg-slate-600" />
              <span className="text-emerald-400 font-medium">{metaInfo.totalDays} Days ({metaInfo.sundaysCount} Sundays)</span>
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          <button
            onClick={loadInstantDemo}
            title="Reload Demo Data"
            className="hidden sm:inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-800/80 hover:bg-slate-700/80 text-xs text-slate-300 hover:text-white border border-slate-700/70 transition shadow-sm"
          >
            <Sparkles className="w-3.5 h-3.5 text-amber-400" /> Load Demo
          </button>

          <button
            onClick={exportPayrollToExcel}
            className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-medium text-xs shadow-lg shadow-emerald-700/20 transition-all hover:scale-[1.02]"
          >
            <Download className="w-3.5 h-3.5" /> Export Excel
          </button>
        </div>
      </header>

      {/* Secondary Navigation */}
      <div className="border-b border-slate-800/80 bg-[#0B0F19]/40 backdrop-blur-md px-4 sm:px-8 py-2.5 flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-1.5 p-1 rounded-xl bg-slate-900/80 border border-slate-800">
          {[
            { id: "dashboard", label: "Dashboard", icon: TrendingUp },
            { id: "salary", label: "Salary Calculation Sheet", icon: DollarSign },
            { id: "attendance", label: "Daily Punch Inspector", icon: Calendar },
            { id: "master", label: `Master Roster (${master.length})`, icon: Users },
            { id: "upload", label: "Biometric Uploader", icon: Upload }
          ].map((tab) => {
            const Icon = tab.icon;
            const active = activeTab === tab.id;
            return (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-xs font-semibold transition-all duration-200 ${
                  active
                    ? "bg-gradient-to-r from-indigo-600 to-violet-600 text-white shadow-md shadow-indigo-600/30"
                    : "text-slate-400 hover:text-slate-200 hover:bg-slate-800/50"
                }`}
              >
                <Icon className={`w-3.5 h-3.5 ${active ? "text-white" : "text-slate-400"}`} />
                <span>{tab.label}</span>
              </button>
            );
          })}
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={() => setShowInactive(!showInactive)}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border transition ${
              showInactive
                ? "bg-amber-500/10 text-amber-300 border-amber-500/30"
                : "bg-slate-800/80 text-slate-400 border-slate-700/80 hover:text-slate-200"
            }`}
          >
            {showInactive ? <Eye className="w-3.5 h-3.5 text-amber-400" /> : <EyeOff className="w-3.5 h-3.5 text-slate-400" />}
            <span>Show Inactive ({master.filter((m) => String(m.active).toLowerCase() === "no").length})</span>
          </button>

          <div className="relative">
            <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-500" />
            <input
              type="text"
              placeholder="Search ID, Name..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="pl-8 pr-3 py-1.5 text-xs bg-slate-900/90 border border-slate-700/70 rounded-lg text-slate-200 placeholder-slate-500 focus:outline-none focus:ring-1 focus:ring-indigo-500 w-44 sm:w-56"
            />
            {searchQuery && (
              <button onClick={() => setSearchQuery("")} className="absolute right-2.5 top-2 text-slate-400 hover:text-slate-200">
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8 space-y-6">
        {parseLogs.length > 0 && (
          <div className="p-3 bg-indigo-950/40 border border-indigo-800/60 rounded-xl text-xs text-indigo-200 flex items-center justify-between shadow-lg">
            <div className="flex items-center gap-2">
              <Activity className="w-4 h-4 text-indigo-400 flex-shrink-0 animate-pulse" />
              <span>{parseLogs[parseLogs.length - 1]}</span>
            </div>
            <button onClick={() => setParseLogs([])} className="text-slate-400 hover:text-white text-[11px] underline ml-4">
              Dismiss
            </button>
          </div>
        )}

        {/* TAB 1: DASHBOARD */}
        {activeTab === "dashboard" && (
          <div className="space-y-6">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <div className="p-4 rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 hover:border-emerald-500/50 transition-all duration-300 hover:-translate-y-1 hover:shadow-xl hover:shadow-emerald-500/10 group">
                <div className="flex items-center justify-between text-xs font-semibold text-slate-400">
                  <span>Net Payable Payroll</span>
                  <div className="p-2 rounded-xl bg-emerald-500/10 text-emerald-400 group-hover:scale-110 transition">
                    <DollarSign className="w-4 h-4" />
                  </div>
                </div>
                <div className="mt-2 text-2xl sm:text-3xl font-extrabold text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 to-teal-300 font-mono">
                  {formatINR(kpiStats.totalNet)}
                </div>
                <div className="mt-1 text-[11px] text-slate-400 flex items-center justify-between">
                  <span>Base: {formatINR(kpiStats.totalBase)}</span>
                  <span className="text-emerald-400 font-medium">Active Only</span>
                </div>
              </div>

              <div className="p-4 rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 hover:border-blue-500/50 transition-all duration-300 hover:-translate-y-1 hover:shadow-xl hover:shadow-blue-500/10 group">
                <div className="flex items-center justify-between text-xs font-semibold text-slate-400">
                  <span>Net Total Adjustment (AD)</span>
                  <div className="p-2 rounded-xl bg-blue-500/10 text-blue-400 group-hover:scale-110 transition">
                    <Award className="w-4 h-4" />
                  </div>
                </div>
                <div className={`mt-2 text-2xl sm:text-3xl font-extrabold font-mono ${kpiStats.totalAdjustment >= 0 ? "text-blue-400" : "text-rose-400"}`}>
                  {kpiStats.totalAdjustment >= 0 ? "+" : ""}
                  {formatINR(kpiStats.totalAdjustment)}
                </div>
                <div className="mt-1 text-[11px] text-slate-400">Bonus + OT - Short Ded</div>
              </div>

              <div className="p-4 rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 hover:border-amber-500/50 transition-all duration-300 hover:-translate-y-1 hover:shadow-xl hover:shadow-amber-500/10 group">
                <div className="flex items-center justify-between text-xs font-semibold text-slate-400">
                  <span>Total Bonus Paid</span>
                  <div className="p-2 rounded-xl bg-amber-500/10 text-amber-400 group-hover:scale-110 transition">
                    <Award className="w-4 h-4" />
                  </div>
                </div>
                <div className="mt-2 text-2xl sm:text-3xl font-extrabold text-amber-400 font-mono">
                  {formatINR(kpiStats.totalBonusPay)}
                </div>
                <div className="mt-1 text-[11px] text-slate-400">23+ Present Days Tier</div>
              </div>

              <div className="p-4 rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 hover:border-indigo-500/50 transition-all duration-300 hover:-translate-y-1 hover:shadow-xl hover:shadow-indigo-500/10 group">
                <div className="flex items-center justify-between text-xs font-semibold text-slate-400">
                  <span>Active Headcount</span>
                  <div className="p-2 rounded-xl bg-indigo-500/10 text-indigo-400 group-hover:scale-110 transition">
                    <Users className="w-4 h-4" />
                  </div>
                </div>
                <div className="mt-2 text-2xl sm:text-3xl font-extrabold text-white font-mono">
                  {kpiStats.totalActiveEmp}{" "}
                  <span className="text-xs text-slate-400 font-sans font-normal">/ {master.length} in Roster</span>
                </div>
                <div className="mt-1 text-[11px] text-slate-400">
                  {kpiStats.totalMonthlyWorkers} Duty | {kpiStats.totalDayWorkers} Day Workers
                </div>
              </div>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <div className="p-3.5 rounded-xl bg-slate-900/50 border border-slate-800/70 flex items-center justify-between">
                <div>
                  <div className="text-[11px] text-slate-400">Sunday Overtime Days</div>
                  <div className="text-lg font-bold text-amber-400 font-mono">{kpiStats.totalSundayOTDays} Shifts</div>
                </div>
                <Sun className="w-4 h-4 text-amber-400" />
              </div>

              <div className="p-3.5 rounded-xl bg-slate-900/50 border border-slate-800/70 flex items-center justify-between">
                <div>
                  <div className="text-[11px] text-slate-400">Cumulative Absent Days</div>
                  <div className="text-lg font-bold text-rose-400 font-mono">{formatDec(kpiStats.totalAbsent, 1)} Days</div>
                </div>
                <AlertTriangle className="w-4 h-4 text-rose-400" />
              </div>

              <div className="p-3.5 rounded-xl bg-slate-900/50 border border-slate-800/70 flex items-center justify-between">
                <div>
                  <div className="text-[11px] text-slate-400">Total Short Hours</div>
                  <div className="text-lg font-bold text-amber-400 font-mono">{formatDec(kpiStats.totalShortHours, 1)} Hrs</div>
                </div>
                <Clock className="w-4 h-4 text-amber-400" />
              </div>

              <div className="p-3.5 rounded-xl bg-slate-900/50 border border-slate-800/70 flex items-center justify-between">
                <div>
                  <div className="text-[11px] text-slate-400">Total Extra Hours</div>
                  <div className="text-lg font-bold text-indigo-400 font-mono">+{formatDec(kpiStats.totalExtraHours, 1)} Hrs</div>
                </div>
                <TrendingUp className="w-4 h-4 text-indigo-400" />
              </div>
            </div>

            {/* Preview Table */}
            <div className="rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 overflow-hidden shadow-2xl">
              <div className="p-4 sm:p-5 border-b border-slate-800 flex flex-wrap justify-between items-center gap-3">
                <div>
                  <h3 className="font-bold text-sm sm:text-base text-white flex items-center gap-2">
                    <Layers className="w-4 h-4 text-indigo-400" /> Monthly Payroll Summary Preview
                  </h3>
                  <p className="text-xs text-slate-400">
                    Live evaluated employee payroll records. Dalpat Singh: 252.30 Hrs, 26 Present, Bonus Pay evaluated.
                  </p>
                </div>
                <button
                  onClick={() => setActiveTab("salary")}
                  className="text-xs text-indigo-400 hover:text-indigo-300 font-semibold flex items-center gap-1 group"
                >
                  <span>Open Full Calculation Sheet</span>
                  <ChevronRight className="w-4 h-4 group-hover:translate-x-0.5 transition" />
                </button>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs border-collapse font-mono">
                  <thead className="bg-slate-950/80 text-slate-400 uppercase text-[10px] tracking-wider border-b border-slate-800 font-semibold font-sans">
                    <tr>
                      <th className="py-3 px-4">Emp ID</th>
                      <th className="py-3 px-4">Employee Name</th>
                      <th className="py-3 px-3">Type</th>
                      <th className="py-3 px-3 text-right">Base Salary</th>
                      <th className="py-3 px-3 text-center">P / A Days</th>
                      <th className="py-3 px-3 text-right">Punch Hrs</th>
                      <th className="py-3 px-3 text-center text-amber-400">Bonus Pay</th>
                      <th className="py-3 px-3 text-right">Adjustment (AD)</th>
                      <th className="py-3 px-4 text-right text-emerald-400 font-bold">NET SALARY</th>
                      <th className="py-3 px-4 font-sans">Remarks</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/60">
                    {displayedPayroll.slice(0, 10).map((row) => (
                      <tr
                        key={row.empId}
                        className={`hover:bg-slate-800/30 transition-colors ${!row.isActive ? "opacity-40 bg-slate-950/40" : ""}`}
                      >
                        <td className="py-2.5 px-4 font-bold text-indigo-400">{row.empId}</td>
                        <td className="py-2.5 px-4 font-sans font-semibold text-slate-200">
                          {row.name}
                          {!row.isActive && (
                            <span className="ml-2 text-[9px] bg-rose-500/20 text-rose-300 px-1.5 py-0.5 rounded border border-rose-500/30">
                              INACTIVE
                            </span>
                          )}
                        </td>
                        <td className="py-2.5 px-3 font-sans">
                          <span
                            className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                              row.isDayWorker
                                ? "bg-amber-500/10 text-amber-300 border border-amber-500/30"
                                : "bg-indigo-500/10 text-indigo-300 border border-indigo-500/30"
                            }`}
                          >
                            {row.isDayWorker ? "Day Worker" : `${row.dutyHours}h`}
                          </span>
                        </td>
                        <td className="py-2.5 px-3 text-right text-slate-300">{formatINR(row.monthlySalary)}</td>
                        <td className="py-2.5 px-3 text-center font-sans">
                          <span className="text-emerald-400 font-bold">{row.normalPresentDays}</span> /{" "}
                          <span className="text-rose-400 font-bold">{row.absentDays}</span>
                        </td>
                        <td className="py-2.5 px-3 text-right font-bold text-slate-200">
                          {formatDec(row.actualTotalHours, 2)}h
                        </td>
                        <td className="py-2.5 px-3 text-center">
                          {row.bonusDaySalary > 0 ? (
                            <span className="text-amber-400 font-bold">
                              +{formatINR(row.bonusDaySalary)} ({row.bonusDays}d)
                            </span>
                          ) : (
                            <span className="text-slate-500 font-mono">₹0</span>
                          )}
                        </td>
                        <td className={`py-2.5 px-3 text-right font-bold ${row.totalAdjustment >= 0 ? "text-blue-400" : "text-rose-400"}`}>
                          {row.totalAdjustment >= 0 ? "+" : ""}
                          {formatINR(row.totalAdjustment)}
                        </td>
                        <td className="py-2.5 px-4 text-right font-extrabold text-emerald-400 text-sm">
                          {formatINR(row.netSalary)}
                        </td>
                        <td className="py-2.5 px-4 font-sans text-[11px] text-slate-400">{row.remarks}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: DETAILED SALARY SHEET */}
        {activeTab === "salary" && (
          <div className="space-y-4">
            <div className="rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 p-4 sm:p-5 flex flex-wrap justify-between items-center gap-3">
              <div>
                <h2 className="font-bold text-base text-white flex items-center gap-2">
                  <DollarSign className="w-5 h-5 text-emerald-400" /> Full Industrial Salary Calculation Sheet
                </h2>
                <p className="text-xs text-slate-400">
                  Daily punch hours parsed from biometric Sub-Rows (`2 + dayIndex`). Normal present days are maintained without deduction for OT or short hours.
                </p>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={exportPayrollToExcel}
                  className="px-3.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white font-semibold text-xs flex items-center gap-1.5 shadow-md shadow-emerald-700/20"
                >
                  <Download className="w-3.5 h-3.5" /> Export Excel Sheet
                </button>
              </div>
            </div>

            <div className="rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 overflow-hidden shadow-2xl">
              <div className="overflow-x-auto max-h-[720px]">
                <table className="w-full text-left text-xs border-collapse font-mono">
                  <thead className="sticky top-0 z-20 bg-[#0A0D14] text-slate-300 uppercase text-[10px] tracking-wider border-b border-slate-800 font-bold font-sans">
                    <tr>
                      <th className="py-3 px-3">Emp ID</th>
                      <th className="py-3 px-3">Name</th>
                      <th className="py-3 px-2">Type</th>
                      <th className="py-3 px-2 text-right">Base (₹)</th>
                      <th className="py-3 px-2 text-right">Daily/Hr</th>
                      <th className="py-3 px-2 text-center text-emerald-400">P Days</th>
                      <th className="py-3 px-2 text-center text-rose-400">A Days</th>
                      <th className="py-3 px-2 text-center text-amber-400">Sun OT</th>
                      <th className="py-3 px-2 text-right">Punch H</th>
                      <th className="py-3 px-2 text-right">Req H</th>
                      <th className="py-3 px-2 text-right text-rose-400">Short H</th>
                      <th className="py-3 px-2 text-right text-indigo-400">Extra H</th>
                      <th className="py-3 px-2 text-right text-rose-400">Absent Ded</th>
                      <th className="py-3 px-2 text-right text-rose-400">Short Ded</th>
                      <th className="py-3 px-2 text-right text-indigo-400">Extra Pay</th>
                      <th className="py-3 px-2 text-right text-amber-400">Sun Pay</th>
                      <th className="py-3 px-2 text-center text-amber-300">Bonus Pay</th>
                      <th className="py-3 px-3 text-right bg-slate-950 text-emerald-400 font-extrabold text-sm">
                        NET SALARY
                      </th>
                      <th className="py-3 px-3 font-sans">Remarks</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50">
                    {displayedPayroll.length === 0 ? (
                      <tr>
                        <td colSpan="19" className="text-center py-12 text-slate-500 font-sans">
                          No matching payroll records found.
                        </td>
                      </tr>
                    ) : (
                      displayedPayroll.map((emp) => (
                        <tr
                          key={emp.empId}
                          className={`hover:bg-slate-800/30 transition-colors ${!emp.isActive ? "opacity-35 bg-slate-950/60" : ""}`}
                        >
                          <td className="py-2 px-3 font-bold text-indigo-400">{emp.empId}</td>
                          <td className="py-2 px-3 font-sans font-semibold text-slate-200">
                            {emp.name}
                            {!emp.isActive && <span className="block text-[9px] text-rose-400 font-mono">INACTIVE (EXCLUDED)</span>}
                          </td>
                          <td className="py-2 px-2 font-sans">
                            <span
                              className={`px-1.5 py-0.5 rounded text-[9px] font-bold ${
                                emp.isDayWorker
                                  ? "bg-amber-500/10 text-amber-300 border border-amber-500/30"
                                  : "bg-indigo-500/10 text-indigo-300 border border-indigo-500/30"
                              }`}
                            >
                              {emp.isDayWorker ? "Day" : `${emp.dutyHours}h`}
                            </span>
                          </td>
                          <td className="py-2 px-2 text-right text-slate-300">{formatINR(emp.monthlySalary)}</td>
                          <td className="py-2 px-2 text-right text-[11px] text-slate-400">
                            {emp.isDayWorker ? `${Math.round(emp.dailySalary)}/d` : `${Math.round(emp.hourlyRate)}/h`}
                          </td>
                          <td className="py-2 px-2 text-center text-emerald-400 font-bold">{emp.normalPresentDays}</td>
                          <td className="py-2 px-2 text-center text-rose-400 font-bold">{emp.absentDays}</td>
                          <td className="py-2 px-2 text-center text-amber-400 font-bold">
                            {emp.sundayExtraDays > 0 ? `+${emp.sundayExtraDays}` : "0"}
                          </td>
                          <td className="py-2 px-2 text-right text-slate-300">{formatDec(emp.actualTotalHours, 1)}</td>
                          <td className="py-2 px-2 text-right text-slate-500">{formatDec(emp.requiredHours, 1)}</td>
                          <td className="py-2 px-2 text-right text-rose-400">{emp.shortHours > 0 ? formatDec(emp.shortHours, 1) : "-"}</td>
                          <td className="py-2 px-2 text-right text-indigo-400">{emp.extraHours > 0 ? formatDec(emp.extraHours, 1) : "-"}</td>
                          <td className="py-2 px-2 text-right text-rose-400">
                            {emp.absentDeduction > 0 ? `-${formatINR(emp.absentDeduction)}` : "-"}
                          </td>
                          <td className="py-2 px-2 text-right text-rose-400">
                            {emp.shortHourDeduction > 0 ? `-${formatINR(emp.shortHourDeduction)}` : "-"}
                          </td>
                          <td className="py-2 px-2 text-right text-indigo-400">
                            {emp.extraHourSalary > 0 ? `+${formatINR(emp.extraHourSalary)}` : "-"}
                          </td>
                          <td className="py-2 px-2 text-right text-amber-400">
                            {emp.sundayExtraSalary > 0 ? `+${formatINR(emp.sundayExtraSalary)}` : "-"}
                          </td>
                          <td className="py-2 px-2 text-center">
                            {emp.bonusDaySalary > 0 ? (
                              <span className="text-amber-400 font-bold">
                                +{formatINR(emp.bonusDaySalary)} ({emp.bonusDays}d)
                              </span>
                            ) : (
                              <span className="text-slate-500 font-mono">₹0</span>
                            )}
                          </td>
                          <td className="py-2 px-3 text-right bg-slate-950/80 font-extrabold text-emerald-400 text-sm">
                            {formatINR(emp.netSalary)}
                          </td>
                          <td className="py-2 px-3 font-sans text-[10px] text-slate-400">{emp.remarks}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* TAB 3: DAILY PUNCH INSPECTOR */}
        {activeTab === "attendance" && (
          <div className="space-y-4">
            <div className="rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 p-4 sm:p-5 flex flex-wrap justify-between items-center gap-3">
              <div>
                <h2 className="font-bold text-base text-white flex items-center gap-2">
                  <Calendar className="w-5 h-5 text-indigo-400" /> Daily Punch & Biometric Timeline
                </h2>
                <p className="text-xs text-slate-400">
                  Physical column mapping: Day 1 punch is extracted at column index 2.
                </p>
              </div>

              <div className="flex items-center gap-1.5 bg-slate-950/70 p-1 rounded-xl border border-slate-800">
                {["ALL", "P", "PW", "HD", "A", "W"].map((st) => (
                  <button
                    key={st}
                    onClick={() => setAttendanceFilter(st)}
                    className={`px-2.5 py-1 rounded-lg text-[11px] font-bold transition ${
                      attendanceFilter === st
                        ? "bg-indigo-600 text-white shadow-md shadow-indigo-600/30"
                        : "text-slate-400 hover:text-white"
                    }`}
                  >
                    {st === "ALL" ? "All Punches" : st}
                  </button>
                ))}
              </div>
            </div>

            <div className="rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 overflow-hidden shadow-2xl">
              <div className="overflow-x-auto max-h-[650px]">
                <table className="w-full text-left text-xs border-collapse font-mono">
                  <thead className="sticky top-0 z-20 bg-[#0A0D14] text-slate-400 uppercase text-[10px] tracking-wider border-b border-slate-800 font-sans font-bold">
                    <tr>
                      <th className="py-3 px-4">Day</th>
                      <th className="py-3 px-3">Emp ID</th>
                      <th className="py-3 px-4">Employee Name</th>
                      <th className="py-3 px-3 text-center">Status</th>
                      <th className="py-3 px-3 text-center">First In</th>
                      <th className="py-3 px-3 text-center">Last Out</th>
                      <th className="py-3 px-3 text-right">Hours</th>
                      <th className="py-3 px-4 text-center">Raw Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/40">
                    {displayedAttendance.length === 0 ? (
                      <tr>
                        <td colSpan="8" className="text-center py-12 text-slate-500 font-sans">
                          No matching biometric punches found.
                        </td>
                      </tr>
                    ) : (
                      displayedAttendance.slice(0, 300).map((row, idx) => (
                        <tr
                          key={idx}
                          className={`hover:bg-slate-800/30 transition-colors ${row.isSunday ? "bg-amber-950/20" : ""}`}
                        >
                          <td className="py-2 px-4 font-semibold text-slate-300 font-sans">
                            Day {row.dayIndex}{" "}
                            {row.isSunday && (
                              <span className="text-amber-400 font-bold ml-1 text-[10px] bg-amber-500/10 px-1 py-0.5 rounded border border-amber-500/30">
                                SUN
                              </span>
                            )}
                          </td>
                          <td className="py-2 px-3 font-bold text-indigo-400">{row.empId}</td>
                          <td className="py-2 px-4 font-sans font-semibold text-slate-200">{row.name}</td>
                          <td className="py-2 px-3 text-center font-sans">
                            {row.evaluatedStatus === "P" && (
                              <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-500/10 text-emerald-300 border border-emerald-500/30">
                                P
                              </span>
                            )}
                            {row.evaluatedStatus === "PW" && (
                              <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-purple-500/20 text-purple-300 border border-purple-500/40">
                                PW (Sun OT)
                              </span>
                            )}
                            {row.evaluatedStatus === "HD" && (
                              <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-500/10 text-amber-300 border border-amber-500/30">
                                HD
                              </span>
                            )}
                            {row.evaluatedStatus === "A" && (
                              <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-rose-500/10 text-rose-300 border border-rose-500/30">
                                A
                              </span>
                            )}
                            {row.evaluatedStatus === "W" && (
                              <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-800 text-slate-400 border border-slate-700">
                                W
                              </span>
                            )}
                          </td>
                          <td className="py-2 px-3 text-center text-slate-300">{row.firstIn}</td>
                          <td className="py-2 px-3 text-center text-slate-300">{row.lastOut}</td>
                          <td className="py-2 px-3 text-right font-bold text-slate-200">{formatDec(row.actualHours, 2)}h</td>
                          <td className="py-2 px-4 text-center text-slate-500 text-[11px]">{row.rawStatus || "-"}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* TAB 4: EMPLOYEE MASTER */}
        {activeTab === "master" && (
          <div className="space-y-4">
            <div className="rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 p-4 sm:p-5 flex flex-wrap justify-between items-center gap-3">
              <div>
                <h2 className="font-bold text-base text-white flex items-center gap-2">
                  <Users className="w-5 h-5 text-indigo-400" /> Employee Master Roster ({master.length} Total Records)
                </h2>
                <p className="text-xs text-slate-400">
                  Dual-Input: Bulk import spreadsheets or add individual employees. Automatically persists in LocalStorage.
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <input
                  type="file"
                  ref={fileMasterInputRef}
                  onChange={handleBulkMasterUpload}
                  accept=".xlsx, .xls, .csv"
                  className="hidden"
                />
                <button
                  onClick={() => fileMasterInputRef.current?.click()}
                  className="px-3 py-1.5 rounded-lg bg-indigo-600/20 hover:bg-indigo-600/30 text-indigo-300 border border-indigo-500/30 text-xs font-semibold flex items-center gap-1.5 transition"
                >
                  <FileUp className="w-3.5 h-3.5" /> Bulk Upload Master
                </button>

                <button
                  onClick={handleDownloadMasterTemplate}
                  className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-medium border border-slate-700 transition flex items-center gap-1.5"
                >
                  <Download className="w-3.5 h-3.5" /> Template
                </button>

                <button
                  onClick={handleResetMaster}
                  className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white text-xs font-medium border border-slate-700 transition flex items-center gap-1.5"
                >
                  <RotateCcw className="w-3.5 h-3.5 text-amber-400" /> Reset Factory Roster
                </button>

                <button
                  onClick={() => {
                    setEditingEmp(null);
                    setIsEmpModalOpen(true);
                  }}
                  className="px-3.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold shadow-lg shadow-indigo-600/30 flex items-center gap-1.5"
                >
                  <Plus className="w-3.5 h-3.5" /> Add Single Employee
                </button>
              </div>
            </div>

            <div className="rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 overflow-hidden shadow-2xl">
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-950 text-slate-400 uppercase text-[10px] tracking-wider border-b border-slate-800 font-bold">
                    <tr>
                      <th className="py-3 px-4">Emp ID</th>
                      <th className="py-3 px-4">Employee Name</th>
                      <th className="py-3 px-3">Duty Hours</th>
                      <th className="py-3 px-3 text-right">Base Salary (₹)</th>
                      <th className="py-3 px-3 text-center">Own Tiffin (-0.5h)</th>
                      <th className="py-3 px-3 text-center">Bonus Days</th>
                      <th className="py-3 px-3 text-center">Active Status</th>
                      <th className="py-3 px-4 text-center">Actions</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50">
                    {displayedMaster.map((emp) => (
                      <tr
                        key={emp.empId}
                        className={`hover:bg-slate-800/30 transition-colors ${
                          String(emp.active).toLowerCase() === "no" ? "opacity-35 bg-slate-950/40" : ""
                        }`}
                      >
                        <td className="py-2.5 px-4 font-mono font-bold text-indigo-400">{emp.empId}</td>
                        <td className="py-2.5 px-4 font-semibold text-slate-200">{emp.name}</td>
                        <td className="py-2.5 px-3">
                          {Number(emp.dutyHours) <= 0 ? (
                            <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-500/10 text-amber-300 border border-amber-500/30">
                              Day Worker (0h)
                            </span>
                          ) : (
                            <span className="font-mono text-slate-300 font-medium">{emp.dutyHours} Hours</span>
                          )}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-200 font-medium">{formatINR(emp.salary)}</td>
                        <td className="py-2.5 px-3 text-center">
                          <span
                            className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                              String(emp.ownTiffin).toLowerCase() === "yes"
                                ? "bg-emerald-500/10 text-emerald-300 border border-emerald-500/30"
                                : "bg-slate-800 text-slate-400"
                            }`}
                          >
                            {emp.ownTiffin}
                          </span>
                        </td>
                        <td className="py-2.5 px-3 text-center">
                          <span
                            className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                              String(emp.bonusDays).toLowerCase() === "yes"
                                ? "bg-indigo-500/10 text-indigo-300 border border-indigo-500/30"
                                : "bg-slate-800 text-slate-400"
                            }`}
                          >
                            {emp.bonusDays}
                          </span>
                        </td>
                        <td className="py-2.5 px-3 text-center">
                          <span
                            className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                              String(emp.active).toLowerCase() === "yes"
                                ? "bg-emerald-500/10 text-emerald-300 border border-emerald-500/30"
                                : "bg-rose-500/10 text-rose-300 border border-rose-500/30"
                            }`}
                          >
                            {emp.active}
                          </span>
                        </td>
                        <td className="py-2.5 px-4 text-center">
                          <div className="flex items-center justify-center space-x-2">
                            <button
                              onClick={() => {
                                setEditingEmp(emp);
                                setIsEmpModalOpen(true);
                              }}
                              className="p-1 rounded hover:bg-slate-800 text-slate-400 hover:text-indigo-400 transition"
                            >
                              <Edit3 className="w-3.5 h-3.5" />
                            </button>
                            <button
                              onClick={() => handleDeleteEmp(emp.empId)}
                              className="p-1 rounded hover:bg-slate-800 text-slate-400 hover:text-rose-400 transition"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* TAB 5: BIOMETRIC UPLOADER */}
        {activeTab === "upload" && (
          <div className="max-w-3xl mx-auto space-y-6">
            <div className="rounded-2xl bg-slate-900/70 backdrop-blur-xl border border-slate-800/80 p-6 sm:p-8 shadow-2xl">
              <h2 className="font-bold text-lg text-white flex items-center gap-2">
                <Upload className="w-5 h-5 text-indigo-400" /> Biometric Attendance Excel Importer
              </h2>
              <p className="text-xs text-slate-400 mt-1">
                Reads raw biometric .xls reports. Uses exact physical column extraction: Column 0 is the label, Column 1 is merged padding, and Day 1 punch is extracted at column index 2.
              </p>

              <div
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => {
                  e.preventDefault();
                  if (e.dataTransfer.files?.[0]) parseBiometricWorkbook(e.dataTransfer.files[0]);
                }}
                className="mt-6 border-2 border-dashed border-slate-700/80 hover:border-indigo-500/80 rounded-2xl p-10 text-center transition-all bg-slate-950/40 group cursor-pointer"
              >
                <div className="w-14 h-14 rounded-2xl bg-indigo-600/10 border border-indigo-500/20 text-indigo-400 flex items-center justify-center mx-auto mb-3 group-hover:scale-110 transition">
                  <FileSpreadsheet className="w-7 h-7" />
                </div>
                <p className="text-sm font-semibold text-slate-200">Drag & drop biometric monthly report (.xlsx / .xls)</p>
                <p className="text-xs text-slate-500 mt-1">Zero column shift engine with automated header alignment</p>

                <div className="mt-5">
                  <label className="inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white font-semibold text-xs cursor-pointer shadow-lg shadow-indigo-600/30 transition">
                    <Upload className="w-4 h-4" /> Select Excel File
                    <input
                      type="file"
                      accept=".xlsx, .xls"
                      className="hidden"
                      onChange={(e) => {
                        if (e.target.files?.[0]) parseBiometricWorkbook(e.target.files[0]);
                      }}
                    />
                  </label>
                </div>
              </div>

              <div className="mt-6 pt-5 border-t border-slate-800/80 flex flex-wrap items-center justify-between gap-3 text-xs text-slate-400">
                <div className="flex items-center gap-2">
                  <Sparkles className="w-4 h-4 text-amber-400" />
                  <span>Test with reference industrial dataset:</span>
                </div>
                <button
                  onClick={loadInstantDemo}
                  className="px-3 py-1.5 rounded-lg bg-indigo-500/10 hover:bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 font-semibold transition"
                >
                  Generate Reference Dataset
                </button>
              </div>
            </div>

            <div className="rounded-2xl bg-slate-900/40 border border-slate-800/80 p-5 text-xs text-slate-400 space-y-2">
              <h4 className="font-bold text-slate-200 text-sm flex items-center gap-1.5">
                <ShieldCheck className="w-4 h-4 text-emerald-400" /> Biometric Engine Specifications
              </h4>
              <ul className="list-disc list-inside space-y-1 text-slate-400">
                <li><strong className="text-slate-300">Physical Column 2:</strong> Day 1 punch starts strictly at physical column 2 (2 + dayIndex).</li>
                <li><strong className="text-slate-300">Preserved Punch Hours:</strong> Total monthly hours are the direct sum of daily duration strings (Dalpat Singh = 252.30 Hrs).</li>
                <li><strong className="text-slate-300">Bonus Days Logic:</strong> Calculated on non-Sunday present days (26+ &#8594; 3d, 24-25 &#8594; 2d, 23 &#8594; 1d).</li>
                <li><strong className="text-slate-300">Strict Inactive Exclusion:</strong> Inactive staff contribute ₹0 to all financial summaries and KPIs.</li>
              </ul>
            </div>
          </div>
        )}
      </main>

      {/* Modal: Add / Edit Employee */}
      {isEmpModalOpen && (
        <EmployeeModal
          initialData={editingEmp}
          onClose={() => {
            setIsEmpModalOpen(false);
            setEditingEmp(null);
          }}
          onSave={handleSaveSingleEmp}
        />
      )}
    </div>
  );
}

// Sub-Component: Employee Edit/Add Modal
function EmployeeModal({ initialData, onClose, onSave }) {
  const [formData, setFormData] = useState(
    initialData || {
      empId: "",
      name: "",
      dutyHours: 9.5,
      salary: 16000,
      ownTiffin: "No",
      bonusDays: "Yes",
      active: "Yes",
      notes: ""
    }
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/75 backdrop-blur-md p-4">
      <div className="bg-[#0D111D] border border-slate-800 rounded-2xl w-full max-w-md overflow-hidden shadow-2xl animate-in fade-in zoom-in-95 duration-150">
        <div className="p-4 bg-slate-900/90 border-b border-slate-800 flex justify-between items-center">
          <h3 className="font-bold text-white text-sm">
            {initialData ? `Edit Employee #${initialData.empId}` : "Add New Employee to Master"}
          </h3>
          <button onClick={onClose} className="text-slate-400 hover:text-white">
            <X className="w-4 h-4" />
          </button>
        </div>

        <form
          onSubmit={(e) => {
            e.preventDefault();
            onSave({
              ...formData,
              dutyHours: Number(formData.dutyHours) || 0,
              salary: Number(formData.salary) || 0
            });
          }}
          className="p-5 space-y-4 text-xs"
        >
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block text-slate-300 font-semibold mb-1">
                Emp ID <span className="text-rose-500">*</span>
              </label>
              <input
                type="text"
                required
                disabled={!!initialData}
                value={formData.empId}
                onChange={(e) => setFormData({ ...formData, empId: e.target.value })}
                placeholder="e.g. 94"
                className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-slate-200 focus:outline-none focus:ring-1 focus:ring-indigo-500 disabled:opacity-50"
              />
            </div>
            <div>
              <label className="block text-slate-300 font-semibold mb-1">
                Full Name <span className="text-rose-500">*</span>
              </label>
              <input
                type="text"
                required
                value={formData.name}
                onChange={(e) => setFormData({ ...formData, name: e.target.value.toUpperCase() })}
                placeholder="e.g. SURESH KUMAR"
                className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-slate-200 focus:outline-none focus:ring-1 focus:ring-indigo-500 uppercase"
              />
            </div>
          </div>

          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block text-slate-300 font-semibold mb-1">Duty Hours (0 = Day Worker)</label>
              <input
                type="number"
                step="0.5"
                min="0"
                max="24"
                value={formData.dutyHours}
                onChange={(e) => setFormData({ ...formData, dutyHours: e.target.value })}
                className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-slate-200 focus:outline-none focus:ring-1 focus:ring-indigo-500"
              />
            </div>
            <div>
              <label className="block text-slate-300 font-semibold mb-1">Monthly Salary (₹)</label>
              <input
                type="number"
                step="500"
                min="0"
                value={formData.salary}
                onChange={(e) => setFormData({ ...formData, salary: e.target.value })}
                className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-slate-200 focus:outline-none focus:ring-1 focus:ring-indigo-500"
              />
            </div>
          </div>

          <div className="grid grid-cols-3 gap-3 pt-2">
            <div>
              <label className="block text-slate-300 font-semibold mb-1">Own Tiffin</label>
              <select
                value={formData.ownTiffin}
                onChange={(e) => setFormData({ ...formData, ownTiffin: e.target.value })}
                className="w-full px-2 py-2 bg-slate-900 border border-slate-700 rounded-lg text-slate-200 focus:outline-none focus:ring-1 focus:ring-indigo-500"
              >
                <option value="No">No</option>
                <option value="Yes">Yes (-0.5h)</option>
              </select>
            </div>

            <div>
              <label className="block text-slate-300 font-semibold mb-1">Bonus Days</label>
              <select
                value={formData.bonusDays}
                onChange={(e) => setFormData({ ...formData, bonusDays: e.target.value })}
                className="w-full px-2 py-2 bg-slate-900 border border-slate-700 rounded-lg text-slate-200 focus:outline-none focus:ring-1 focus:ring-indigo-500"
              >
                <option value="Yes">Yes</option>
                <option value="No">No</option>
              </select>
            </div>

            <div>
              <label className="block text-slate-300 font-semibold mb-1">Active Status</label>
              <select
                value={formData.active}
                onChange={(e) => setFormData({ ...formData, active: e.target.value })}
                className="w-full px-2 py-2 bg-slate-900 border border-slate-700 rounded-lg text-slate-200 focus:outline-none focus:ring-1 focus:ring-indigo-500"
              >
                <option value="Yes">Yes</option>
                <option value="No">No</option>
              </select>
            </div>
          </div>

          <div>
            <label className="block text-slate-300 font-semibold mb-1">Notes / Department</label>
            <input
              type="text"
              value={formData.notes || ""}
              onChange={(e) => setFormData({ ...formData, notes: e.target.value })}
              placeholder="e.g. Machine Maintenance"
              className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-slate-200 focus:outline-none focus:ring-1 focus:ring-indigo-500"
            />
          </div>

          <div className="pt-4 border-t border-slate-800 flex justify-end gap-2">
            <button type="button" onClick={onClose} className="px-4 py-2 bg-slate-800 hover:bg-slate-700 rounded-lg text-slate-300 font-medium">
              Cancel
            </button>
            <button type="submit" className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 rounded-lg text-white font-semibold shadow-md shadow-indigo-600/30">
              Save Employee
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}