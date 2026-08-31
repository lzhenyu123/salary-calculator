# salary-calculator
A web app to calculate and manage salary information
import React, { useState, useEffect, useCallback } from "react";
import { Plus, Trash2, ChevronRight, Receipt, Users, Save, GraduationCap, Briefcase, CalendarDays, X, UploadCloud, Pencil, RefreshCw } from "lucide-react";
import * as XLSX from "xlsx";

const uid = () => Math.random().toString(36).slice(2, 10);

const emptyEmployee = (type = "fulltime") => ({
  id: uid(),
  name: "",
  title: "",
  type, // "fulltime" | "intern" | "admin" | "boss"
  subconRole: "Technician", // "Manager" | "Head Technician" | "Service Controller" | "Technician" | "Other"
  adminConfirmed: true, // for type "admin": true = 转正, false = 没转正
  basic: type === "intern" ? 1500 : 4000,
  totalWorkDays: 26,
  daysWorked: 26,
  onTimeDays: 0,
  onTimeRate: 10,
  leaveQuota: 8,
  leaveDates: [],
  epfEnabled: true,
  epfRate: 11,
  epfManualAmount: null,
  socsoEnabled: type === "fulltime",
  socsoRate: 1.2166666667,
  socsoManualAmount: null,
  eisEnabled: type === "fulltime",
  eisRate: 0.1944444444,
  eisManualAmount: null,
  pcb: 0,
  // subcon
  subconAmount: 0,
  subconItems: [],
});

function currency(n) {
  const v = Number(n) || 0;
  const sign = v < 0 ? "-" : "";
  const truncated = Math.floor(Math.abs(v) * 100 + 1e-6) / 100;
  return sign + truncated.toLocaleString("en-MY", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function computeBasicPayroll(e) {
  const actsLikeIntern = e.type === "intern" || (e.type === "admin" && e.adminConfirmed === false);
  const rawBasic = Number(e.basic) || 0;
  const totalWorkDays = Number(e.totalWorkDays) || 0;
  const daysWorked = Number(e.daysWorked) || 0;
  const isProrated = actsLikeIntern && totalWorkDays > 0;
  const dailyRate = isProrated ? rawBasic / totalWorkDays : 0;
  const basic = isProrated ? dailyRate * daysWorked : rawBasic;
  const gross = basic;

  const isIntern = actsLikeIntern;
  const epfAuto = gross * ((Number(e.epfRate) || 0) / 100);
  const socsoAuto = gross * ((Number(e.socsoRate) || 0) / 100);
  const eisAuto = gross * ((Number(e.eisRate) || 0) / 100);
  const epf = !isIntern && e.epfEnabled ? (e.epfManualAmount != null && e.epfManualAmount !== "" ? Number(e.epfManualAmount) : epfAuto) : 0;
  const socso = !isIntern && e.socsoEnabled ? (e.socsoManualAmount != null && e.socsoManualAmount !== "" ? Number(e.socsoManualAmount) : socsoAuto) : 0;
  const eis = !isIntern && e.eisEnabled ? (e.eisManualAmount != null && e.eisManualAmount !== "" ? Number(e.eisManualAmount) : eisAuto) : 0;
  const pcb = isIntern ? 0 : Number(e.pcb) || 0;

  const leaveDates = e.leaveDates || [];
  const currentYear = new Date().getFullYear();
  const leaveDatesThisYear = leaveDates.filter((d) => new Date(d).getFullYear() === currentYear);
  const leaveQuota = Number(e.leaveQuota) || 0;
  const leaveDaysUsed = leaveDatesThisYear.length;
  const leaveExcessDays = Math.max(leaveDaysUsed - leaveQuota, 0);
  const leaveDailyRate = totalWorkDays > 0 ? rawBasic / totalWorkDays : 0;

  let leaveDeduction = 0;
  let absentDays = 0;
  let uncoveredAbsentDays = 0;
  let leaveExhausted = false;
  if (!actsLikeIntern) {
    absentDays = Math.max(totalWorkDays - daysWorked, 0);
    leaveExhausted = leaveDaysUsed >= leaveQuota;
    if (leaveExhausted && absentDays > 0 && totalWorkDays > 0) {
      uncoveredAbsentDays = absentDays;
      const netBasicAfterStatutory = basic - epf - socso - eis;
      const proratedNetBasic = (netBasicAfterStatutory / totalWorkDays) * daysWorked;
      leaveDeduction = netBasicAfterStatutory - proratedNetBasic;
    }
  }

  const totalDeduction = epf + socso + eis + pcb + leaveDeduction;
  const net = gross - totalDeduction;
  return { basic, gross, epf, socso, eis, epfAuto, socsoAuto, eisAuto, pcb, leaveDaysUsed, leaveQuota, leaveExcessDays, leaveExhausted, leaveDailyRate, leaveDeduction, absentDays, uncoveredAbsentDays, totalDeduction, net, isProrated, dailyRate, totalWorkDays, daysWorked, rawBasic, currentYear, actsLikeIntern };
}

// ---- Subcon commission engine ----
function parseNum(v) {
  if (typeof v === "number") return v;
  const cleaned = String(v == null ? "" : v).replace(/[^0-9.\-]/g, "");
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function tierRate(value, tiers) {
  // tiers: array of [minInclusive, rate], sorted ascending; picks the highest tier the value qualifies for
  let rate = 0;
  for (const [min, r] of tiers) {
    if (value >= min) rate = r;
  }
  return rate;
}

function unitLinesDescribe(i) {
  const manual = parseNum(i.manualAmount) || 0;
  const lines = i.unitLines || [];
  const parts = [];
  if (manual !== 0) parts.push(`手动 RM${manual.toLocaleString()}`);
  lines.forEach((l) => {
    const rate = parseNum(l.rate) || 0;
    const qty = parseNum(l.qty) || 0;
    if (rate !== 0 || qty !== 0) parts.push(`RM${rate.toLocaleString()} × ${qty}`);
  });
  return parts.length ? parts.join(" + ") : "尚未输入金额";
}

const COMMISSION_TYPES = [
  {
    id: "sales_commission",
    bulkFill: true,
    sumDuplicateNames: true,
    label: "Sales Commission",
    roles: (e) => e.type === "fulltime",
    excelUpload: true,
    fields: [
      { key: "personalSales", label: "Sales", unit: "RM" },
      { key: "companySales", label: "Total Sales", unit: "RM" },
      { key: "margin", label: "Margin", unit: "%" },
    ],
    compute: (i, e) => {
      const personalSales = parseNum(i.personalSales) || 0;
      const companySales = parseNum(i.companySales) || 0;
      const margin = parseNum(i.margin) || 0;
      const isManager = e.subconRole === "Manager";
      let baseRate;
      let salesForCommission = personalSales;
      if (isManager) {
        let cutoffCap;
        if (companySales >= 500000) cutoffCap = 100000;
        else if (companySales >= 400000) cutoffCap = 90000;
        else cutoffCap = 80000; // 300K or below (and anything under 400K)
        salesForCommission = Math.min(personalSales, cutoffCap);
        baseRate = tierRate(salesForCommission, [[0, 0.02], [30001, 0.03], [50001, 0.04], [80001, 0.05]]);
      } else {
        baseRate = tierRate(personalSales, [[0, 0.02], [30001, 0.03], [50001, 0.04], [80001, 0.05]]);
      }
      let marginAdj = 0;
      if (margin >= 35 && margin <= 39) marginAdj = -0.01;
      else if (margin > 0 && margin < 35) marginAdj = -0.02;
      const finalRate = Math.max(baseRate + marginAdj, 0);
      return salesForCommission * finalRate;
    },
    describe: (i, e) => {
      const personalSales = parseNum(i.personalSales) || 0;
      const companySales = parseNum(i.companySales) || 0;
      const margin = parseNum(i.margin) || 0;
      const isManager = e.subconRole === "Manager";
      let baseRatePct, cutoffText = null, salesUsed = personalSales;
      if (isManager) {
        let cutoffCap;
        if (companySales >= 500000) cutoffCap = 100000;
        else if (companySales >= 400000) cutoffCap = 90000;
        else cutoffCap = 80000;
        salesUsed = Math.min(personalSales, cutoffCap);
        baseRatePct = tierRate(salesUsed, [[0, 2], [30001, 3], [50001, 4], [80001, 5]]);
        cutoffText =
          personalSales > cutoffCap
            ? `Cutoff：Total Sales RM${companySales.toLocaleString()} → 个人业绩封顶 RM${cutoffCap.toLocaleString()}（实际 RM${personalSales.toLocaleString()} 超过封顶，按 RM${cutoffCap.toLocaleString()} 算）`
            : `Cutoff 上限 RM${cutoffCap.toLocaleString()}（个人业绩 RM${personalSales.toLocaleString()} 未超过，按实际金额算）`;
      } else {
        baseRatePct = tierRate(personalSales, [[0, 2], [30001, 3], [50001, 4], [80001, 5]]);
      }
      let marginAdj = 0;
      if (margin >= 35 && margin <= 39) marginAdj = -1;
      else if (margin > 0 && margin < 35) marginAdj = -2;
      const finalRate = Math.max(baseRatePct + marginAdj, 0);
      return `${cutoffText ? cutoffText + " · " : ""}基础 ${baseRatePct}%${marginAdj !== 0 ? ` − margin调整 ${Math.abs(marginAdj)}%` : ""} = 最终 ${finalRate}% × RM${salesUsed.toLocaleString()}`;
    },
  },
  {
    id: "technician_commission",
    bulkFill: true,
    sumDuplicateNames: true,
    label: "Technician Commission",
    roles: (e) => e.type === "fulltime" && e.subconRole !== "Manager",
    excelUpload: true,
    fields: [
      { key: "jobCount", label: "QTY", unit: "件" },
      { key: "sales", label: "Sales", unit: "RM" },
    ],
    compute: (i) => {
      const jobCount = parseNum(i.jobCount) || 0;
      const sales = parseNum(i.sales) || 0;
      const rate = tierRate(jobCount, [[0, 0.01], [31, 0.015], [61, 0.02], [101, 0.025]]);
      return sales * rate;
    },
    describe: (i) => {
      const jobCount = parseNum(i.jobCount) || 0;
      const sales = parseNum(i.sales) || 0;
      const ratePct = tierRate(jobCount, [[0, 1], [31, 1.5], [61, 2], [101, 2.5]]);
      return `QTY ${jobCount} → ${ratePct}% × RM${sales.toLocaleString()}`;
    },
  },
  {
    id: "boss_personal",
    bulkFill: true,
    sumDuplicateNames: true,
    label: "Boss no here (Sales)",
    roles: (e) => e.type === "fulltime",
    excelUpload: true,
    fields: [
      { key: "personalSales", label: "Sales", unit: "RM" },
      { key: "date", label: "日期", unit: "" },
    ],
    compute: (i) => (parseNum(i.personalSales) || 0) * 0.03,
    describe: (i) => {
      const s = parseNum(i.personalSales) || 0;
      return `3% × RM${s.toLocaleString()}${i.date ? ` · ${i.date}` : ""}`;
    },
  },
  {
    id: "boss_daily_split",
    bulkFill: true,
    label: "Boss no here (当天总业绩均分)",
    roles: (e) => e.type === "fulltime" || e.type === "admin",
    excelUpload: true,
    fields: [
      { key: "date", label: "日期", unit: "" },
      { key: "dailyTotalSales", label: "当天总业绩", unit: "RM" },
      { key: "presentCount", label: "当天到场人数", unit: "人" },
    ],
    compute: (i) => {
      const total = parseNum(i.dailyTotalSales) || 0;
      const present = parseNum(i.presentCount) || 1;
      return (total * 0.03) / present;
    },
    describe: (i) => {
      const total = parseNum(i.dailyTotalSales) || 0;
      const present = parseNum(i.presentCount) || 1;
      return `3% × RM${total.toLocaleString()} ÷ ${present} 人${i.date ? ` · ${i.date}` : ""}`;
    },
  },
  {
    id: "boss_intern",
    label: "Boss no here (实习生)",
    roles: (e) => e.type === "intern",
    excelUpload: true,
    fields: [{ key: "days", label: "出勤天数", unit: "天" }],
    compute: (i) => (parseNum(i.days) || 0) * 10,
    describe: (i) => `RM10 × ${parseNum(i.days) || 0} 天`,
  },
  {
    id: "video_commission",
    label: "Video Commission",
    roles: (e) => e.type === "fulltime" || e.type === "admin",
    excelUpload: true,
    fields: [{ key: "manualAmount", label: "手动输入金额", unit: "RM" }],
    compute: (i) => parseNum(i.manualAmount) || 0,
    describe: (i) => `手动输入 RM${(parseNum(i.manualAmount) || 0).toLocaleString()}`,
  },
  {
    id: "brake_fluid",
    bulkFill: true,
    label: "Brake Fluid",
    roles: (e) => e.type === "fulltime",
    excelUpload: true,
    fields: [
      { key: "bottles", label: "瓶数", unit: "瓶" },
      { key: "employeeCount", label: "本期员工人数", unit: "人" },
    ],
    compute: (i) => {
      const bottles = parseNum(i.bottles) || 0;
      const count = parseNum(i.employeeCount) || 1;
      return (bottles * 10) / count;
    },
    describe: (i) => {
      const bottles = parseNum(i.bottles) || 0;
      const count = parseNum(i.employeeCount) || 1;
      return `${bottles} 瓶 × RM10 ÷ ${count} 人`;
    },
  },
  {
    id: "bar",
    label: "Bar (Agent)",
    roles: (e) => e.type === "fulltime",
    excelUpload: true,
    bulkFill: true,
    sumDuplicateNames: true,
    fields: [{ key: "salePrice", label: "Total Price", unit: "RM" }],
    compute: (i) => (parseNum(i.salePrice) || 0) * 0.06,
    describe: (i) => `6% × RM${(parseNum(i.salePrice) || 0).toLocaleString()}`,
  },
  {
    id: "remap",
    bulkFill: true,
    sumDuplicateNames: true,
    label: "Remap（Employee）",
    roles: (e) => e.type === "fulltime",
    excelUpload: true,
    fields: [{ key: "quantity", label: "Inv Qty", unit: "件" }],
    compute: (i) => (parseNum(i.quantity) || 0) * 50,
    describe: (i) => `${parseNum(i.quantity) || 0} 件 × RM50`,
  },
  {
    id: "intercooler_agent",
    bulkFill: true,
    label: "KWY/RTC/KC Intercooler (Agent)",
    roles: (e) => e.type === "fulltime",
    excelUpload: true,
    fields: [{ key: "quantity", label: "Inv Qty", unit: "件" }],
    compute: (i) => (parseNum(i.quantity) || 0) * 50,
    describe: (i) => `${parseNum(i.quantity) || 0} 件 × RM50`,
  },
  {
    id: "used_part_2f",
    label: "2ND FLOOR USED PART (Agent)",
    roles: (e) => e.type === "fulltime",
    excelUpload: true,
    bulkFill: true,
    sumDuplicateNames: true,
    fields: [{ key: "salePrice", label: "Total Price", unit: "RM" }],
    compute: (i) => (parseNum(i.salePrice) || 0) * 0.06,
    describe: (i) => `6% × RM${(parseNum(i.salePrice) || 0).toLocaleString()}`,
  },
  {
    id: "pump_bearing",
    label: "Pump Bearing",
    roles: (e) => e.type === "fulltime" || e.type === "admin",
    excelUpload: false,
    hasUnitLines: true,
    fields: [{ key: "manualAmount", label: "手动输入金额", unit: "RM" }],
    compute: (i) => {
      const manual = parseNum(i.manualAmount) || 0;
      const lines = i.unitLines || [];
      const linesTotal = lines.reduce((s, l) => s + (parseNum(l.rate) || 0) * (parseNum(l.qty) || 0), 0);
      return manual + linesTotal;
    },
    describe: unitLinesDescribe,
  },
  {
    id: "skim_disc",
    label: "Skim Disc",
    roles: (e) => e.type === "fulltime" || e.type === "admin",
    excelUpload: false,
    hasUnitLines: true,
    fields: [{ key: "manualAmount", label: "手动输入金额", unit: "RM" }],
    compute: (i) => {
      const manual = parseNum(i.manualAmount) || 0;
      const lines = i.unitLines || [];
      const linesTotal = lines.reduce((s, l) => s + (parseNum(l.rate) || 0) * (parseNum(l.qty) || 0), 0);
      return manual + linesTotal;
    },
    describe: unitLinesDescribe,
  },
  {
    id: "dyno",
    label: "Dyno",
    roles: (e) => e.type === "fulltime" || e.type === "admin",
    excelUpload: true,
    hasUnitLines: true,
    fields: [{ key: "manualAmount", label: "手动输入金额", unit: "RM" }],
    compute: (i) => {
      const manual = parseNum(i.manualAmount) || 0;
      const lines = i.unitLines || [];
      const linesTotal = lines.reduce((s, l) => s + (parseNum(l.rate) || 0) * (parseNum(l.qty) || 0), 0);
      return manual + linesTotal;
    },
    describe: unitLinesDescribe,
  },
  {
    id: "walnut_blasting",
    label: "Walnut Blasting",
    roles: (e) => e.type === "fulltime" || e.type === "admin",
    excelUpload: false,
    hasUnitLines: true,
    fields: [{ key: "manualAmount", label: "手动输入金额", unit: "RM" }],
    compute: (i) => {
      const manual = parseNum(i.manualAmount) || 0;
      const lines = i.unitLines || [];
      const linesTotal = lines.reduce((s, l) => s + (parseNum(l.rate) || 0) * (parseNum(l.qty) || 0), 0);
      return manual + linesTotal;
    },
    describe: unitLinesDescribe,
  },
  {
    id: "injector_service",
    label: "Injector Service",
    roles: (e) => e.type === "fulltime" || e.type === "admin",
    excelUpload: false,
    hasUnitLines: true,
    fields: [{ key: "manualAmount", label: "手动输入金额", unit: "RM" }],
    compute: (i) => {
      const manual = parseNum(i.manualAmount) || 0;
      const lines = i.unitLines || [];
      const linesTotal = lines.reduce((s, l) => s + (parseNum(l.rate) || 0) * (parseNum(l.qty) || 0), 0);
      return manual + linesTotal;
    },
    describe: unitLinesDescribe,
  },
  {
    id: "supervisor",
    label: "Manager",
    roles: (e) => e.type === "fulltime" && e.subconRole === "Manager",
    excelUpload: true,
    fields: [{ key: "companySales", label: "Sales Performance Header", unit: "RM" }],
    compute: (i) => {
      const s = parseNum(i.companySales) || 0;
      const rate = s >= 500000 ? 0.007 : s >= 400000 ? 0.006 : s >= 300000 ? 0.005 : 0;
      return s * rate;
    },
    describe: (i) => {
      const s = parseNum(i.companySales) || 0;
      const ratePct = s >= 500000 ? 0.7 : s >= 400000 ? 0.6 : s >= 300000 ? 0.5 : 0;
      return `${ratePct}% × RM${s.toLocaleString()}`;
    },
  },
  {
    id: "head_technician_bonus",
    label: "Head Technician",
    roles: (e) => e.type === "fulltime" && e.subconRole === "Head Technician",
    excelUpload: true,
    fields: [{ key: "companySales", label: "Sales Performance Details", unit: "RM" }],
    compute: (i) => {
      const s = parseNum(i.companySales) || 0;
      const rate = s >= 500000 ? 0.008 : s >= 400000 ? 0.007 : s >= 300000 ? 0.006 : 0;
      return s * rate;
    },
    describe: (i) => {
      const s = parseNum(i.companySales) || 0;
      const ratePct = s >= 500000 ? 0.8 : s >= 400000 ? 0.7 : s >= 300000 ? 0.6 : 0;
      return `${ratePct}% × RM${s.toLocaleString()}`;
    },
  },
  {
    id: "service_controller_bonus",
    label: "Service Controller",
    roles: (e) => e.type === "fulltime" && e.subconRole === "Service Controller",
    excelUpload: true,
    fields: [{ key: "companySales", label: "Monthly Sales", unit: "RM" }],
    compute: (i) => {
      const s = parseNum(i.companySales) || 0;
      if (s >= 500000) return 1500;
      if (s >= 400000) return 800;
      if (s >= 350000) return 525;
      if (s >= 300000) return 300;
      return 0;
    },
    describe: (i) => {
      const s = parseNum(i.companySales) || 0;
      const amt = s >= 500000 ? 1500 : s >= 400000 ? 800 : s >= 350000 ? 525 : s >= 300000 ? 300 : 0;
      return `Monthly Sales RM${s.toLocaleString()} → 固定 RM${amt.toLocaleString()}`;
    },
  },
  {
    id: "core_points",
    label: "核心分",
    roles: (e) => e.type === "fulltime" || e.type === "admin",
    excelUpload: false,
    fields: (e) =>
      e.type === "admin"
        ? [
            { key: "monthlySales", label: "Monthly Sales", unit: "RM" },
            { key: "headcount", label: "参与人数", unit: "人" },
            { key: "adminCount", label: "Admin人数", unit: "人" },
          ]
        : [
            { key: "monthlySales", label: "Monthly Sales", unit: "RM" },
            { key: "headcount", label: "参与人数", unit: "人" },
          ],
    compute: (i, e) => {
      const s = parseNum(i.monthlySales) || 0;
      const headcount = parseNum(i.headcount) || 1;
      const rate = s >= 450000 ? 0.01 : s >= 400000 ? 0.008 : s >= 350000 ? 0.005 : 0;
      const coreAmount = s * rate;
      let perPerson = coreAmount / headcount;
      if (e.type === "admin") {
        const adminCount = parseNum(i.adminCount) || 1;
        perPerson = perPerson / adminCount;
      }
      return perPerson;
    },
    describe: (i, e) => {
      const s = parseNum(i.monthlySales) || 0;
      const headcount = parseNum(i.headcount) || 1;
      const ratePct = s >= 450000 ? 1 : s >= 400000 ? 0.8 : s >= 350000 ? 0.5 : 0;
      let text = `${ratePct}% × RM${s.toLocaleString()} ÷ ${headcount} 人`;
      if (e.type === "admin") {
        const adminCount = parseNum(i.adminCount) || 1;
        text += ` ÷ ${adminCount} 个Admin`;
      }
      return text;
    },
  },
];

function getCommissionType(id) {
  return COMMISSION_TYPES.find((t) => t.id === id);
}

function commissionFields(type, e) {
  return typeof type.fields === "function" ? type.fields(e) : type.fields;
}

function computeSubconPayroll(e) {
  const subconAmount = Number(e.subconAmount) || 0;
  const onTimeBonus = e.type === "intern" ? (Number(e.onTimeDays) || 0) * (Number(e.onTimeRate) || 0) : 0;
  const items = (e.subconItems || []).map((item) => {
    const type = getCommissionType(item.typeId);
    const amount = type ? type.compute(item.inputs || {}, e) : 0;
    return { ...item, amount, label: type ? type.label : item.typeId };
  });
  const itemsTotal = items.reduce((sum, it) => sum + (Number(it.amount) || 0), 0);
  const gross = subconAmount + onTimeBonus + itemsTotal;
  return { onTimeBonus, gross, net: gross, subconAmount, items, itemsTotal };
}

function computePayroll(e) {
  const basicCalc = computeBasicPayroll(e);
  const subconCalc = computeSubconPayroll(e);
  const totalNet = basicCalc.net + subconCalc.net;
  return { basic: basicCalc, subcon: subconCalc, totalNet };
}

const FieldRow = ({ label, unit, value, onChange, step = 1, min = 0, disabled }) => (
  <div className={`flex items-center justify-between py-2.5 border-b border-[#E4DFD3] ${disabled ? "opacity-40" : ""}`}>
    <span className="text-[17px] md:text-[20px] text-[#6B6558] tracking-wide">{label}</span>
    <div className="flex items-center gap-1.5">
      <input
        type="number"
        step={step}
        min={min}
        value={value}
        disabled={disabled}
        onChange={(ev) => onChange(ev.target.value)}
        className="w-28 text-right bg-transparent font-mono text-[18px] md:text-[21px] text-[#1F2937] focus:outline-none focus:bg-[#F0ECE0] rounded px-1.5 py-0.5 transition-colors disabled:cursor-not-allowed"
      />
      <span className="text-[16px] md:text-[19px] text-[#A39D8C] w-10">{unit}</span>
    </div>
  </div>
);

const ToggleRow = ({ label, checked, onChange, children }) => (
  <div className="border-b border-[#E4DFD3]">
    <div className="flex items-center justify-between py-2.5">
      <span className="text-[17px] md:text-[20px] text-[#6B6558] tracking-wide">{label}</span>
      <button
        onClick={() => onChange(!checked)}
        className={`w-9 h-5 rounded-full relative transition-colors ${checked ? "bg-[#3D6B60]" : "bg-[#D9D2BE]"}`}
      >
        <span
          className={`absolute top-0.5 w-4 h-4 rounded-full bg-white transition-transform ${
            checked ? "translate-x-4" : "translate-x-0.5"
          }`}
        />
      </button>
    </div>
    {checked && children}
  </div>
);

const LedgerLine = ({ label, value, sign, emphasis }) => (
  <div className="flex items-baseline justify-between py-1.5">
    <span className={`text-[17px] md:text-[20px] ${emphasis ? "text-[#1F2937] font-bold" : "text-[#8A8371]"}`}>{label}</span>
    <span className={`font-mono text-[18px] md:text-[21px] ${sign === "-" ? "text-[#A8453B]" : "text-[#1F2937]"} ${emphasis ? "font-extrabold text-[20px] md:text-[23px]" : ""}`}>
      {sign === "-" ? "\u2212" : ""}RM {currency(value)}
    </span>
  </div>
);

export default function PayrollSystem() {
  const [employees, setEmployees] = useState([]);
  const [selectedId, setSelectedId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saveState, setSaveState] = useState("idle");
  const [newLeaveDate, setNewLeaveDate] = useState("");
  const [leaveViewYear, setLeaveViewYear] = useState(new Date().getFullYear());
  const [payView, setPayView] = useState("basic");
  const [period, setPeriod] = useState(() => {
    const d = new Date();
    const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
    return `${months[d.getMonth()]}${d.getFullYear()}`;
  });
  const [periodInput, setPeriodInput] = useState(period);
  const [periodManifest, setPeriodManifest] = useState([]);
  const [showNewPeriodInput, setShowNewPeriodInput] = useState(false);
  const [newPeriodName, setNewPeriodName] = useState("");
  const [editingPeriod, setEditingPeriod] = useState(null);
  const [editingPeriodValue, setEditingPeriodValue] = useState("");
  const [confirmAction, setConfirmAction] = useState(null); // { type: 'delete'|'rename'|'create', period, newName }
  const [draggedPeriodIndex, setDraggedPeriodIndex] = useState(null);
  const [draggedSubconItemId, setDraggedSubconItemId] = useState(null);
  const [leaderboardTick, setLeaderboardTick] = useState(0);

  const reorderArray = (arr, fromIdx, toIdx) => {
    const copy = [...arr];
    const [moved] = copy.splice(fromIdx, 1);
    copy.splice(toIdx, 0, moved);
    return copy;
  };

  const periodKey = (p) => `payroll:period:${p.trim()}`;

  const refreshManifest = useCallback(async () => {
    const manifestRes = await window.storage.get("payroll:period-manifest", false).catch(() => null);
    const manifest = manifestRes && manifestRes.value ? JSON.parse(manifestRes.value) : [];
    setPeriodManifest(manifest);
    return manifest;
  }, []);

  const loadPeriod = useCallback(async (p, carryFromHint) => {
    setLoading(true);
    try {
      const res = await window.storage.get(periodKey(p), false);
      if (res && res.value) {
        const list = JSON.parse(res.value);
        setEmployees(list);
        setSelectedId(list.length ? list[0].id : null);
        setLoading(false);
        return;
      }
      // no snapshot for this period yet - try to carry forward roster.
      // Prefer the period we were just viewing (carryFromHint); fall back to the
      // most recently saved period in the manifest.
      const manifest = await refreshManifest();
      const candidates = [];
      if (carryFromHint && carryFromHint !== p) candidates.push(carryFromHint);
      if (manifest.length > 0) {
        const lastPeriod = manifest[manifest.length - 1];
        if (lastPeriod !== p && !candidates.includes(lastPeriod)) candidates.push(lastPeriod);
      }
      for (const src of candidates) {
        const baseRes = await window.storage.get(periodKey(src), false).catch(() => null);
        if (baseRes && baseRes.value) {
          const baseList = JSON.parse(baseRes.value);
          const carried = baseList.map((emp) => ({
            ...emp,
            daysWorked: emp.totalWorkDays,
            onTimeDays: 0,
            pcb: 0,
            subconAmount: 0,
            subconItems: [],
          }));
          setEmployees(carried);
          setSelectedId(carried.length ? carried[0].id : null);
          setLoading(false);
          return;
        }
      }
      setEmployees([]);
      setSelectedId(null);
    } catch (e) {
      setEmployees([]);
      setSelectedId(null);
    } finally {
      setLoading(false);
    }
  }, [refreshManifest]);

  useEffect(() => {
    loadPeriod(period);
    refreshManifest();
  }, []);

  const switchPeriod = (newPeriod) => {
    const trimmed = newPeriod.trim();
    if (!trimmed || trimmed === period) return;
    const fromPeriod = period;
    setPeriod(trimmed);
    setPeriodInput(trimmed);
    loadPeriod(trimmed, fromPeriod);
  };

  const createNewPeriod = () => {
    const trimmed = newPeriodName.trim();
    if (!trimmed) return;
    setConfirmAction({ type: "create", newName: trimmed });
  };

  const renamePeriod = async (oldName, newNameRaw) => {
    const newName = newNameRaw.trim();
    setEditingPeriod(null);
    if (!newName || newName === oldName) return;
    try {
      const res = await window.storage.get(periodKey(oldName), false).catch(() => null);
      const data = res && res.value ? res.value : JSON.stringify([]);
      await window.storage.set(periodKey(newName), data, false);
      await window.storage.delete(periodKey(oldName), false).catch(() => null);
      const manifest = periodManifest.map((p) => (p === oldName ? newName : p));
      await window.storage.set("payroll:period-manifest", JSON.stringify(manifest), false);
      setPeriodManifest(manifest);
      if (period === oldName) {
        setPeriod(newName);
        setPeriodInput(newName);
      }
    } catch (e) {
      // ignore
    }
  };

  const deletePeriod = async (name) => {
    try {
      await window.storage.delete(periodKey(name), false).catch(() => null);
      const manifest = periodManifest.filter((p) => p !== name);
      await window.storage.set("payroll:period-manifest", JSON.stringify(manifest), false);
      setPeriodManifest(manifest);
      if (period === name) {
        const fallback = manifest[manifest.length - 1];
        if (fallback) {
          switchPeriod(fallback);
        } else {
          setEmployees([]);
          setSelectedId(null);
        }
      }
    } catch (e) {
      // ignore
    }
  };

  const reorderPeriods = async (fromIdx, toIdx) => {
    if (fromIdx === toIdx || fromIdx == null || toIdx == null) return;
    const reordered = reorderArray(periodManifest, fromIdx, toIdx);
    setPeriodManifest(reordered);
    try {
      await window.storage.set("payroll:period-manifest", JSON.stringify(reordered), false);
    } catch (e) {
      // ignore
    }
  };

  const persist = useCallback(
    async (list) => {
      setSaveState("saving");
      try {
        await window.storage.set(periodKey(period), JSON.stringify(list), false);
        const manifestRes = await window.storage.get("payroll:period-manifest", false).catch(() => null);
        let manifest = manifestRes && manifestRes.value ? JSON.parse(manifestRes.value) : [];
        manifest = manifest.filter((p) => p !== period);
        manifest.push(period);
        await window.storage.set("payroll:period-manifest", JSON.stringify(manifest), false);
        setPeriodManifest(manifest);
        setSaveState("saved");
        setTimeout(() => setSaveState("idle"), 1200);
      } catch (e) {
        setSaveState("idle");
      }
    },
    [period]
  );

  const updateEmployees = (updater) => {
    setEmployees((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      persist(next);
      return next;
    });
  };

  const addEmployee = (type) => {
    const e = emptyEmployee(type);
    updateEmployees((prev) => [...prev, e]);
    setSelectedId(e.id);
  };

  const removeEmployee = (id) => {
    updateEmployees((prev) => prev.filter((e) => e.id !== id));
    if (selectedId === id) setSelectedId(null);
  };

  const patchEmployee = (id, patch) => {
    updateEmployees((prev) => prev.map((e) => (e.id === id ? { ...e, ...patch } : e)));
  };

  const normalizeKey = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9\u4e00-\u9fa5]/g, "");
  const truncate2 = (n) => {
    const v = Number(n);
    if (isNaN(v)) return n;
    return Math.floor(v * 100 + 1e-6) / 100;
  };

  const [uploadProgress, setUploadProgress] = useState({});
  const [uploadResult, setUploadResult] = useState({}); // key -> { ok: bool, message: string }

  // Types where uploading once fills in every matching employee that already
  // has this commission item (not just the employee you clicked upload on).
  const BULK_FILL_CONFIG = {
    brake_fluid: { fields: [{ key: "bottles", matchers: ["invqty", "inv qty", "qty"], mode: "sum" }] },
    sales_commission: {
      fields: [
        { key: "personalSales", matchers: ["sales"], mode: "sumPerRow" },
        { key: "companySales", matchers: ["sales"], mode: "sum" },
      ],
    },
    technician_commission: {
      fields: [
        { key: "jobCount", matchers: ["qty"], mode: "sumPerRow" },
        { key: "sales", matchers: ["sales"], mode: "sumPerRow" },
      ],
    },
    boss_personal: { fields: [{ key: "personalSales", matchers: ["sales"], mode: "sumPerRow" }] },
    boss_daily_split: { fields: [{ key: "dailyTotalSales", matchers: ["sales"], mode: "bottomTotal" }] },
    bar: { fields: [{ key: "salePrice", matchers: ["totalprice", "total price", "sales", "price"], mode: "sumPerRow" }] },
    used_part_2f: { fields: [{ key: "salePrice", matchers: ["totalprice", "total price", "sales", "price"], mode: "sumPerRow" }] },
    remap: { fields: [{ key: "quantity", matchers: ["invqty", "inv qty", "qty"], mode: "sumPerRow" }] },
    intercooler_agent: { fields: [{ key: "quantity", matchers: ["invqty", "inv qty", "qty"], mode: "custom" }] },
  };

  const findColumn = (keys, matchers) => keys.find((k) => matchers.some((m) => normalizeKey(k).includes(normalizeKey(m))));

  const handleBulkExcelUpload = (event, typeId) => {
    try {
      const file = event.target.files && event.target.files[0];
      event.target.value = "";
      const config = BULK_FILL_CONFIG[typeId];
      if (!file) {
        setUploadResult((p) => ({ ...p, [typeId]: { ok: false, message: "没有选到文件，请重新点击按钮选择 Excel" } }));
        return;
      }
      if (!config) {
        setUploadResult((p) => ({ ...p, [typeId]: { ok: false, message: "这个项目还不支持批量上传（内部设置缺失）" } }));
        return;
      }
      setUploadProgress((p) => ({ ...p, [typeId]: 1 }));
      setUploadResult((p) => { const n = { ...p }; delete n[typeId]; return n; });
      const reader = new FileReader();
      reader.onprogress = (evt) => {
        if (evt.lengthComputable) {
          setUploadProgress((p) => ({ ...p, [typeId]: Math.max(1, Math.round((evt.loaded / evt.total) * 100)) }));
        }
      };
      const finish = () => setUploadProgress((p) => { const n = { ...p }; delete n[typeId]; return n; });
      const fail = (message) => {
        setUploadResult((p) => ({ ...p, [typeId]: { ok: false, message } }));
        finish();
      };
      reader.onerror = () => fail("读取文件失败，请重试");
      reader.onload = (evt) => {
        try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (!rows.length) { fail("Excel 里没有读到任何数据行"); return; }
        const keys = Object.keys(rows[0]);
        const needsName = config.fields.some((f) => f.mode === "perRow" || f.mode === "sumPerRow" || f.mode === "custom");
        const nameKey = findColumn(keys, ["name", "salesman", "agent", "姓名"]);
        if (needsName && !nameKey) { fail(`找不到姓名栏（表头需含 Name/Salesman/Agent/姓名）。现有栏位：${keys.join("、")}`); return; }
        const rowMap = {};
        const rowsByName = {};
        if (nameKey) {
          rows.forEach((r) => {
            if (r[nameKey] !== "") {
              const k = normalizeKey(r[nameKey]);
              rowMap[k] = r;
              if (!rowsByName[k]) rowsByName[k] = [];
              rowsByName[k].push(r);
            }
          });
        }
        const bottomRow = rows[rows.length - 1];

        const missingCols = [];
        const sharedValues = {}; // field key -> computed shared value (bottomTotal / sum)
        config.fields.forEach((f) => {
          const col = findColumn(keys, f.matchers);
          if (!col) { missingCols.push(f.matchers.join("/")); return; }
          if (f.mode === "bottomTotal") {
            if (bottomRow[col] !== "") sharedValues[f.key] = truncate2(bottomRow[col]);
          } else if (f.mode === "sum") {
            const total = rows.reduce((s, r) => s + (Number(r[col]) || 0), 0);
            sharedValues[f.key] = truncate2(total);
          }
        });

        const dateCol = findColumn(keys, ["date", "日期"]);

        // special pre-processing for KWY/RTC/KC Intercooler (Agent):
        // only count rows whose Item Description mentions one of the three keywords,
        // excluding rows where Total Price is 0 or the description contains "SET".
        let intercoolerQtyByName = null;
        if (typeId === "intercooler_agent") {
          const descCol = findColumn(keys, ["item description", "itemdescription", "description"]);
          const qtyCol = findColumn(keys, ["invqty", "inv qty", "qty"]);
          const priceCol = findColumn(keys, ["totalprice", "total price", "price"]);
          intercoolerQtyByName = {};
          if (descCol && qtyCol) {
            rows.forEach((r) => {
              const desc = String(r[descCol] || "").toUpperCase();
              const mentionsIntercooler = ["KWY INTERCOOLER", "RTC INTERCOOLER", "KC INTERCOOLER"].some((k) => desc.includes(k));
              if (!mentionsIntercooler) return;
              if (desc.includes("SET")) return;
              const price = priceCol ? Number(r[priceCol]) || 0 : 1; // if no price column found, don't exclude on price
              if (priceCol && price === 0) return;
              const nameVal = nameKey ? r[nameKey] : "";
              if (nameVal === "") return;
              const k = normalizeKey(nameVal);
              intercoolerQtyByName[k] = (intercoolerQtyByName[k] || 0) + (Number(r[qtyCol]) || 0);
            });
          }
        }

        let matchedCount = 0;
        let candidateCount = 0;
        const uploadedAt = Date.now();
        updateEmployees((prev) =>
          prev.map((emp) => {
            const items = emp.subconItems || [];
            const matchingIdxs = items.map((it, idx) => (it.typeId === typeId ? idx : -1)).filter((idx) => idx !== -1);
            if (matchingIdxs.length === 0) return emp;
            candidateCount++;
            const newItems = [...items];
            let matchedThis = false;

            if (typeId === "intercooler_agent") {
              const qty = intercoolerQtyByName ? intercoolerQtyByName[normalizeKey(emp.name)] : undefined;
              matchingIdxs.forEach((idx) => {
                const item = items[idx];
                newItems[idx] = { ...item, inputs: { ...item.inputs, quantity: qty != null ? truncate2(qty) : 0 }, uploadedAt };
              });
              if (qty != null) matchedThis = true;
            } else if (typeId === "boss_personal" && matchingIdxs.length > 1) {
              // multiple dated entries for the same employee this period -
              // match each item to its own date's row(s) instead of summing everything into one item.
              const empRows = (rowsByName[normalizeKey(emp.name)] || []).slice();
              const salesCol = findColumn(keys, ["sales"]);
              const usedRowIdx = new Set();

              // first pass: items that already have a date - match by date
              matchingIdxs.forEach((idx) => {
                const item = items[idx];
                const itemDate = (item.inputs?.date || "").trim();
                if (!itemDate || !dateCol) return;
                const matchIdxs = empRows
                  .map((r, ri) => (!usedRowIdx.has(ri) && String(r[dateCol]).trim() === itemDate ? ri : -1))
                  .filter((ri) => ri !== -1);
                if (matchIdxs.length) {
                  matchIdxs.forEach((ri) => usedRowIdx.add(ri));
                  const total = matchIdxs.reduce((s, ri) => s + (Number(empRows[ri][salesCol]) || 0), 0);
                  newItems[idx] = { ...item, inputs: { ...item.inputs, personalSales: truncate2(total) }, uploadedAt };
                  matchedThis = true;
                }
              });
              // second pass: items without a date yet - assign remaining rows in order
              matchingIdxs.forEach((idx) => {
                const item = items[idx];
                const itemDate = (item.inputs?.date || "").trim();
                if (itemDate) return; // already handled above (or intentionally dated but unmatched - leave as is)
                const rowIdx = empRows.findIndex((r, ri) => !usedRowIdx.has(ri));
                if (rowIdx !== -1) {
                  usedRowIdx.add(rowIdx);
                  const row = empRows[rowIdx];
                  const newInputs = { ...item.inputs, personalSales: truncate2(row[salesCol]) };
                  if (dateCol && row[dateCol] !== "") newInputs.date = row[dateCol];
                  newItems[idx] = { ...item, inputs: newInputs, uploadedAt };
                  matchedThis = true;
                } else {
                  newItems[idx] = { ...item, inputs: { ...item.inputs, personalSales: 0 }, uploadedAt };
                }
              });
            } else {
              matchingIdxs.forEach((idx) => {
                const item = items[idx];
                const newInputs = { ...item.inputs };
                config.fields.forEach((f) => {
                  if (f.mode === "bottomTotal" || f.mode === "sum") {
                    // shared across everyone - if this upload doesn't have the column, reset to 0
                    newInputs[f.key] = sharedValues[f.key] != null ? sharedValues[f.key] : 0;
                    if (sharedValues[f.key] != null) matchedThis = true;
                  } else if (f.mode === "sumPerRow") {
                    const col = findColumn(keys, f.matchers);
                    const matchingRows = col ? rowsByName[normalizeKey(emp.name)] : null;
                    if (matchingRows && matchingRows.length) {
                      const total = matchingRows.reduce((s, r) => s + (Number(r[col]) || 0), 0);
                      newInputs[f.key] = truncate2(total);
                      matchedThis = true;
                    } else {
                      // this employee's name isn't in this upload - reset instead of keeping stale value
                      newInputs[f.key] = 0;
                    }
                  } else {
                    const col = findColumn(keys, f.matchers);
                    const row = col ? rowMap[normalizeKey(emp.name)] : null;
                    if (row && row[col] !== "") {
                      newInputs[f.key] = truncate2(row[col]);
                      matchedThis = true;
                    } else {
                      // this employee's name isn't in this upload - reset instead of keeping stale value
                      newInputs[f.key] = 0;
                    }
                  }
                });
                newItems[idx] = { ...item, inputs: newInputs, uploadedAt };
              });
            }

            if (matchedThis) matchedCount++;
            return { ...emp, subconItems: newItems };
          })
        );

        if (candidateCount === 0) {
          fail("目前没有任何员工加了这个 Subcon 项目，请先帮至少一位员工加上这个项目再上传");
          return;
        }
        if (matchedCount === 0) {
          fail(`Excel 里的姓名没有跟系统里任何员工对上，其余人的数字已重置为 0（读到的姓名：${Object.keys(rowMap).length ? "有" : "无"} 笔，请检查姓名拼写是否一致）`);
          return;
        }
        let msg = `已更新 ${matchedCount}/${candidateCount} 位员工`;
        if (matchedCount < candidateCount) msg += `，其余 ${candidateCount - matchedCount} 位没在表格里找到，已重置为 0`;
        if (missingCols.length) msg += `（找不到栏位：${missingCols.join("、")}）`;
        setUploadResult((p) => ({ ...p, [typeId]: { ok: true, message: msg } }));
        finish();
        } catch (err) {
          fail("解析 Excel 失败：" + (err && err.message ? err.message : "文件格式不支持"));
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (outerErr) {
      setUploadResult((p) => ({ ...p, [typeId]: { ok: false, message: "上传出错：" + (outerErr && outerErr.message ? outerErr.message : "未知错误") } }));
    }
  };

  const handleCommissionExcelUpload = (event, item, type, fields) => {
    const file = event.target.files && event.target.files[0];
    event.target.value = "";
    if (!file || !selected) return;
    setUploadProgress((p) => ({ ...p, [item.id]: 1 }));
    setUploadResult((p) => { const n = { ...p }; delete n[item.id]; return n; });
    const reader = new FileReader();
    reader.onprogress = (evt) => {
      if (evt.lengthComputable) {
        setUploadProgress((p) => ({ ...p, [item.id]: Math.max(1, Math.round((evt.loaded / evt.total) * 100)) }));
      }
    };
    const finish = () => setUploadProgress((p) => { const n = { ...p }; delete n[item.id]; return n; });
    const fail = (message) => {
      setUploadResult((p) => ({ ...p, [item.id]: { ok: false, message } }));
      finish();
    };
    reader.onerror = () => fail("读取文件失败，请重试");
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (!rows.length) { fail("Excel 里没有读到任何数据行"); return; }
        const keys = Object.keys(rows[0]);
        const nameKey = keys.find((k) => normalizeKey(k).includes("name") || normalizeKey(k).includes("agent") || normalizeKey(k).includes("姓名"));
        if (!nameKey) { fail(`找不到姓名栏（表头需含 Name/Agent/姓名）。现有栏位：${keys.join("、")}`); return; }
        const targetName = normalizeKey(selected.name);
        const matchingRows = rows.filter((r) => normalizeKey(r[nameKey]) === targetName);
        const newInputs = { ...item.inputs };
        const missingCols = [];
        const uploadedAt = Date.now();
        if (!matchingRows.length) {
          fields.forEach((f) => { newInputs[f.key] = 0; });
          const items = (selected.subconItems || []).map((it) => (it.id === item.id ? { ...it, inputs: newInputs, uploadedAt } : it));
          patchEmployee(selected.id, { subconItems: items });
          setUploadResult((p) => ({ ...p, [item.id]: { ok: false, message: `Excel 里找不到「${selected.name}」这个名字，数字已重置为 0` } }));
          finish();
          return;
        }
        const row = matchingRows[0];
        fields.forEach((f) => {
          const colKey = keys.find((k) => {
            const nk = normalizeKey(k);
            return nk.includes(normalizeKey(f.label)) || nk.includes(normalizeKey(f.key));
          });
          if (!colKey) { missingCols.push(f.label); return; }
          if (type.sumDuplicateNames && matchingRows.length > 1) {
            const total = matchingRows.reduce((s, r) => s + (Number(r[colKey]) || 0), 0);
            newInputs[f.key] = truncate2(total);
          } else if (row[colKey] !== "") {
            newInputs[f.key] = truncate2(row[colKey]);
          } else {
            missingCols.push(f.label);
          }
        });
        const items = (selected.subconItems || []).map((it) => (it.id === item.id ? { ...it, inputs: newInputs, uploadedAt } : it));
        patchEmployee(selected.id, { subconItems: items });
        let msg = type.sumDuplicateNames && matchingRows.length > 1 ? `已填入（合计 ${matchingRows.length} 笔记录）` : "已填入";
        if (missingCols.length) msg += `（找不到栏位：${missingCols.join("、")}）`;
        setUploadResult((p) => ({ ...p, [item.id]: { ok: true, message: msg } }));
        finish();
      } catch (err) {
        fail("解析 Excel 失败：" + (err && err.message ? err.message : "文件格式不支持"));
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const selected = employees.find((e) => e.id === selectedId) || null;
  const calc = selected ? computePayroll(selected) : null;
  const actsLikeIntern = selected ? selected.type === "intern" || (selected.type === "admin" && selected.adminConfirmed === false) : false;
  const totalNet = employees.reduce((sum, e) => sum + computePayroll(e).totalNet, 0);
  const fulltimeCount = employees.filter((e) => e.type === "fulltime").length;
  const internCount = employees.filter((e) => e.type === "intern").length;
  const adminCount = employees.filter((e) => e.type === "admin").length;
  const bossCount = employees.filter((e) => e.type === "boss").length;

  const buildLeaderboard = (typeId) => {
    const withItem = employees.filter((e) => (e.subconItems || []).some((it) => it.typeId === typeId));
    const rows = withItem
      .map((e) => {
        const items = (e.subconItems || []).filter((it) => it.typeId === typeId);
        const type = getCommissionType(typeId);
        let sales = 0;
        let commission = 0;
        items.forEach((it) => {
          const inputs = it.inputs || {};
          sales += parseNum(inputs.personalSales || inputs.sales || 0);
          commission += type ? type.compute(inputs, e) : 0;
        });
        return { name: e.name || "未命名员工", sales, commission };
      })
      .sort((a, b) => b.commission - a.commission)
      .slice(0, 3);
    return { rows, candidateCount: withItem.length };
  };

  const salesCommissionLeaderboard = buildLeaderboard("sales_commission");
  const technicianCommissionLeaderboard = buildLeaderboard("technician_commission");

  useEffect(() => {
    setPayView("basic");
  }, [selectedId]);

  if (loading) {
    return (
      <div className="min-h-screen bg-[#F7F4EC] flex items-center justify-center">
        <div className="text-[#8A8371] font-mono text-sm tracking-widest">加载中…</div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#F7F4EC] text-[#1F2937] font-semibold" style={{ fontFamily: "'IBM Plex Sans', 'PingFang SC', system-ui, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
        input[type=number]::-webkit-inner-spin-button,
        input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
        input[type=number] { -moz-appearance: textfield; }
      `}</style>

      <header className="border-b border-[#DDD5C2] bg-[#F7F4EC]/95 backdrop-blur sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-6 py-5 flex items-center justify-between">
          <div className="flex items-baseline gap-3">
            <h1 className="text-[26px] md:text-[29px] font-extrabold tracking-tight" style={{ fontFamily: "'Source Serif 4', serif" }}>
              薪资台账
            </h1>
            <span className="text-[16px] md:text-[19px] text-[#A39D8C] font-mono">Payroll Ledger · Malaysia</span>
          </div>
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-1.5 text-[15px] md:text-[18px] text-[#A39D8C] w-16">
              {saveState === "saving" && <><Save size={12} className="animate-pulse" /> 保存中</>}
              {saveState === "saved" && <><Save size={12} /> 已保存</>}
            </div>
          </div>
        </div>
        <div className="max-w-6xl mx-auto px-6 pb-3 flex items-center gap-1.5 overflow-x-auto">
          {periodManifest.map((p, idx) =>
            editingPeriod === p ? (
              <input
                key={p}
                autoFocus
                value={editingPeriodValue}
                onChange={(e) => setEditingPeriodValue(e.target.value)}
                onBlur={() => {
                  const trimmed = editingPeriodValue.trim();
                  if (trimmed && trimmed !== p) {
                    setConfirmAction({ type: "rename", period: p, newName: trimmed });
                  } else {
                    setEditingPeriod(null);
                  }
                }}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    const trimmed = editingPeriodValue.trim();
                    if (trimmed && trimmed !== p) {
                      setConfirmAction({ type: "rename", period: p, newName: trimmed });
                    } else {
                      setEditingPeriod(null);
                    }
                  }
                  if (e.key === "Escape") setEditingPeriod(null);
                }}
                className="shrink-0 text-[14px] md:text-[17px] font-mono font-bold px-3 py-1 rounded-t-md bg-white border border-blue-800 focus:outline-none w-36"
              />
            ) : (
              <div
                key={p}
                draggable
                onDragStart={() => setDraggedPeriodIndex(idx)}
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => {
                  e.preventDefault();
                  reorderPeriods(draggedPeriodIndex, idx);
                  setDraggedPeriodIndex(null);
                }}
                onDragEnd={() => setDraggedPeriodIndex(null)}
                className={`shrink-0 flex items-center gap-1.5 text-[14px] md:text-[17px] font-mono font-bold pl-3 pr-1.5 py-1 rounded-t-md transition-colors group cursor-move ${
                  draggedPeriodIndex === idx ? "opacity-40" : ""
                } ${p === period ? "bg-blue-800 text-white" : "bg-blue-100 text-blue-800 hover:bg-blue-200"}`}
              >
                <button
                  onClick={() => switchPeriod(p)}
                  onDoubleClick={() => {
                    setEditingPeriod(p);
                    setEditingPeriodValue(p);
                  }}
                >
                  {p}
                </button>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setEditingPeriod(p);
                    setEditingPeriodValue(p);
                  }}
                  className={`opacity-0 group-hover:opacity-70 hover:opacity-100 transition-opacity ${p === period ? "text-white" : "text-blue-800"}`}
                  title="改名字"
                >
                  <Pencil size={11} />
                </button>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setConfirmAction({ type: "delete", period: p });
                  }}
                  className={`opacity-0 group-hover:opacity-70 hover:opacity-100 transition-opacity ${p === period ? "text-white" : "text-blue-800"}`}
                  title="删除"
                >
                  <X size={12} />
                </button>
              </div>
            )
          )}
          {showNewPeriodInput ? (
            <input
              autoFocus
              value={newPeriodName}
              onChange={(e) => setNewPeriodName(e.target.value)}
              onBlur={createNewPeriod}
              onKeyDown={(e) => {
                if (e.key === "Enter") createNewPeriod();
                if (e.key === "Escape") {
                  setShowNewPeriodInput(false);
                  setNewPeriodName("");
                }
              }}
              placeholder="September2026"
              className="shrink-0 text-[14px] md:text-[17px] font-mono px-3 py-1 rounded-t-md bg-white border border-[#C9A227] focus:outline-none w-40"
            />
          ) : (
            <button
              onClick={() => setShowNewPeriodInput(true)}
              className="shrink-0 text-[14px] md:text-[17px] px-2.5 py-1 rounded-t-md text-[#A39D8C] hover:bg-[#EFE9DA] hover:text-[#3D6B60] transition-colors"
              title="新增月份 sheet"
            >
              <Plus size={16} />
            </button>
          )}
        </div>
      </header>

      {confirmAction && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 px-4">
          <div className="bg-white rounded-lg p-6 max-w-sm w-full shadow-xl">
            <p className="text-[16px] md:text-[19px] text-[#1F2937] mb-5 leading-relaxed">
              {confirmAction.type === "delete"
                ? `确定要删除「${confirmAction.period}」这个月份的记录吗？此操作无法恢复。`
                : confirmAction.type === "create"
                ? `确定要新增「${confirmAction.newName}」这个月份吗？`
                : `确定要把「${confirmAction.period}」改名为「${confirmAction.newName}」吗？`}
            </p>
            <div className="flex justify-end gap-2">
              <button
                onClick={() => {
                  setConfirmAction(null);
                  setEditingPeriod(null);
                }}
                className="px-4 py-1.5 rounded-full text-[14px] md:text-[17px] font-bold bg-[#EFE9DA] text-[#6B6558] hover:bg-[#E4DFD3] transition-colors"
              >
                取消
              </button>
              <button
                onClick={() => {
                  if (confirmAction.type === "delete") {
                    deletePeriod(confirmAction.period);
                  } else if (confirmAction.type === "create") {
                    setShowNewPeriodInput(false);
                    setNewPeriodName("");
                    switchPeriod(confirmAction.newName);
                  } else {
                    renamePeriod(confirmAction.period, confirmAction.newName);
                  }
                  setConfirmAction(null);
                }}
                className="px-4 py-1.5 rounded-full text-[14px] md:text-[17px] font-bold bg-rose-600 text-white hover:bg-rose-700 transition-colors"
              >
                确认
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-6xl mx-auto px-6 py-8 grid grid-cols-[280px_1fr] gap-8">
        <aside>
          <div className="flex items-center justify-between mb-3">
            <span className="text-[15px] md:text-[18px] uppercase tracking-widest text-[#A39D8C] flex items-center gap-1.5">
              <Users size={13} /> 员工 ({employees.length})
            </span>
          </div>

          <div className="mb-3 rounded-lg bg-gradient-to-br from-[#1F2937] to-[#3D6B60] p-3">
            <div className="flex items-center justify-between mb-2">
              <span className="text-[11px] md:text-[14px] uppercase tracking-widest text-[#C9A227] font-bold">
                🏆 Sales Commission 排行榜
              </span>
              <button
                onClick={() => setLeaderboardTick((t) => t + 1)}
                className="text-[#C9A227] hover:text-white transition-colors"
                title="刷新排名"
              >
                <RefreshCw size={13} />
              </button>
            </div>
            <div className="text-[11px] md:text-[14px] text-white/40 mb-1">
              诊断：共 {employees.length} 位员工，其中 {salesCommissionLeaderboard.candidateCount} 位加了这个项目
            </div>
            {salesCommissionLeaderboard.rows.length === 0 ? (
              <div className="text-[13px] md:text-[16px] text-white/50">目前还没有人加 Sales Commission 项目</div>
            ) : (
              <div className="space-y-1.5">
                {salesCommissionLeaderboard.rows.map((r, idx) => (
                  <div key={idx} className="flex items-center justify-between text-white">
                    <span className="flex items-center gap-1.5 text-[14px] md:text-[17px] font-bold">
                      <span>{idx === 0 ? "🥇" : idx === 1 ? "🥈" : "🥉"}</span>
                      {r.name}
                    </span>
                    <span className="flex flex-col items-end">
                      <span className="font-mono text-[14px] md:text-[17px] text-[#C9A227] font-bold">RM {currency(r.commission)}</span>
                      <span className="font-mono text-[11px] md:text-[14px] text-white/50">Sales RM {currency(r.sales)}</span>
                    </span>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div className="mb-4 rounded-lg bg-gradient-to-br from-[#1F2937] to-[#3D6B60] p-3">
            <div className="flex items-center justify-between mb-2">
              <span className="text-[11px] md:text-[14px] uppercase tracking-widest text-[#C9A227] font-bold">
                🏆 Technician Commission 排行榜
              </span>
              <button
                onClick={() => setLeaderboardTick((t) => t + 1)}
                className="text-[#C9A227] hover:text-white transition-colors"
                title="刷新排名"
              >
                <RefreshCw size={13} />
              </button>
            </div>
            <div className="text-[11px] md:text-[14px] text-white/40 mb-1">
              诊断：共 {employees.length} 位员工，其中 {technicianCommissionLeaderboard.candidateCount} 位加了这个项目
            </div>
            {technicianCommissionLeaderboard.rows.length === 0 ? (
              <div className="text-[13px] md:text-[16px] text-white/50">目前还没有人加 Technician Commission 项目</div>
            ) : (
              <div className="space-y-1.5">
                {technicianCommissionLeaderboard.rows.map((r, idx) => (
                  <div key={idx} className="flex items-center justify-between text-white">
                    <span className="flex items-center gap-1.5 text-[14px] md:text-[17px] font-bold">
                      <span>{idx === 0 ? "🥇" : idx === 1 ? "🥈" : "🥉"}</span>
                      {r.name}
                    </span>
                    <span className="flex flex-col items-end">
                      <span className="font-mono text-[14px] md:text-[17px] text-[#C9A227] font-bold">RM {currency(r.commission)}</span>
                      <span className="font-mono text-[11px] md:text-[14px] text-white/50">Sales RM {currency(r.sales)}</span>
                    </span>
                  </div>
                ))}
              </div>
            )}
          </div>
          <div className="grid grid-cols-2 gap-2 mb-3">
            <button
              onClick={() => addEmployee("fulltime")}
              className="flex items-center justify-center gap-1 text-[15px] md:text-[18px] font-bold rounded-full py-1.5 bg-emerald-100 text-emerald-700 hover:bg-emerald-200 transition-colors"
            >
              员工
            </button>
            <button
              onClick={() => addEmployee("intern")}
              className="flex items-center justify-center gap-1 text-[15px] md:text-[18px] font-bold rounded-full py-1.5 bg-amber-100 text-amber-700 hover:bg-amber-200 transition-colors"
            >
              实习生
            </button>
            <button
              onClick={() => addEmployee("admin")}
              className="flex items-center justify-center gap-1 text-[15px] md:text-[18px] font-bold rounded-full py-1.5 bg-purple-100 text-purple-700 hover:bg-purple-200 transition-colors"
            >
              Admin
            </button>
            <button
              onClick={() => addEmployee("boss")}
              className="flex items-center justify-center gap-1 text-[15px] md:text-[18px] font-bold rounded-full py-1.5 bg-rose-100 text-rose-700 hover:bg-rose-200 transition-colors"
            >
              老板
            </button>
          </div>

          <div className="space-y-1">
            {employees.length === 0 && (
              <div className="text-[17px] md:text-[20px] text-[#A39D8C] py-6 text-center border border-dashed border-[#DDD5C2] rounded-md">
                还没有员工<br />点击上方按钮新增
              </div>
            )}
            {employees.map((e) => {
              const c = computePayroll(e);
              const active = e.id === selectedId;
              return (
                <button
                  key={e.id}
                  onClick={() => setSelectedId(e.id)}
                  className={`w-full text-left px-3 py-2.5 rounded-md transition-colors group relative ${
                    active ? "bg-[#1F2937] text-[#F7F4EC]" : "hover:bg-[#EFE9DA] text-[#1F2937]"
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="text-[18px] md:text-[21px] font-bold truncate flex items-center gap-1.5">
                      <span
                        className={`w-2.5 h-2.5 rounded-full shrink-0 ${
                          e.type === "intern"
                            ? "bg-amber-500"
                            : e.type === "admin"
                            ? "bg-purple-500"
                            : e.type === "boss"
                            ? "bg-rose-500"
                            : "bg-emerald-500"
                        }`}
                      />
                      {e.name || "未命名员工"}
                    </span>
                    <ChevronRight size={14} className={active ? "opacity-60" : "opacity-0 group-hover:opacity-40"} />
                  </div>
                  <div className={`text-[16px] md:text-[19px] font-mono mt-0.5 flex items-center gap-1.5 ${active ? "text-[#C9A227]" : "text-[#8A8371]"}`}>
                    RM {currency(c.totalNet)}
                    {Number(e.subconAmount) > 0 && (
                      <span className="text-[12px] md:text-[15px] px-1.5 rounded-full bg-orange-500 text-white font-bold">
                        Subcon
                      </span>
                    )}
                  </div>
                </button>
              );
            })}
          </div>

          {employees.length > 0 && (
            <div className="mt-6 pt-4 border-t border-[#DDD5C2]">
              <div className="text-[15px] md:text-[18px] uppercase tracking-widest text-[#A39D8C] mb-1.5">
                本期实发合计 · {fulltimeCount}员工 / {internCount}实习 / {adminCount}Admin / {bossCount}老板
              </div>
              <div className="font-mono text-[24px] md:text-[27px] font-extrabold" style={{ fontFamily: "'Source Serif 4', serif" }}>
                RM {currency(totalNet)}
              </div>
            </div>
          )}
        </aside>

        <main>
          {!selected ? (
            <div className="h-full min-h-[400px] flex items-center justify-center text-[#A39D8C] text-[18px] md:text-[21px] border border-dashed border-[#DDD5C2] rounded-lg">
              选择或新增一位员工，开始计算薪资
            </div>
          ) : (
            <div className="grid grid-cols-[1fr_320px] gap-6">
              <div className="bg-white/60 border border-[#E4DFD3] rounded-lg p-6">
                <div className="flex items-center justify-between mb-1">
                  <div className="flex items-center gap-2 flex-1">
                    <input
                      value={selected.name}
                      onChange={(e) => patchEmployee(selected.id, { name: e.target.value })}
                      placeholder="员工姓名"
                      className="text-[22px] md:text-[25px] font-extrabold bg-transparent focus:outline-none placeholder:text-[#C4BDA9] flex-1"
                      style={{ fontFamily: "'Source Serif 4', serif" }}
                    />
                    {Number(selected.subconAmount) > 0 && (
                      <span className="text-[14px] md:text-[17px] font-bold px-2 py-0.5 rounded-full bg-orange-500 text-white">
                        Subcon
                      </span>
                    )}
                  </div>
                  <button
                    onClick={() => removeEmployee(selected.id)}
                    className="text-[#C4877D] hover:text-[#A8453B] transition-colors p-1"
                    title="删除员工"
                  >
                    <Trash2 size={16} />
                  </button>
                </div>
                <div className="flex items-center gap-2 mb-5">
                  <span className="text-[14px] md:text-[17px] text-[#A39D8C] uppercase tracking-widest">薪资类型</span>
                  <div className="flex gap-2 text-[15px] md:text-[18px]">
                    <button
                      onClick={() => setPayView("basic")}
                      className={`px-3 py-1 rounded-full font-bold transition-colors ${
                        payView === "basic" ? "bg-blue-600 text-white" : "bg-blue-100 text-blue-700"
                      }`}
                    >
                      Basic
                    </button>
                    <button
                      onClick={() => setPayView("subcon")}
                      className={`px-3 py-1 rounded-full font-bold transition-colors ${
                        payView === "subcon" ? "bg-orange-600 text-white" : "bg-orange-100 text-orange-700"
                      }`}
                    >
                      Subcon
                    </button>
                    <button
                      onClick={() => setPayView("total")}
                      className={`px-3 py-1 rounded-full font-bold transition-colors ${
                        payView === "total" ? "bg-emerald-700 text-white" : "bg-emerald-100 text-emerald-800"
                      }`}
                    >
                      Total · RM {currency(calc.totalNet)}
                    </button>
                  </div>
                </div>

                <div className="flex items-center gap-3 mb-4">
                  <input
                    value={selected.title}
                    onChange={(e) => patchEmployee(selected.id, { title: e.target.value })}
                    placeholder="职位（选填）"
                    className="text-[17px] md:text-[20px] text-[#8A8371] bg-transparent focus:outline-none placeholder:text-[#C4BDA9] flex-1"
                  />
                  <div className="flex gap-1.5 text-[14px] md:text-[17px] flex-wrap">
                    <button
                      onClick={() => patchEmployee(selected.id, { type: "fulltime" })}
                      className={`px-2.5 py-0.5 rounded-full font-bold transition-colors ${
                        selected.type === "fulltime" ? "bg-emerald-600 text-white" : "bg-emerald-100 text-emerald-700"
                      }`}
                    >
                      员工
                    </button>
                    <button
                      onClick={() => patchEmployee(selected.id, { type: "intern" })}
                      className={`px-2.5 py-0.5 rounded-full font-bold transition-colors ${
                        actsLikeIntern ? "bg-amber-600 text-white" : "bg-amber-100 text-amber-700"
                      }`}
                    >
                      实习生
                    </button>
                    <button
                      onClick={() => patchEmployee(selected.id, { type: "admin" })}
                      className={`px-2.5 py-0.5 rounded-full font-bold transition-colors ${
                        selected.type === "admin" ? "bg-purple-600 text-white" : "bg-purple-100 text-purple-700"
                      }`}
                    >
                      Admin
                    </button>
                    <button
                      onClick={() => patchEmployee(selected.id, { type: "boss" })}
                      className={`px-2.5 py-0.5 rounded-full font-bold transition-colors ${
                        selected.type === "boss" ? "bg-rose-600 text-white" : "bg-rose-100 text-rose-700"
                      }`}
                    >
                      老板
                    </button>
                  </div>
                </div>

                {selected.type === "fulltime" && (
                  <div className="flex items-center gap-2 mb-4">
                    <span className="text-[13px] md:text-[16px] text-[#A39D8C] uppercase tracking-widest">职务</span>
                    <select
                      value={selected.subconRole}
                      onChange={(e) => patchEmployee(selected.id, { subconRole: e.target.value })}
                      className="text-[14px] md:text-[17px] bg-transparent border border-[#DDD5C2] rounded px-2 py-1 focus:outline-none focus:border-[#3D6B60]"
                    >
                      <option value="Technician">Technician</option>
                      <option value="Manager">Manager</option>
                      <option value="Head Technician">Head Technician</option>
                      <option value="Service Controller">Service Controller</option>
                      <option value="Other">Other</option>
                    </select>
                  </div>
                )}

                {selected.type === "admin" && (
                  <div className="flex items-center gap-2 mb-4">
                    <span className="text-[13px] md:text-[16px] text-[#A39D8C] uppercase tracking-widest">状态</span>
                    <div className="flex gap-1.5 text-[13px] md:text-[16px]">
                      <button
                        onClick={() => patchEmployee(selected.id, { adminConfirmed: true })}
                        className={`px-2.5 py-0.5 rounded-full font-bold transition-colors ${
                          selected.adminConfirmed !== false ? "bg-purple-700 text-white" : "bg-purple-100 text-purple-700"
                        }`}
                      >
                        转正
                      </button>
                      <button
                        onClick={() => patchEmployee(selected.id, { adminConfirmed: false })}
                        className={`px-2.5 py-0.5 rounded-full font-bold transition-colors ${
                          selected.adminConfirmed === false ? "bg-purple-700 text-white" : "bg-purple-100 text-purple-700"
                        }`}
                      >
                        没转正
                      </button>
                    </div>
                  </div>
                )}

                {payView === "total" ? (
                  <div className="py-8 px-6 border border-dashed border-[#DDD5C2] rounded-lg space-y-3">
                    <div className="flex items-center justify-between text-[16px] md:text-[19px]">
                      <span className="text-blue-700 font-bold">Basic 净额</span>
                      <span className="font-mono">RM {currency(calc.basic.net)}</span>
                    </div>
                    <div className="pl-3 border-l-2 border-orange-200 space-y-1">
                      <span className="text-orange-700 font-bold text-[16px] md:text-[19px]">Subcon 明细</span>
                      {calc.subcon.items.length === 0 && calc.subcon.subconAmount === 0 && calc.subcon.onTimeBonus === 0 ? (
                        <div className="text-[14px] md:text-[17px] text-[#A39D8C]">暂无 Subcon 项目</div>
                      ) : (
                        <>
                          {calc.subcon.items.map((it) => (
                            <div key={it.id} className="flex items-center justify-between text-[14px] md:text-[17px]">
                              <span className="text-[#6B6558]">{it.label}</span>
                              <span className="font-mono">RM {currency(it.amount)}</span>
                            </div>
                          ))}
                          {calc.subcon.subconAmount !== 0 && (
                            <div className="flex items-center justify-between text-[14px] md:text-[17px]">
                              <span className="text-[#6B6558]">其他手动金额</span>
                              <span className="font-mono">RM {currency(calc.subcon.subconAmount)}</span>
                            </div>
                          )}
                          {calc.subcon.onTimeBonus !== 0 && (
                            <div className="flex items-center justify-between text-[14px] md:text-[17px]">
                              <span className="text-[#6B6558]">准时补贴</span>
                              <span className="font-mono">RM {currency(calc.subcon.onTimeBonus)}</span>
                            </div>
                          )}
                        </>
                      )}
                      <div className="flex items-center justify-between text-[15px] md:text-[18px] pt-1">
                        <span className="text-orange-700 font-bold">Subcon 净额</span>
                        <span className="font-mono font-bold">RM {currency(calc.subcon.net)}</span>
                      </div>
                    </div>
                    <div className="flex items-center justify-between text-[18px] md:text-[21px] pt-3 border-t border-[#DDD5C2]">
                      <span className="text-emerald-800 font-extrabold">Total 合计</span>
                      <span className="font-mono font-extrabold">RM {currency(calc.totalNet)}</span>
                    </div>
                  </div>
                ) : payView === "subcon" ? (
                  <div className="space-y-4">
                    <div className="flex items-center gap-2 flex-wrap">
                      {COMMISSION_TYPES.filter((t) => t.roles(selected)).map((t) => (
                        <button
                          key={t.id}
                          onClick={() => {
                            const fields = commissionFields(t, selected);
                            const inputs = {};
                            fields.forEach((f) => (inputs[f.key] = ""));
                            const newItem = { id: uid(), typeId: t.id, inputs };
                            patchEmployee(selected.id, { subconItems: [...(selected.subconItems || []), newItem] });
                          }}
                          className="text-[13px] md:text-[16px] px-2.5 py-1 rounded-full border border-orange-300 text-orange-700 hover:border-orange-600 hover:bg-orange-50 transition-colors"
                        >
                          + {t.label}
                        </button>
                      ))}
                    </div>

                    {(selected.subconItems || []).length === 0 && (
                      <div className="text-[15px] md:text-[18px] text-[#A39D8C] text-center py-6 border border-dashed border-[#DDD5C2] rounded-lg">
                        点击上方按钮加入一个 Subcon 项目
                      </div>
                    )}

                    {(selected.subconItems || []).map((item, idx, arr) => {
                      const type = getCommissionType(item.typeId);
                      if (!type) return null;
                      const fields = commissionFields(type, selected);
                      const amount = type.compute(item.inputs || {}, selected);
                      const description = type.describe ? type.describe(item.inputs || {}, selected) : null;
                      const isRepeatOfPrevious = idx > 0 && arr[idx - 1].typeId === item.typeId;
                      return (
                        <div
                          key={item.id}
                          draggable
                          onDragStart={() => setDraggedSubconItemId(item.id)}
                          onDragOver={(e) => e.preventDefault()}
                          onDrop={(e) => {
                            e.preventDefault();
                            const items = selected.subconItems || [];
                            const fromIdx = items.findIndex((it) => it.id === draggedSubconItemId);
                            const toIdx = items.findIndex((it) => it.id === item.id);
                            if (fromIdx !== -1 && toIdx !== -1 && fromIdx !== toIdx) {
                              patchEmployee(selected.id, { subconItems: reorderArray(items, fromIdx, toIdx) });
                            }
                            setDraggedSubconItemId(null);
                          }}
                          onDragEnd={() => setDraggedSubconItemId(null)}
                          className={`border border-[#E4DFD3] rounded-lg p-4 bg-white/60 cursor-move ${
                            draggedSubconItemId === item.id ? "opacity-40" : ""
                          } ${isRepeatOfPrevious ? "mt-[-8px]" : ""}`}
                        >
                          <div className="flex items-center justify-between mb-2">
                            {isRepeatOfPrevious ? (
                              <span className="text-[12px] md:text-[15px] text-[#A39D8C]">又一笔</span>
                            ) : (
                              <span className="text-[15px] md:text-[18px] font-bold text-orange-700">{type.label}</span>
                            )}
                            <div className="flex items-center gap-3">
                              <span className="font-mono text-[15px] md:text-[18px] text-[#1F2937]">RM {currency(amount)}</span>
                              <button
                                onClick={() => {
                                  const emptyInputs = {};
                                  fields.forEach((f) => (emptyInputs[f.key] = ""));
                                  const newItem = { id: uid(), typeId: item.typeId, inputs: emptyInputs };
                                  const items = [...(selected.subconItems || [])];
                                  const idx = items.findIndex((it) => it.id === item.id);
                                  items.splice(idx + 1, 0, newItem);
                                  patchEmployee(selected.id, { subconItems: items });
                                }}
                                className="text-[#3D6B60] hover:text-[#2A4C44]"
                                title="再加一笔同类型的记录"
                              >
                                <Plus size={16} />
                              </button>
                              <button
                                onClick={() => {
                                  const items = (selected.subconItems || []).filter((it) => it.id !== item.id);
                                  patchEmployee(selected.id, { subconItems: items });
                                }}
                                className="text-[#A39D8C] hover:text-[#A8453B]"
                              >
                                <X size={16} />
                              </button>
                            </div>
                          </div>
                          {type.excelUpload && (
                            <>
                              <div className="flex items-center gap-2 mb-1 flex-wrap">
                                <label className="inline-flex items-center gap-1.5 text-[12px] md:text-[15px] text-[#3D6B60] border border-[#3D6B60]/30 hover:border-[#3D6B60] rounded-full px-2.5 py-1 cursor-pointer transition-colors">
                                  <UploadCloud size={13} />
                                  {uploadProgress[type.bulkFill ? type.id : item.id] != null
                                    ? `上传中 ${uploadProgress[type.bulkFill ? type.id : item.id]}%`
                                    : type.bulkFill
                                    ? "上传 Excel（自动分配给所有该项目的人）"
                                    : "上传 Excel 自动填入"}
                                  <input
                                    type="file"
                                    accept=".xlsx,.xls"
                                    className="hidden"
                                    onChange={(e) =>
                                      type.bulkFill ? handleBulkExcelUpload(e, type.id) : handleCommissionExcelUpload(e, item, type, fields)
                                    }
                                  />
                                </label>
                                {item.uploadedAt && (
                                  <span className="text-[11px] md:text-[14px] text-[#A39D8C]">
                                    上次上传：{new Date(item.uploadedAt).toLocaleString("zh-CN", { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" })}
                                  </span>
                                )}
                              </div>
                              {uploadResult[type.bulkFill ? type.id : item.id] && (
                                <div
                                  className={`text-[12px] md:text-[15px] mb-2 ${
                                    uploadResult[type.bulkFill ? type.id : item.id].ok ? "text-[#3D6B60]" : "text-[#A8453B]"
                                  }`}
                                >
                                  {uploadResult[type.bulkFill ? type.id : item.id].message}
                                </div>
                              )}
                            </>
                          )}
                          <div className="grid grid-cols-2 gap-x-4">
                            {fields.map((f) => (
                              <div key={f.key} className="flex items-center justify-between py-1.5 border-b border-[#E4DFD3]">
                                <span className="text-[13px] md:text-[16px] text-[#6B6558]">{f.label}</span>
                                <div className="flex items-center gap-1">
                                  <input
                                    value={item.inputs?.[f.key] ?? ""}
                                    onChange={(e) => {
                                      const val = e.target.value;
                                      const syncedTypes = {
                                        brake_fluid: ["employeeCount", "bottles"],
                                        core_points: ["monthlySales", "headcount", "adminCount"],
                                        sales_commission: ["companySales"],
                                      };
                                      const syncFields = syncedTypes[type.id];
                                      if (type.id === "boss_daily_split" && f.key === "date") {
                                        // joining a date group - inherit that group's total/headcount if it already exists
                                        let inherited = null;
                                        (function findGroup(list) {
                                          for (const emp2 of list) {
                                            for (const it2 of emp2.subconItems || []) {
                                              if (it2.typeId === "boss_daily_split" && it2.id !== item.id && val.trim() !== "" && (it2.inputs?.date || "").trim() === val.trim()) {
                                                inherited = it2.inputs;
                                                return;
                                              }
                                            }
                                          }
                                        })(employees);
                                        const items = (selected.subconItems || []).map((it) =>
                                          it.id === item.id
                                            ? {
                                                ...it,
                                                inputs: {
                                                  ...it.inputs,
                                                  date: val,
                                                  ...(inherited ? { dailyTotalSales: inherited.dailyTotalSales, presentCount: inherited.presentCount } : {}),
                                                },
                                              }
                                            : it
                                        );
                                        patchEmployee(selected.id, { subconItems: items });
                                      } else if (type.id === "boss_daily_split" && (f.key === "dailyTotalSales" || f.key === "presentCount")) {
                                        // shared per DATE, not globally - only sync to items with the same date
                                        const itemDate = (item.inputs?.date || "").trim();
                                        updateEmployees((prev) =>
                                          prev.map((emp) => {
                                            const items = emp.subconItems || [];
                                            let changed = false;
                                            const newItems = items.map((it) => {
                                              if (it.typeId !== "boss_daily_split") return it;
                                              const sameItem = it.id === item.id;
                                              const sameDate = itemDate !== "" && (it.inputs?.date || "").trim() === itemDate;
                                              if (!sameItem && !sameDate) return it;
                                              changed = true;
                                              return { ...it, inputs: { ...it.inputs, [f.key]: val } };
                                            });
                                            return changed ? { ...emp, subconItems: newItems } : emp;
                                          })
                                        );
                                      } else if (syncFields && syncFields.includes(f.key)) {
                                        // shared totals for the whole period - sync across every employee with the same item type.
                                        // the item actually being edited is always updated by its own id (so this is safe even
                                        // when this employee has multiple items of the same type); other employees fall back to
                                        // their first item of this type.
                                        updateEmployees((prev) =>
                                          prev.map((emp) => {
                                            const items = emp.subconItems || [];
                                            const isCurrentEmp = emp.id === selected.id;
                                            const i = isCurrentEmp ? items.findIndex((it) => it.id === item.id) : items.findIndex((it) => it.typeId === type.id);
                                            if (i === -1) return emp;
                                            const newItems = [...items];
                                            newItems[i] = { ...items[i], inputs: { ...items[i].inputs, [f.key]: val } };
                                            return { ...emp, subconItems: newItems };
                                          })
                                        );
                                      } else {
                                        const items = (selected.subconItems || []).map((it) =>
                                          it.id === item.id ? { ...it, inputs: { ...it.inputs, [f.key]: val } } : it
                                        );
                                        patchEmployee(selected.id, { subconItems: items });
                                      }
                                    }}
                                    className="w-24 text-right bg-transparent font-mono text-[13px] md:text-[16px] focus:outline-none focus:bg-[#F0ECE0] rounded px-1 py-0.5"
                                  />
                                  <span className="text-[11px] md:text-[14px] text-[#A39D8C] w-8">{f.unit}</span>
                                </div>
                              </div>
                            ))}
                          </div>
                          {type.hasUnitLines && (
                            <div className="mt-2 pt-2 border-t border-[#E4DFD3]">
                              <div className="text-[11px] md:text-[14px] text-[#A39D8C] uppercase tracking-widest mb-1">RM/Unit × QTY（可加多个）</div>
                              {(item.inputs?.unitLines || []).map((line, lineIdx) => (
                                <div key={lineIdx} className="flex items-center gap-1.5 py-1">
                                  <input
                                    value={line.rate ?? ""}
                                    onChange={(e) => {
                                      const lines = [...(item.inputs?.unitLines || [])];
                                      lines[lineIdx] = { ...lines[lineIdx], rate: e.target.value };
                                      const items = (selected.subconItems || []).map((it) =>
                                        it.id === item.id ? { ...it, inputs: { ...it.inputs, unitLines: lines } } : it
                                      );
                                      patchEmployee(selected.id, { subconItems: items });
                                    }}
                                    placeholder="RM/Unit"
                                    className="w-20 text-right bg-transparent border-b border-[#DDD5C2] font-mono text-[13px] md:text-[16px] focus:outline-none focus:border-[#3D6B60] px-1"
                                  />
                                  <span className="text-[12px] md:text-[15px] text-[#A39D8C]">×</span>
                                  <input
                                    value={line.qty ?? ""}
                                    onChange={(e) => {
                                      const lines = [...(item.inputs?.unitLines || [])];
                                      lines[lineIdx] = { ...lines[lineIdx], qty: e.target.value };
                                      const items = (selected.subconItems || []).map((it) =>
                                        it.id === item.id ? { ...it, inputs: { ...it.inputs, unitLines: lines } } : it
                                      );
                                      patchEmployee(selected.id, { subconItems: items });
                                    }}
                                    placeholder="QTY"
                                    className="w-16 text-right bg-transparent border-b border-[#DDD5C2] font-mono text-[13px] md:text-[16px] focus:outline-none focus:border-[#3D6B60] px-1"
                                  />
                                  <span className="text-[12px] md:text-[15px] text-[#A39D8C] flex-1 text-right font-mono">
                                    = RM {currency((parseNum(line.rate) || 0) * (parseNum(line.qty) || 0))}
                                  </span>
                                  <button
                                    onClick={() => {
                                      const lines = (item.inputs?.unitLines || []).filter((_, idx) => idx !== lineIdx);
                                      const items = (selected.subconItems || []).map((it) =>
                                        it.id === item.id ? { ...it, inputs: { ...it.inputs, unitLines: lines } } : it
                                      );
                                      patchEmployee(selected.id, { subconItems: items });
                                    }}
                                    className="text-[#A39D8C] hover:text-[#A8453B]"
                                  >
                                    <X size={14} />
                                  </button>
                                </div>
                              ))}
                              <button
                                onClick={() => {
                                  const lines = [...(item.inputs?.unitLines || []), { rate: "", qty: "" }];
                                  const items = (selected.subconItems || []).map((it) =>
                                    it.id === item.id ? { ...it, inputs: { ...it.inputs, unitLines: lines } } : it
                                  );
                                  patchEmployee(selected.id, { subconItems: items });
                                }}
                                className="text-[12px] md:text-[15px] text-[#3D6B60] hover:text-[#2A4C44] flex items-center gap-1 mt-1"
                              >
                                <Plus size={12} /> 加一行
                              </button>
                            </div>
                          )}
                          {description && (
                            <div className="mt-2 text-[12px] md:text-[15px] text-[#8A6D3D] bg-orange-50 rounded px-2 py-1.5 leading-relaxed">
                              {description}
                            </div>
                          )}
                        </div>
                      );
                    })}

                    <div className="flex items-center gap-1.5 pt-2">
                      <span className="text-[15px] md:text-[18px] text-[#6B6558]">其他手动金额</span>
                      <input
                        type="number"
                        value={selected.subconAmount}
                        onChange={(e) => patchEmployee(selected.id, { subconAmount: e.target.value })}
                        className="w-28 text-right bg-transparent border-b border-[#DDD5C2] font-mono text-[16px] md:text-[19px] focus:outline-none focus:border-[#C9A227] px-1"
                      />
                      <span className="text-[14px] md:text-[17px] text-[#A39D8C]">RM</span>
                    </div>

                    {selected.type === "intern" && (
                      <div className="max-w-xs">
                        <FieldRow label="准时到岗天数 (9:15前)" unit="天" value={selected.onTimeDays} onChange={(v) => patchEmployee(selected.id, { onTimeDays: v })} step={1} />
                        <FieldRow label="准时补贴单价" unit="RM/天" value={selected.onTimeRate} onChange={(v) => patchEmployee(selected.id, { onTimeRate: v })} step={1} />
                      </div>
                    )}
                  </div>
                ) : (
                <div className="mt-2">
                  <div className="text-[15px] md:text-[18px] uppercase tracking-widest text-[#A39D8C] mb-1">收入</div>
                  <FieldRow
                    label={actsLikeIntern ? "底薪 Basic（整月）" : "底薪 Basic"}
                    unit="RM"
                    value={selected.basic}
                    onChange={(v) => patchEmployee(selected.id, { basic: v })}
                    step={50}
                  />
                  {!actsLikeIntern && (
                    <>
                      <FieldRow label="当月上班天数" unit="天" value={selected.totalWorkDays} onChange={(v) => patchEmployee(selected.id, { totalWorkDays: v })} step={1} />
                      <FieldRow label="实际出勤天数" unit="天" value={selected.daysWorked} onChange={(v) => patchEmployee(selected.id, { daysWorked: v })} step={0.5} />
                      {calc.basic.absentDays > 0 && (
                        <div className={`flex items-center justify-between py-2 text-[16px] md:text-[19px] border-b border-[#E4DFD3] ${calc.basic.uncoveredAbsentDays > 0 ? "text-[#A8453B]" : "text-[#8A8371]"}`}>
                          <span>
                            缺勤 {calc.basic.absentDays} 天
                            {calc.basic.uncoveredAbsentDays > 0
                              ? `，年假已用完，其中 ${calc.basic.uncoveredAbsentDays} 天扣薪`
                              : "，年假余额足够覆盖，暂不扣薪"}
                          </span>
                        </div>
                      )}
                    </>
                  )}
                  {actsLikeIntern && (
                    <>
                      <FieldRow label="当月应上班天数" unit="天" value={selected.totalWorkDays} onChange={(v) => patchEmployee(selected.id, { totalWorkDays: v })} step={1} />
                      <FieldRow label="实际出勤天数" unit="天" value={selected.daysWorked} onChange={(v) => patchEmployee(selected.id, { daysWorked: v })} step={0.5} />
                      <div className="flex items-center justify-between py-2 text-[16px] md:text-[19px] text-[#8A6D3D] border-b border-[#E4DFD3]">
                        <span>按出勤折算底薪</span>
                        <span className="font-mono">
                          RM {currency(selected.basic)} ÷ {selected.totalWorkDays || 0} × {selected.daysWorked || 0} = RM {currency(calc.basic.basic)}
                        </span>
                      </div>
                    </>
                  )}

                  {!actsLikeIntern && (
                    <>
                      <div className="text-[15px] md:text-[18px] uppercase tracking-widest text-[#A39D8C] mb-1 mt-5">法定扣除</div>
                      <ToggleRow
                        label="EPF 公积金"
                        checked={selected.epfEnabled}
                        onChange={(v) => patchEmployee(selected.id, { epfEnabled: v })}
                      >
                        <FieldRow label="员工比例" unit="%" value={selected.epfRate} onChange={(v) => patchEmployee(selected.id, { epfRate: v })} step={0.5} />
                        <FieldRow
                          label="金额（可手动改）"
                          unit="RM"
                          value={selected.epfManualAmount != null ? selected.epfManualAmount : Number(currency(calc.basic.epfAuto).replace(/,/g, ""))}
                          onChange={(v) => patchEmployee(selected.id, { epfManualAmount: v })}
                          step={1}
                        />
                      </ToggleRow>
                      <ToggleRow
                        label="SOCSO 社险"
                        checked={selected.socsoEnabled}
                        onChange={(v) => patchEmployee(selected.id, { socsoEnabled: v })}
                      >
                        <FieldRow label="员工比例（估算）" unit="%" value={selected.socsoRate} onChange={(v) => patchEmployee(selected.id, { socsoRate: v })} step={0.0001} />
                        <FieldRow
                          label="金额（可手动改）"
                          unit="RM"
                          value={selected.socsoManualAmount != null ? selected.socsoManualAmount : Number(currency(calc.basic.socsoAuto).replace(/,/g, ""))}
                          onChange={(v) => patchEmployee(selected.id, { socsoManualAmount: v })}
                          step={1}
                        />
                      </ToggleRow>
                      <ToggleRow
                        label="EIS 就业保险"
                        checked={selected.eisEnabled}
                        onChange={(v) => patchEmployee(selected.id, { eisEnabled: v })}
                      >
                        <FieldRow label="员工比例" unit="%" value={selected.eisRate} onChange={(v) => patchEmployee(selected.id, { eisRate: v })} step={0.0001} />
                        <FieldRow
                          label="金额（可手动改）"
                          unit="RM"
                          value={selected.eisManualAmount != null ? selected.eisManualAmount : Number(currency(calc.basic.eisAuto).replace(/,/g, ""))}
                          onChange={(v) => patchEmployee(selected.id, { eisManualAmount: v })}
                          step={1}
                        />
                      </ToggleRow>

                      <div className="text-[15px] md:text-[18px] uppercase tracking-widest text-[#A39D8C] mb-1 mt-5">其他</div>
                      <FieldRow label="PCB 所得税" unit="RM" value={selected.pcb} onChange={(v) => patchEmployee(selected.id, { pcb: v })} step={10} />

                      <div className="text-[15px] md:text-[18px] uppercase tracking-widest text-[#A39D8C] mb-1 mt-5 flex items-center gap-1.5">
                        <CalendarDays size={12} /> 年假 Annual Leave
                      </div>
                      <FieldRow label="每年年假额度" unit="天" value={selected.leaveQuota} onChange={(v) => patchEmployee(selected.id, { leaveQuota: v })} step={1} />

                  <div className="py-2 border-b border-[#E4DFD3]">
                    <div className="flex items-center gap-2 mb-2">
                      <input
                        type="date"
                        value={newLeaveDate}
                        onChange={(e) => setNewLeaveDate(e.target.value)}
                        className="text-[17px] md:text-[20px] bg-transparent border border-[#DDD5C2] rounded px-2 py-1 focus:outline-none focus:border-[#3D6B60] text-[#1F2937]"
                      />
                      <button
                        onClick={() => {
                          if (!newLeaveDate) return;
                          const dates = [...(selected.leaveDates || []), newLeaveDate].sort();
                          patchEmployee(selected.id, { leaveDates: dates });
                          setNewLeaveDate("");
                        }}
                        className="flex items-center gap-1 text-[16px] md:text-[19px] text-[#3D6B60] hover:text-[#2A4C44] border border-[#3D6B60]/30 hover:border-[#3D6B60] rounded-md px-2.5 py-1 transition-colors"
                      >
                        <Plus size={13} /> 记录请假
                      </button>
                    </div>

                    {(() => {
                      const years = Array.from(
                        new Set([...(selected.leaveDates || []).map((d) => new Date(d).getFullYear()), calc.basic.currentYear])
                      ).sort((a, b) => a - b);
                      const datesForYear = (selected.leaveDates || []).filter(
                        (d) => new Date(d).getFullYear() === leaveViewYear
                      );
                      const yearIndex = years.indexOf(leaveViewYear);
                      return (
                        <>
                          <div className="flex items-center justify-between mb-2">
                            <button
                              onClick={() => yearIndex > 0 && setLeaveViewYear(years[yearIndex - 1])}
                              disabled={yearIndex <= 0}
                              className="text-[#A39D8C] hover:text-[#3D6B60] disabled:opacity-30 disabled:hover:text-[#A39D8C] px-1"
                            >
                              ‹
                            </button>
                            <span className="text-[17px] md:text-[20px] font-mono text-[#6B6558]">
                              {leaveViewYear}年
                              {leaveViewYear === calc.basic.currentYear && <span className="text-[14px] md:text-[17px] text-[#3D6B60] ml-1">（本年度）</span>}
                            </span>
                            <button
                              onClick={() => yearIndex < years.length - 1 && setLeaveViewYear(years[yearIndex + 1])}
                              disabled={yearIndex >= years.length - 1}
                              className="text-[#A39D8C] hover:text-[#3D6B60] disabled:opacity-30 disabled:hover:text-[#A39D8C] px-1"
                            >
                              ›
                            </button>
                          </div>

                          {datesForYear.length === 0 ? (
                            <div className="text-[16px] md:text-[19px] text-[#A39D8C] py-1 text-center">{leaveViewYear}年没有请假记录</div>
                          ) : (
                            <div className="flex flex-wrap gap-1.5 mb-2">
                              {datesForYear.map((d) => {
                                const i = (selected.leaveDates || []).indexOf(d);
                                return (
                                  <span
                                    key={`${d}-${i}`}
                                    className="flex items-center gap-1 text-[16px] md:text-[19px] font-mono bg-[#EFE9DA] text-[#6B6558] rounded-full pl-2.5 pr-1 py-1"
                                  >
                                    {d}
                                    <button
                                      onClick={() => {
                                        const dates = (selected.leaveDates || []).filter((_, idx) => idx !== i);
                                        patchEmployee(selected.id, { leaveDates: dates });
                                      }}
                                      className="text-[#A39D8C] hover:text-[#A8453B] transition-colors"
                                    >
                                      <X size={12} />
                                    </button>
                                  </span>
                                );
                              })}
                            </div>
                          )}
                        </>
                      );
                    })()}

                    <div className={`text-[16px] md:text-[19px] font-mono ${calc.basic.leaveExcessDays > 0 ? "text-[#A8453B]" : "text-[#8A8371]"}`}>
                      {calc.basic.currentYear}年度已用 {calc.basic.leaveDaysUsed} / {calc.basic.leaveQuota} 天
                      {actsLikeIntern && calc.basic.leaveExcessDays > 0 && (
                        <span> · 超出 {calc.basic.leaveExcessDays} 天，按日薪 RM {currency(calc.basic.leaveDailyRate)} 扣款</span>
                      )}
                      {!actsLikeIntern && calc.basic.leaveExcessDays > 0 && (
                        <span> · 年假余额已用完，缺勤将开始扣薪</span>
                      )}
                    </div>
                  </div>
                    </>
                  )}
                </div>
                )}
              </div>

              <div className="bg-[#1F2937] text-[#F7F4EC] rounded-lg p-6 h-fit sticky top-24">
                {payView !== "basic" && (
                  <div className="flex items-center gap-2 mb-4 text-[#C9A227]">
                    <Receipt size={15} />
                    <span className="text-[15px] md:text-[18px] uppercase tracking-widest">本期结算单 Payslip</span>
                  </div>
                )}
                <div className="divide-y divide-white/10">
                  {payView === "total" ? (
                    <div className="pb-2 mb-1">
                      <LedgerLine label="Basic 净额" value={calc.basic.net} />
                      <LedgerLine label="Subcon 净额" value={calc.subcon.net} />
                    </div>
                  ) : payView === "subcon" ? (
                    <div className="pb-2 mb-1">
                      {calc.subcon.items.map((it) => (
                        <LedgerLine key={it.id} label={it.label} value={it.amount} />
                      ))}
                      {calc.subcon.subconAmount !== 0 && <LedgerLine label="其他手动金额" value={calc.subcon.subconAmount} />}
                      {selected.type === "intern" && (
                        <LedgerLine label={`准时补贴 (${selected.onTimeDays || 0}天)`} value={calc.subcon.onTimeBonus} />
                      )}
                      <div className="text-[14px] md:text-[17px] text-white/35 pt-1">不扣 EPF / SOCSO / EIS，全额发放</div>
                    </div>
                  ) : (
                    <>
                  <div className="pb-2 mb-1">
                    <LedgerLine
                      label={calc.basic.isProrated ? `底薪 Basic (${calc.basic.daysWorked}/${calc.basic.totalWorkDays}天)` : "底薪 Basic"}
                      value={calc.basic.basic}
                    />
                  </div>
                  <div className="py-2">
                    <LedgerLine label="应发合计 Gross" value={calc.basic.gross} emphasis />
                  </div>
                  <div className="py-2">
                    {!actsLikeIntern && selected.epfEnabled && <LedgerLine label="EPF 公积金" value={calc.basic.epf} sign="-" />}
                    {!actsLikeIntern && selected.socsoEnabled && <LedgerLine label="SOCSO 社险" value={calc.basic.socso} sign="-" />}
                    {!actsLikeIntern && selected.eisEnabled && <LedgerLine label="EIS 就业保险" value={calc.basic.eis} sign="-" />}
                    {!actsLikeIntern && calc.basic.pcb > 0 && <LedgerLine label="PCB 所得税" value={calc.basic.pcb} sign="-" />}
                    {!actsLikeIntern && calc.basic.leaveDeduction > 0 && (
                      <LedgerLine
                        label={`缺勤扣薪 (${calc.basic.uncoveredAbsentDays}天)`}
                        value={calc.basic.leaveDeduction}
                        sign="-"
                      />
                    )}
                    {(actsLikeIntern ||
                      (!selected.epfEnabled && !selected.socsoEnabled && !selected.eisEnabled && calc.basic.pcb === 0 && calc.basic.leaveDeduction === 0)) && (
                      <div className="text-[16px] md:text-[19px] text-white/40 py-1">无扣除项</div>
                    )}
                  </div>
                    </>
                  )}
                  <div className="pt-3">
                    <div className="flex items-baseline justify-between">
                      <span className="text-[17px] md:text-[20px] text-[#C9A227] tracking-wide">
                        {payView === "total" ? "Total Net Pay" : "实发工资 Net Pay"}
                      </span>
                      <span className="font-mono text-[29px] md:text-[32px] font-extrabold text-[#C9A227]" style={{ fontFamily: "'Source Serif 4', serif" }}>
                        RM {currency(payView === "total" ? calc.totalNet : payView === "subcon" ? calc.subcon.net : calc.basic.net)}
                      </span>
                    </div>
                  </div>
                </div>
                {payView !== "basic" && (
                  <p className="text-[15px] md:text-[18px] text-white/35 mt-5 leading-relaxed">
                    EPF/SOCSO/EIS 比例为常见估算值，SOCSO 官方按工资级距表计算而非纯百分比。请对照 KWSP / PERKESO 最新表核实。员工的缺勤天数会先由年假余额抵扣，年假用完后才开始扣薪；实习生按出勤天数直接折算底薪。
                  </p>
                )}
              </div>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
