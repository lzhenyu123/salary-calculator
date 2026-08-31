import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import "./App.css";

const uid = () => Math.random().toString(36).slice(2, 10);

const emptyEmployee = (type = "fulltime") => ({
 id: uid(),
 name: "",
 title: "",
 type,
 basic: type === "intern" ? 1500 : 4000,
 totalWorkDays: 26,
 daysWorked: 26,
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
 leaveQuota: 8,
 leaveDates: [],
 subconAmount: 0,
});

function currency(n) {
 const v = Number(n) || 0;
 const sign = v < 0 ? "-" : "";
 const truncated = Math.floor(Math.abs(v) * 100 + 1e-6) / 100;
 return sign + truncated.toLocaleString("en-MY", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function computeBasicPayroll(e) {
 const actsLikeIntern = e.type === "intern";
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
 const totalDeduction = epf + socso + eis + pcb;
 const net = gross - totalDeduction;
 return { basic, gross, epf, socso, eis, epfAuto, socsoAuto, eisAuto, pcb, totalDeduction, net, isProrated, dailyRate };
}

function computePayroll(e) {
 const basicCalc = computeBasicPayroll(e);
 const subconAmount = Number(e.subconAmount) || 0;
 const totalNet = basicCalc.net + subconAmount;
 return { basic: basicCalc, subcon: subconAmount, totalNet };
}

const FieldRow = ({ label, unit, value, onChange, step = 1, min = 0, disabled }) => (
 <div className={`flex items-center justify-between py-2.5 border-b border-gray-200 ${disabled ? "opacity-40" : ""}`}>
 <span className="text-base text-gray-700">{label}</span>
 <div className="flex items-center gap-1.5">
 <input
 type="number"
 step={step}
 min={min}
 value={value}
 disabled={disabled}
 onChange={(ev) => onChange(ev.target.value)}
 className="w-28 text-right bg-transparent font-mono text-lg focus:outline-none focus:bg-blue-50 rounded px-1.5 py-0.5 transition-colors disabled:cursor-not-allowed"
 />
 <span className="text-base text-gray-500 w-10">{unit}</span>
 </div>
 </div>
);

const LedgerLine = ({ label, value, sign, emphasis }) => (
 <div className="flex items-baseline justify-between py-1.5">
 <span className={`text-base ${emphasis ? "text-white font-bold" : "text-gray-300"}`}>{label}</span>
 <span className={`font-mono text-lg ${sign === "-" ? "text-red-400" : "text-white"} ${emphasis ? "font-extrabold text-xl" : ""}`}>
 {sign === "-" ? "−" : ""}RM {currency(value)}
 </span>
 </div>
);

export default function PayrollSystem() {
 const [employees, setEmployees] = useState([]);
 const [selectedId, setSelectedId] = useState(null);
 const [loading, setLoading] = useState(true);

 useEffect(() => {
 setLoading(false);
 }, []);

 const updateEmployees = (updater) => {
 setEmployees((prev) => {
 const next = typeof updater === "function" ? updater(prev) : updater;
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

 const selected = employees.find((e) => e.id === selectedId) || null;
 const calc = selected ? computePayroll(selected) : null;
 const actsLikeIntern = selected ? selected.type === "intern" : false;
 const totalNet = employees.reduce((sum, e) => sum + computePayroll(e).totalNet, 0);

 if (loading) {
 return <div className="min-h-screen bg-gray-100 flex items-center justify-center">Loading...</div>;
 }

 return (
 <div className="min-h-screen bg-gradient-to-br from-gray-900 to-gray-800 text-white font-sans">
 <header className="border-b border-gray-700 bg-gray-900/95 backdrop-blur sticky top-0 z-10">
 <div className="max-w-6xl mx-auto px-6 py-6 flex items-center justify-between">
 <div className="flex items-baseline gap-3">
 <h1 className="text-4xl font-extrabold tracking-tight">薪资台账</h1>
 <span className="text-lg text-gray-400 font-mono">Payroll Ledger</span>
 </div>
 </div>
 </header>

 <div className="max-w-6xl mx-auto px-6 py-8 grid grid-cols-[280px_1fr] gap-8">
 <aside>
 <div className="flex items-center justify-between mb-3">
 <span className="text-sm uppercase tracking-widest text-gray-400">
 👥 员工 ({employees.length})
 </span>
 </div>

 <div className="grid grid-cols-2 gap-2 mb-3">
 <button
 onClick={() => addEmployee("fulltime")}
 className="px-3 py-2 rounded-full text-sm font-bold bg-emerald-600 text-white hover:bg-emerald-700 transition-colors"
 >
 + 员工
 </button>
 <button
 onClick={() => addEmployee("intern")}
 className="px-3 py-2 rounded-full text-sm font-bold bg-amber-600 text-white hover:bg-amber-700 transition-colors"
 >
 + 实习生
 </button>
 </div>

 <div className="space-y-1">
 {employees.length === 0 && (
 <div className="text-base text-gray-400 py-6 text-center border border-dashed border-gray-600 rounded-md">
 点击上方按钮新增
 </div>
 )}
 {employees.map((e) => {
 const c = computePayroll(e);
 const active = e.id === selectedId;
 return (
 <button
 key={e.id}
 onClick={() => setSelectedId(e.id)}
 className={`w-full text-left px-3 py-2.5 rounded-md transition-colors ${
 active ? "bg-blue-600 text-white" : "hover:bg-gray-700 text-gray-300"
 }`}
 >
 <div className="text-lg font-bold">{e.name || "未命名"}</div>
 <div className="text-sm font-mono text-yellow-400">RM {currency(c.totalNet)}</div>
 </button>
 );
 })}
 </div>

 {employees.length > 0 && (
 <div className="mt-6 pt-4 border-t border-gray-600">
 <div className="text-sm uppercase tracking-widest text-gray-400 mb-2">本期合计</div>
 <div className="font-mono text-3xl font-extrabold text-yellow-400">
 RM {currency(totalNet)}
 </div>
 </div>
 )}
 </aside>

 <main>
 {!selected ? (
 <div className="h-full min-h-[400px] flex items-center justify-center text-gray-400 text-xl border border-dashed border-gray-600 rounded-lg">
 选择或新增员工
 </div>
 ) : (
 <div className="grid grid-cols-[1fr_320px] gap-6">
 <div className="bg-gray-800/50 border border-gray-700 rounded-lg p-6">
 <div className="flex items-center justify-between mb-5">
 <input
 value={selected.name}
 onChange={(e) => patchEmployee(selected.id, { name: e.target.value })}
 placeholder="员工姓名"
 className="text-3xl font-extrabold bg-transparent focus:outline-none placeholder:text-gray-600 text-white"
 />
 <button
 onClick={() => removeEmployee(selected.id)}
 className="text-red-400 hover:text-red-300 px-2"
 >
 ✕
 </button>
 </div>

 <div className="flex gap-2 mb-5">
 <button
 onClick={() => patchEmployee(selected.id, { type: "fulltime" })}
 className={`px-3 py-1 rounded-full text-sm font-bold ${
 selected.type === "fulltime" ? "bg-emerald-600 text-white" : "bg-gray-700"
 }`}
 >
 员工
 </button>
 <button
 onClick={() => patchEmployee(selected.id, { type: "intern" })}
 className={`px-3 py-1 rounded-full text-sm font-bold ${
 selected.type === "intern" ? "bg-amber-600 text-white" : "bg-gray-700"
 }`}
 >
 实习生
 </button>
 </div>

 <div className="space-y-4">
 <div>
 <h3 className="text-sm uppercase tracking-widest text-gray-400 mb-2">💰 收入</h3>
 <FieldRow
 label="底薪 Basic"
 unit="RM"
 value={selected.basic}
 onChange={(v) => patchEmployee(selected.id, { basic: v })}
 step={50}
 />
 <FieldRow label="应上班天数" unit="天" value={selected.totalWorkDays} onChange={(v) => patchEmployee(selected.id, { totalWorkDays: v })} step={1} />
 <FieldRow label="实际出勤天数" unit="天" value={selected.daysWorked} onChange={(v) => patchEmployee(selected.id, { daysWorked: v })} step={0.5} />
 </div>

 {!actsLikeIntern && (
 <div>
 <h3 className="text-sm uppercase tracking-widest text-gray-400 mb-2">📋 法定扣除</h3>
 <FieldRow label="EPF (%)" unit="%" value={selected.epfRate} onChange={(v) => patchEmployee(selected.id, { epfRate: v })} step={0.5} />
 <FieldRow label="SOCSO (%)" unit="%" value={selected.socsoRate} onChange={(v) => patchEmployee(selected.id, { socsoRate: v })} step={0.0001} />
 <FieldRow label="EIS (%)" unit="%" value={selected.eisRate} onChange={(v) => patchEmployee(selected.id, { eisRate: v })} step={0.0001} />
 <FieldRow label="PCB 所得税" unit="RM" value={selected.pcb} onChange={(v) => patchEmployee(selected.id, { pcb: v })} step={10} />
 </div>
 )}

 <div>
 <h3 className="text-sm uppercase tracking-widest text-gray-400 mb-2">➕ 其他</h3>
 <FieldRow label="年假额度" unit="天" value={selected.leaveQuota} onChange={(v) => patchEmployee(selected.id, { leaveQuota: v })} step={1} />
 <FieldRow label="额外补款" unit="RM" value={selected.subconAmount} onChange={(v) => patchEmployee(selected.id, { subconAmount: v })} step={10} />
 </div>
 </div>
 </div>

 <div className="bg-gray-900 border border-gray-700 rounded-lg p-6 h-fit sticky top-24">
 <div className="flex items-center gap-2 mb-4 text-yellow-400">
 <span className="text-2xl">📝</span>
 <span className="text-sm uppercase tracking-widest">结算单</span>
 </div>
 <div className="divide-y divide-gray-700 space-y-2">
 <div>
 <LedgerLine
 label={calc.basic.isProrated ? `底薪 (${selected.daysWorked}/${selected.totalWorkDays}天)` : "底薪"}
 value={calc.basic.basic}
 />
 </div>
 <div className="pt-2">
 <LedgerLine label="应发合计" value={calc.basic.gross} emphasis />
 </div>
 <div className="py-2 space-y-1">
 {calc.basic.epf > 0 && <LedgerLine label="EPF 公积金" value={calc.basic.epf} sign="-" />}
 {calc.basic.socso > 0 && <LedgerLine label="SOCSO 社险" value={calc.basic.socso} sign="-" />}
 {calc.basic.eis > 0 && <LedgerLine label="EIS 就业保险" value={calc.basic.eis} sign="-" />}
 {calc.basic.pcb > 0 && <LedgerLine label="PCB 所得税" value={calc.basic.pcb} sign="-" />}
 {calc.subcon > 0 && <LedgerLine label="额外补款" value={calc.subcon} />}
 </div>
 <div className="pt-3">
 <div className="flex items-baseline justify-between">
 <span className="text-base text-yellow-400 tracking-wide">实发工资</span>
 <span className="font-mono text-4xl font-extrabold text-yellow-400">
 RM {currency(calc.totalNet)}
 </span>
 </div>
 </div>
 </div>
 </div>
 </div>
 )}
 </main>
 </div>
 </div>
 );
}

