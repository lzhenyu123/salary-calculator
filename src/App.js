import React, { useState } from "react";
import "./App.css";

const uid = () => Math.random().toString(36).slice(2, 10);

const emptyEmployee = (type = "fulltime") => ({
 id: uid(),
 name: "",
 type,
 basic: type === "intern" ? 1500 : 4000,
 totalWorkDays: 26,
 daysWorked: 26,
 epfRate: 11,
 socsoRate: 1.22,
 eisRate: 0.19,
 pcb: 0,
 leaveQuota: 8,
 subconAmount: 0,
});

function currency(n) {
 const v = Number(n) || 0;
 const sign = v < 0 ? "-" : "";
 const truncated = Math.floor(Math.abs(v) * 100 + 1e-6) / 100;
 return sign + truncated.toLocaleString("en-MY", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function computePayroll(e) {
 const rawBasic = Number(e.basic) || 0;
 const totalWorkDays = Number(e.totalWorkDays) || 0;
 const daysWorked = Number(e.daysWorked) || 0;
 const isProrated = e.type === "intern" && totalWorkDays > 0;
 const basic = isProrated ? (rawBasic / totalWorkDays) * daysWorked : rawBasic;
 const gross = basic;
 const isIntern = e.type === "intern";
 const epf = !isIntern ? gross * ((Number(e.epfRate) || 0) / 100) : 0;
 const socso = !isIntern ? gross * ((Number(e.socsoRate) || 0) / 100) : 0;
 const eis = !isIntern ? gross * ((Number(e.eisRate) || 0) / 100) : 0;
 const pcb = isIntern ? 0 : Number(e.pcb) || 0;
 const totalDeduction = epf + socso + eis + pcb;
 const net = gross - totalDeduction;
 const subconAmount = Number(e.subconAmount) || 0;
 const totalNet = net + subconAmount;
 return { basic, gross, epf, socso, eis, pcb, totalDeduction, net, subconAmount, totalNet };
}

export default function PayrollSystem() {
 const [employees, setEmployees] = useState([]);
 const [selectedId, setSelectedId] = useState(null);

 const updateEmployees = (updater) => {
 setEmployees((prev) => (typeof updater === "function" ? updater(prev) : updater));
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
 const totalNet = employees.reduce((sum, e) => sum + computePayroll(e).totalNet, 0);

 return (
 <div className="min-h-screen bg-gradient-to-br from-gray-900 to-gray-800 text-white font-sans">
 <header className="border-b border-gray-700 bg-gray-900/95 sticky top-0 z-10">
 <div className="max-w-6xl mx-auto px-6 py-6">
 <h1 className="text-4xl font-extrabold">薪资台账</h1>
 <span className="text-lg text-gray-400">Payroll Ledger</span>
 </div>
 </header>

 <div className="max-w-6xl mx-auto px-6 py-8 grid grid-cols-[280px_1fr] gap-8">
 <aside>
 <div className="text-sm uppercase tracking-widest text-gray-400 mb-3">
 👥 员工 ({employees.length})
 </div>

 <div className="grid grid-cols-2 gap-2 mb-4">
 <button
 onClick={() => addEmployee("fulltime")}
 className="px-3 py-2 rounded-lg text-sm font-bold bg-emerald-600 hover:bg-emerald-700"
 >
 + 员工
 </button>
 <button
 onClick={() => addEmployee("intern")}
 className="px-3 py-2 rounded-lg text-sm font-bold bg-amber-600 hover:bg-amber-700"
 >
 + 实习生
 </button>
 </div>

 <div className="space-y-2">
 {employees.length === 0 && (
 <div className="text-gray-400 py-6 text-center text-sm border border-dashed border-gray-600 rounded">
 点击按钮新增员工
 </div>
 )}
 {employees.map((e) => {
 const c = computePayroll(e);
 const active = e.id === selectedId;
 return (
 <button
 key={e.id}
 onClick={() => setSelectedId(e.id)}
 className={`w-full text-left px-3 py-2.5 rounded transition-colors ${
 active ? "bg-blue-600" : "hover:bg-gray-700"
 }`}
 >
 <div className="font-bold">{e.name || "未命名"}</div>
 <div className="text-sm text-yellow-400">RM {currency(c.totalNet)}</div>
 </button>
 );
 })}
 </div>

 {employees.length > 0 && (
 <div className="mt-6 pt-4 border-t border-gray-600">
 <div className="text-sm text-gray-400 mb-2">本期合计</div>
 <div className="text-3xl font-extrabold text-yellow-400">
 RM {currency(totalNet)}
 </div>
 </div>
 )}
 </aside>

 <main>
 {!selected ? (
 <div className="flex items-center justify-center text-gray-400 text-lg border border-dashed border-gray-600 rounded-lg h-96">
 选择或新增员工开始
 </div>
 ) : (
 <div className="grid grid-cols-[1fr_320px] gap-6">
 <div className="bg-gray-800/50 border border-gray-700 rounded-lg p-6">
 <div className="flex justify-between items-center mb-5">
 <input
 value={selected.name}
 onChange={(e) => patchEmployee(selected.id, { name: e.target.value })}
 placeholder="员工姓名"
 className="text-3xl font-bold bg-transparent focus:outline-none text-white placeholder:text-gray-600"
 />
 <button onClick={() => removeEmployee(selected.id)} className="text-red-400 text-2xl">
 ✕
 </button>
 </div>

 <div className="flex gap-2 mb-5">
 <button
 onClick={() => patchEmployee(selected.id, { type: "fulltime" })}
 className={`px-3 py-1 rounded text-sm font-bold ${
 selected.type === "fulltime" ? "bg-emerald-600" : "bg-gray-700"
 }`}
 >
 员工
 </button>
 <button
 onClick={() => patchEmployee(selected.id, { type: "intern" })}
 className={`px-3 py-1 rounded text-sm font-bold ${
 selected.type === "intern" ? "bg-amber-600" : "bg-gray-700"
 }`}
 >
 实习生
 </button>
 </div>

 <div className="space-y-6">
 <div>
 <h3 className="text-sm text-gray-400 mb-3">💰 收入</h3>
 <div className="space-y-2">
 <InputRow label="底薪" unit="RM" value={selected.basic} onChange={(v) => patchEmployee(selected.id, { basic: v })} />
 <InputRow label="应上班天数" unit="天" value={selected.totalWorkDays} onChange={(v) => patchEmployee(selected.id, { totalWorkDays: v })} />
 <InputRow label="实际出勤" unit="天" value={selected.daysWorked} onChange={(v) => patchEmployee(selected.id, { daysWorked: v })} />
 </div>
 </div>

 {selected.type === "fulltime" && (
 <div>
 <h3 className="text-sm text-gray-400 mb-3">📋 扣除</h3>
 <div className="space-y-2">
 <InputRow label="EPF %" unit="%" value={selected.epfRate} onChange={(v) => patchEmployee(selected.id, { epfRate: v })} />
 <InputRow label="SOCSO %" unit="%" value={selected.socsoRate} onChange={(v) => patchEmployee(selected.id, { socsoRate: v })} />
 <InputRow label="EIS %" unit="%" value={selected.eisRate} onChange={(v) => patchEmployee(selected.id, { eisRate: v })} />
 <InputRow label="PCB 所得税" unit="RM" value={selected.pcb} onChange={(v) => patchEmployee(selected.id, { pcb: v })} />
 </div>
 </div>
 )}

 <div>
 <h3 className="text-sm text-gray-400 mb-3">➕ 其他</h3>
 <InputRow label="额外补款" unit="RM" value={selected.subconAmount} onChange={(v) => patchEmployee(selected.id, { subconAmount: v })} />
 </div>
 </div>
 </div>

 <div className="bg-gray-900 border border-gray-700 rounded-lg p-6 h-fit sticky top-24">
 <div className="text-sm text-yellow-400 mb-4 font-bold">📝 结算单</div>
 <div className="space-y-2 divide-y divide-gray-700">
 <div className="pb-2">
 <PayLine label="底薪" value={calc.basic} />
 </div>
 <div className="py-2">
 <PayLine label="应发合计" value={calc.gross} emphasis />
 </div>
 <div className="py-2 space-y-1">
 {calc.epf > 0 && <PayLine label="EPF" value={calc.epf} minus />}
 {calc.socso > 0 && <PayLine label="SOCSO" value={calc.socso} minus />}
 {calc.eis > 0 && <PayLine label="EIS" value={calc.eis} minus />}
 {calc.pcb > 0 && <PayLine label="PCB" value={calc.pcb} minus />}
 {calc.subconAmount > 0 && <PayLine label="补款" value={calc.subconAmount} />}
 </div>
 <div className="pt-3">
 <div className="flex justify-between items-center">
 <span className="text-yellow-400 text-sm">实发工资</span>
 <span className="text-3xl font-extrabold text-yellow-400">RM {currency(calc.totalNet)}</span>
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

const InputRow = ({ label, unit, value, onChange }) => (
 <div className="flex justify-between items-center py-2 border-b border-gray-700">
 <span className="text-gray-300">{label}</span>
 <div className="flex items-center gap-1">
 <input
 type="number"
 value={value}
 onChange={(e) => onChange(e.target.value)}
 className="w-24 text-right bg-gray-700 rounded px-2 py-1 text-sm text-white"
 />
 <span className="text-xs text-gray-400 w-8">{unit}</span>
 </div>
 </div>
);

const PayLine = ({ label, value, emphasis, minus }) => (
 <div className="flex justify-between text-sm">
 <span className={minus ? "text-gray-400" : emphasis ? "text-white font-bold" : "text-gray-300"}>{label}</span>
 <span className={`font-mono ${minus ? "text-red-400" : emphasis ? "text-yellow-400 font-bold" : "text-white"}`}>
 {minus ? "−" : ""} RM {currency(value)}
 </span>
 </div>
);
