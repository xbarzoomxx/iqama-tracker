import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

const STORAGE_KEY = "iqama_tracker_v3";

const INITIAL_DATA = [{"id":2161651126,"name":"JOYNUL ABEDIN - - ABDUL MONAF","nationality":"بنجلاديش","iqamaNumber":"2161651126","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-18","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل صيانة أجهزة كهربائية","jobTitle":"عامل صيانة أجهزة كهربائية","passportNumber":"EM0468275","outsideKingdom":"نعم","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2168878490,"name":"عمار حسن علي الحسن","nationality":"السودان","iqamaNumber":"2168878490","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-26","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مدير عام","jobTitle":"مدير عام","passportNumber":"P08151865","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2499287072,"name":"ليلى عمار حسن علي","nationality":"السودان","iqamaNumber":"2499287072","type":"مرافق","relation":"بنت","gender":"أنثى","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"P08179155","outsideKingdom":"لا","familyHeadId":"2168878490","employerId":"7006539477","company":"انجال المشاعر"},{"id":2374056725,"name":"محمد عمار حسن الحسن","nationality":"السودان","iqamaNumber":"2374056725","type":"مرافق","relation":"ابن","gender":"ذكر","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"P08228098","outsideKingdom":"لا","familyHeadId":"2168878490","employerId":"7006539477","company":"انجال المشاعر"},{"id":2570076238,"name":"يزن عمار حسن علي","nationality":"السودان","iqamaNumber":"2570076238","type":"مرافق","relation":"ابن","gender":"ذكر","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"P11918293","outsideKingdom":"لا","familyHeadId":"2168878490","employerId":"7006539477","company":"انجال المشاعر"},{"id":2374056485,"name":"هسه عمار حسن الحسن","nationality":"السودان","iqamaNumber":"2374056485","type":"مرافق","relation":"بنت","gender":"أنثى","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"P08153195","outsideKingdom":"لا","familyHeadId":"2168878490","employerId":"7006539477","company":"انجال المشاعر"},{"id":2312831726,"name":"هبه محمداحمد محمد صديق","nationality":"السودان","iqamaNumber":"2312831726","type":"مرافق","relation":"زوجة","gender":"أنثى","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"P08151939","outsideKingdom":"لا","familyHeadId":"2168878490","employerId":"7006539477","company":"انجال المشاعر"},{"id":2207144896,"name":"ZAFOR ULLAH ABDUL LATIB","nationality":"بنجلاديش","iqamaNumber":"2207144896","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-08","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني هندسة كهربائية","jobTitle":"فني هندسة كهربائية","passportNumber":"EM0405033","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2253325704,"name":"MOHAMED SHAHALAM MOHAMED SHAMSULHAQUE","nationality":"بنجلاديش","iqamaNumber":"2253325704","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-10-09","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني صيانة ميكانيكية","jobTitle":"فني صيانة ميكانيكية","passportNumber":"EH0582394","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2270598135,"name":"حماده منصور محمد احمد","nationality":"مصر","iqamaNumber":"2270598135","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-12-06","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني ميكانيكي تمديدات صحية","jobTitle":"فني ميكانيكي تمديدات صحية","passportNumber":"A28507403","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2285035610,"name":"حازم حشمت توفيق البسطويسي","nationality":"مصر","iqamaNumber":"2285035610","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-27","renewalStatus":"لم يبدأ","renewalCost":"","notes":"محاسب","jobTitle":"محاسب","passportNumber":"A29095351","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2316800917,"name":"محمد فتحي علي الفقي","nationality":"مصر","iqamaNumber":"2316800917","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-23","renewalStatus":"لم يبدأ","renewalCost":"","notes":"لحام","jobTitle":"لحام","passportNumber":"A04219484","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2368280851,"name":"عمار فضل المولي محمد الزبير","nationality":"السودان","iqamaNumber":"2368280851","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-02","renewalStatus":"لم يبدأ","renewalCost":"","notes":"سباك","jobTitle":"سباك","passportNumber":"P12124172","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2377607409,"name":"احمد حمدى احمد هلال","nationality":"مصر","iqamaNumber":"2377607409","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-11-19","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مهندس  ميكانيكي","jobTitle":"مهندس  ميكانيكي","passportNumber":"A28249488","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2391783970,"name":"بلال احمد حمدي هلال","nationality":"مصر","iqamaNumber":"2391783970","type":"مرافق","relation":"ابن","gender":"ذكر","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"A28249489","outsideKingdom":"نعم","familyHeadId":"2377607409","employerId":"7006539477","company":"انجال المشاعر"},{"id":2422905527,"name":"اروى احمد حمدي هلال","nationality":"مصر","iqamaNumber":"2422905527","type":"مرافق","relation":"بنت","gender":"أنثى","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"A32896771","outsideKingdom":"نعم","familyHeadId":"2377607409","employerId":"7006539477","company":"انجال المشاعر"},{"id":2545044733,"name":"زيد احمد حمدي هلال","nationality":"مصر","iqamaNumber":"2545044733","type":"مرافق","relation":"ابن","gender":"ذكر","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"A32896837","outsideKingdom":"نعم","familyHeadId":"2377607409","employerId":"7006539477","company":"انجال المشاعر"},{"id":2391783988,"name":"ساره محمود خليل محمد","nationality":"مصر","iqamaNumber":"2391783988","type":"مرافق","relation":"زوجة","gender":"أنثى","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"A28249643","outsideKingdom":"نعم","familyHeadId":"2377607409","employerId":"7006539477","company":"انجال المشاعر"},{"id":2383433394,"name":"NAFEES AHMAD MOHD MUSTFA","nationality":"الهند","iqamaNumber":"2383433394","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-28","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني كهربائي أنظمةحمايةكهربائية","jobTitle":"فني كهربائي أنظمةحمايةكهربائية","passportNumber":"V5262084","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2408364129,"name":"احمد جمال البسيوني ابراهيم","nationality":"مصر","iqamaNumber":"2408364129","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-04","renewalStatus":"لم يبدأ","renewalCost":"","notes":"سباك","jobTitle":"سباك","passportNumber":"A23626418","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2424076574,"name":"محمد عبدالله عبدالحليم عبدالله","nationality":"السودان","iqamaNumber":"2424076574","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-02","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مهندس  الكترونيات","jobTitle":"مهندس  الكترونيات","passportNumber":"P09775301","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2426377111,"name":"GAYYUR AHMAD MURSLEEN AHMAD","nationality":"الهند","iqamaNumber":"2426377111","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-20","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني كهربائي تمديدات كهربائية","jobTitle":"فني كهربائي تمديدات كهربائية","passportNumber":"U0688192","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2438640225,"name":"MOHAMMAD RAJU MOJUMDAR HAFEJA","nationality":"بنجلاديش","iqamaNumber":"2438640225","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2024-11-07","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل بناء","jobTitle":"عامل بناء","passportNumber":"EJ0383352","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2444383471,"name":"عبدالعليم قائد ناصر قاسم","nationality":"اليمن","iqamaNumber":"2444383471","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-09-25","renewalStatus":"لم يبدأ","renewalCost":"","notes":"سباك","jobTitle":"سباك","passportNumber":"10928633","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2464717319,"name":"MOHAMMED JOYNUL ABEDIN","nationality":"بنجلاديش","iqamaNumber":"2464717319","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-18","renewalStatus":"لم يبدأ","renewalCost":"","notes":"ميكانيكي مصاعد","jobTitle":"ميكانيكي مصاعد","passportNumber":"EH0762631","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2466025505,"name":"MANJOOR AHMED HAFIZ FAZRUDDIN","nationality":"الهند","iqamaNumber":"2466025505","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-26","renewalStatus":"لم يبدأ","renewalCost":"","notes":"لحام","jobTitle":"لحام","passportNumber":"Y6166044","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2479489730,"name":"FIROZ AHMED SAEED HUSSAIN","nationality":"الهند","iqamaNumber":"2479489730","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-19","renewalStatus":"لم يبدأ","renewalCost":"","notes":"سباك","jobTitle":"سباك","passportNumber":"T3702870","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2495781763,"name":"غسان على جلال احمد","nationality":"السودان","iqamaNumber":"2495781763","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2025-10-01","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل صيانة أجهزة كهربائية","jobTitle":"عامل صيانة أجهزة كهربائية","passportNumber":"P07034750","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2508579659,"name":"FARHAN MANNAN ABDUL MANNAN","nationality":"باكستان","iqamaNumber":"2508579659","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-05-08","renewalStatus":"لم يبدأ","renewalCost":"","notes":"ميكانيكي مصاعد","jobTitle":"ميكانيكي مصاعد","passportNumber":"AR0704682","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2508580137,"name":"HABIBUR RAHMAN","nationality":"بنجلاديش","iqamaNumber":"2508580137","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-24","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني هندسة الكترونية","jobTitle":"فني هندسة الكترونية","passportNumber":"A18031168","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2510535806,"name":"اسد الله محمد مساعد سليمان","nationality":"السودان","iqamaNumber":"2510535806","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-14","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني صيانة أجهزة الكترونية","jobTitle":"فني صيانة أجهزة الكترونية","passportNumber":"P05665810","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2513490108,"name":"MD FOQRUL ISLAM","nationality":"بنجلاديش","iqamaNumber":"2513490108","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2024-02-20","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل صيانة أجهزة كهربائية","jobTitle":"عامل صيانة أجهزة كهربائية","passportNumber":"EN0083184","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2515822423,"name":"IRSHAD AHMED","nationality":"الهند","iqamaNumber":"2515822423","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-16","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني كهربائي تمديدات كهربائية","jobTitle":"فني كهربائي تمديدات كهربائية","passportNumber":"U0136451","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2525086464,"name":"ABDULLAH MUHAMMAD ASLAM KHAN","nationality":"باكستان","iqamaNumber":"2525086464","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-10","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني هندسة كهربائية","jobTitle":"فني هندسة كهربائية","passportNumber":"HB1073993","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2526249459,"name":"ابوبكر حبيب الله محمد فضل الله","nationality":"السودان","iqamaNumber":"2526249459","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2024-09-04","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مهندس  كهربائي","jobTitle":"مهندس  كهربائي","passportNumber":"P08667135","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2526824905,"name":"BABU MANJUR AHMED","nationality":"الهند","iqamaNumber":"2526824905","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-11","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني صيانة ميكانيكية","jobTitle":"فني صيانة ميكانيكية","passportNumber":"U2024984","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2549937999,"name":"MOHAMMED ELIAS","nationality":"بنجلاديش","iqamaNumber":"2549937999","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-11","renewalStatus":"لم يبدأ","renewalCost":"","notes":"سباك","jobTitle":"سباك","passportNumber":"EG0771953","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2566347197,"name":"محمد احمد محمد قسم السيد","nationality":"السودان","iqamaNumber":"2566347197","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-30","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مدير مالي","jobTitle":"مدير مالي","passportNumber":"P09336349","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2567584533,"name":"حسن عطا المنان الحسن محمد","nationality":"السودان","iqamaNumber":"2567584533","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-20","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل بناء","jobTitle":"عامل بناء","passportNumber":"P11068420","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2571845128,"name":"المصطفي محمد علي الضو","nationality":"السودان","iqamaNumber":"2571845128","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-21","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني كهربائي تمديدات كهربائية","jobTitle":"فني كهربائي تمديدات كهربائية","passportNumber":"P10162929","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2575431966,"name":"ايات جاد الله محمد احمد البلوله","nationality":"السودان","iqamaNumber":"2575431966","type":"موظف","relation":"","gender":"أنثى","expiryDate":"2026-06-12","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مهندس  ميكانيكي","jobTitle":"مهندس  ميكانيكي","passportNumber":"P13192846","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2578801967,"name":"MD JONY MIA","nationality":"بنجلاديش","iqamaNumber":"2578801967","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2027-07-10","renewalStatus":"لم يبدأ","renewalCost":"","notes":"ميكانيكي معدات الكرتونية","jobTitle":"ميكانيكي معدات الكرتونية","passportNumber":"A11176721","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2590187510,"name":"محمد سر الختم عوض محمد","nationality":"السودان","iqamaNumber":"2590187510","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-18","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني كهربائي تمديدات كهربائية","jobTitle":"فني كهربائي تمديدات كهربائية","passportNumber":"P11400698","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2628316743,"name":"عبدالحليم محمد عثمان سليمان","nationality":"السودان","iqamaNumber":"2628316743","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-26","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل صيانة أجهزة كهربائية","jobTitle":"عامل صيانة أجهزة كهربائية","passportNumber":"P13512941","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"}];

const STATUS_COLORS = {
  منتهية:        { bg:"#fee2e2", text:"#dc2626", border:"#fca5a5" },
  "تنتهي قريباً":{ bg:"#fef3c7", text:"#d97706", border:"#fcd34d" },
  سارية:         { bg:"#dcfce7", text:"#16a34a", border:"#86efac" },
  "قيد التجديد": { bg:"#dbeafe", text:"#2563eb", border:"#93c5fd" },
  مرافق:         { bg:"#f3e8ff", text:"#7c3aed", border:"#c4b5fd" },
};

const COMPANY_COLORS = {
  "انجال المشاعر": { bg:"#fdf0f0", text:"#6B1A1A", border:"#e8b4b4" },
  "دلتا الماسية":  { bg:"#fffbeb", text:"#b45309", border:"#fcd34d" },
};

const RELATION_ICONS = { زوجة:"💑", ابن:"👦", بنت:"👧" };

function getStatus(r) {
  if (r.type === "مرافق") return "مرافق";
  if (r.renewalStatus === "قيد التجديد") return "قيد التجديد";
  if (!r.expiryDate) return "مرافق";
  const days = getDaysLeft(r.expiryDate);
  if (days < 0) return "منتهية";
  if (days <= 30) return "تنتهي قريباً";
  return "سارية";
}

function getDaysLeft(d) {
  if (!d) return 9999;
  return Math.ceil((new Date(d) - new Date()) / 86400000);
}

function exportToExcel(records) {
  const employees = records.filter(r => r.type !== "مرافق");
  const dependents = records.filter(r => r.type === "مرافق");

  const makeRow = r => ({
    "رقم الإقامة": r.iqamaNumber,
    "الاسم": r.name,
    "الجنسية": r.nationality||"-",
    "الجنس": r.gender||"-",
    "صلة القرابة": r.relation||"-",
    "رب الأسرة": r.familyHeadId||"-",
    "المهنة": r.jobTitle||"-",
    "رقم الجواز": r.passportNumber||"-",
    "تاريخ الانتهاء": r.expiryDate||"-",
    "الأيام المتبقية": r.expiryDate ? getDaysLeft(r.expiryDate) : "-",
    "الحالة": getStatus(r),
    "حالة التجديد": r.renewalStatus,
    "خارج المملكة": r.outsideKingdom||"-",
    "الشركة": r.company||"-",
  });

  const ws1 = XLSX.utils.json_to_sheet(employees.map(makeRow));
  ws1["!cols"] = [{wch:16},{wch:28},{wch:12},{wch:8},{wch:10},{wch:14},{wch:26},{wch:14},{wch:14},{wch:12},{wch:16},{wch:14},{wch:12},{wch:16}];
  const ws2 = XLSX.utils.json_to_sheet(dependents.map(makeRow));
  ws2["!cols"] = ws1["!cols"];

  const statsData = [
    {"البيان":"إجمالي السجلات","العدد":records.length},
    {"البيان":"موظفون","العدد":employees.length},
    {"البيان":"مرافقون","العدد":dependents.length},
    {"البيان":"إقامات سارية","العدد":records.filter(r=>getStatus(r)==="سارية").length},
    {"البيان":"تنتهي قريباً","العدد":records.filter(r=>getStatus(r)==="تنتهي قريباً").length},
    {"البيان":"منتهية","العدد":records.filter(r=>getStatus(r)==="منتهية").length},
    {"البيان":"انجال المشاعر","العدد":records.filter(r=>r.company==="انجال المشاعر").length},
    {"البيان":"دلتا الماسية","العدد":records.filter(r=>r.company==="دلتا الماسية").length},
  ];
  const ws3 = XLSX.utils.json_to_sheet(statsData);
  ws3["!cols"] = [{wch:28},{wch:12}];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, "الموظفون");
  XLSX.utils.book_append_sheet(wb, ws2, "المرافقون");
  XLSX.utils.book_append_sheet(wb, ws3, "ملخص إحصائي");
  XLSX.writeFile(wb, `تقرير_الإقامات_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function exportToPDF(records) {
  const today = new Date().toLocaleDateString("ar-SA",{year:"numeric",month:"long",day:"numeric"});
  const employees = records.filter(r=>r.type!=="مرافق");
  const dependents = records.filter(r=>r.type==="مرافق");

  const makeRows = (list, showRelation=false) => list
    .sort((a,b) => getDaysLeft(a.expiryDate)-getDaysLeft(b.expiryDate))
    .map((r,i) => {
      const st = getStatus(r);
      const sc = STATUS_COLORS[st];
      const days = r.expiryDate ? getDaysLeft(r.expiryDate) : null;
      const cc = COMPANY_COLORS[r.company]||{bg:"#f9f9f9",text:"#374151",border:"#e5e7eb"};
      return `<tr style="background:${i%2===0?"#f9fafb":"#fff"}">
        <td>${i+1}</td>
        <td><strong>${r.name}</strong></td>
        <td>${r.nationality||"-"}</td>
        <td>${r.iqamaNumber}</td>
        ${showRelation ? `<td>${r.relation||"-"}</td><td style="font-size:11px">${r.familyHeadId||"-"}</td>` : `<td>${r.jobTitle||"-"}</td><td>${r.passportNumber||"-"}</td>`}
        <td>${r.expiryDate ? new Date(r.expiryDate).toLocaleDateString("ar-SA") : "-"}</td>
        <td style="font-weight:700;color:${sc.text}">${days===null?"-":days<0?`منتهية منذ ${Math.abs(days)} ي`:`${days} يوم`}</td>
        <td><span style="background:${sc.bg};color:${sc.text};border:1px solid ${sc.border};padding:2px 7px;border-radius:10px;font-size:10px;font-weight:600">${st}</span></td>
        <td><span style="background:${cc.bg};color:${cc.text};padding:2px 7px;border-radius:8px;font-size:10px">${r.company||"-"}</span></td>
      </tr>`;
    }).join("");

  const html = `<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="UTF-8"><title>تقرير الإقامات</title>
  <style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Segoe UI',Tahoma,sans-serif;direction:rtl;color:#1f2937;padding:18px;font-size:11px}
  .header{background:linear-gradient(135deg,#1e3a5f,#2563eb);color:#fff;padding:16px 22px;border-radius:10px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center}
  .header h1{font-size:17px;font-weight:800}.stats{display:grid;grid-template-columns:repeat(8,1fr);gap:8px;margin-bottom:16px}
  .stat{background:#f8faff;border-radius:8px;padding:8px;text-align:center;border-top:3px solid #2563eb}
  .stat-num{font-size:18px;font-weight:800;color:#1e3a5f}.stat-label{font-size:9px;color:#6b7280;margin-top:1px}
  h2{font-size:13px;color:#1e3a5f;margin:14px 0 8px;padding-right:8px;border-right:4px solid #2563eb}
  table{width:100%;border-collapse:collapse;font-size:10px;margin-bottom:14px}thead tr{background:#1e3a5f;color:#fff}
  th{padding:7px 5px;text-align:center;font-weight:700;font-size:10px}td{padding:6px 5px;text-align:center;border-bottom:1px solid #e5e7eb}
  .footer{text-align:center;font-size:9px;color:#9ca3af;margin-top:8px}@media print{body{padding:10px}}</style></head><body>
  <div class="header">
    <div><h1>🪪 تقرير متابعة الإقامات الشامل</h1><p style="opacity:.85;margin-top:3px;font-size:11px">تاريخ التقرير: ${today}</p></div>
    <div style="font-size:12px;text-align:left">إجمالي السجلات: <strong>${records.length}</strong><br>موظفون: <strong>${employees.length}</strong> | مرافقون: <strong>${dependents.length}</strong></div>
  </div>
  <div class="stats">
    <div class="stat" style="border-top-color:#16a34a"><div class="stat-num" style="color:#16a34a">${records.filter(r=>getStatus(r)==="سارية").length}</div><div class="stat-label">✅ سارية</div></div>
    <div class="stat" style="border-top-color:#d97706"><div class="stat-num" style="color:#d97706">${records.filter(r=>getStatus(r)==="تنتهي قريباً").length}</div><div class="stat-label">⚠️ تنتهي قريباً</div></div>
    <div class="stat" style="border-top-color:#dc2626"><div class="stat-num" style="color:#dc2626">${records.filter(r=>getStatus(r)==="منتهية").length}</div><div class="stat-label">❌ منتهية</div></div>
    <div class="stat" style="border-top-color:#7c3aed"><div class="stat-num" style="color:#7c3aed">${dependents.length}</div><div class="stat-label">👨‍👩‍👧 مرافقون</div></div>
    <div class="stat"><div class="stat-num">${employees.length}</div><div class="stat-label">👤 موظفون</div></div>
    <div class="stat" style="border-top-color:#6B1A1A"><div class="stat-num" style="color:#6B1A1A">${records.filter(r=>r.company==="انجال المشاعر").length}</div><div class="stat-label">🏢 انجال</div></div>
    <div class="stat" style="border-top-color:#b45309"><div class="stat-num" style="color:#b45309">${records.filter(r=>r.company==="دلتا الماسية").length}</div><div class="stat-label">🏢 دلتا</div></div>
    <div class="stat"><div class="stat-num">${records.length}</div><div class="stat-label">📋 الإجمالي</div></div>
  </div>
  <h2>👤 الموظفون (${employees.length} موظف)</h2>
  <table><thead><tr><th>#</th><th>الاسم</th><th>الجنسية</th><th>رقم الإقامة</th><th>المهنة</th><th>رقم الجواز</th><th>تاريخ الانتهاء</th><th>الأيام المتبقية</th><th>الحالة</th><th>الشركة</th></tr></thead>
  <tbody>${makeRows(employees,false)}</tbody></table>
  <h2>👨‍👩‍👧 المرافقون (${dependents.length} مرافق)</h2>
  <table><thead><tr><th>#</th><th>الاسم</th><th>الجنسية</th><th>رقم الإقامة</th><th>صلة القرابة</th><th>رقم رب الأسرة</th><th>تاريخ الانتهاء</th><th>الأيام المتبقية</th><th>الحالة</th><th>الشركة</th></tr></thead>
  <tbody>${makeRows(dependents,true)}</tbody></table>
  <div class="footer">نظام متابعة الإقامات — ${today}</div>
  <script>window.onload=()=>window.print();</script></body></html>`;

  const win = window.open(URL.createObjectURL(new Blob([html],{type:"text/html;charset=utf-8"})),"_blank");
  if(!win) alert("يرجى السماح بفتح النوافذ المنبثقة");
}

const emptyForm = { name:"",nationality:"",iqamaNumber:"",type:"موظف",relation:"",gender:"ذكر",expiryDate:"",renewalStatus:"لم يبدأ",renewalCost:"",notes:"",jobTitle:"",passportNumber:"",outsideKingdom:"لا",familyHeadId:"",company:"انجال المشاعر" };

export default function App() {
  const [records, setRecords] = useState([]);
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState(null);
  const [form, setForm] = useState(emptyForm);
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("الكل");
  const [filterCompany, setFilterCompany] = useState("الكل");
  const [filterType, setFilterType] = useState("الكل");
  const [filterNationality, setFilterNationality] = useState("الكل");
  const [sortBy, setSortBy] = useState("expiryDate");
  const [activeTab, setActiveTab] = useState("list");
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [expandedId, setExpandedId] = useState(null);
  const [calcMonths, setCalcMonths] = useState(12);
  const [calcTarget, setCalcTarget] = useState("all");
  const [calcSelectedIds, setCalcSelectedIds] = useState(new Set()); // Set فارغ = الكل محدد
  const [calcIncludeDeps, setCalcIncludeDeps] = useState({});
  const [calcStatusFilter, setCalcStatusFilter] = useState("all"); // all | expired | soon | valid
  const [darkMode, setDarkMode] = useState(false);
  const [modalCard, setModalCard] = useState(null); // null | {label, filter, fKey, color, icon}
  const LOGO = 'data:image/png;base64,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'; // {empId: true/false}

  useEffect(() => {
    try { const s=localStorage.getItem(STORAGE_KEY); if(s){setRecords(JSON.parse(s));return;} } catch{}
    setRecords(INITIAL_DATA);
  }, []);

  useEffect(() => { if(records.length>0) localStorage.setItem(STORAGE_KEY, JSON.stringify(records)); }, [records]);

  const nationalities = ["الكل",...Array.from(new Set(records.map(r=>r.nationality).filter(Boolean))).sort()];

  const handleSubmit = () => {
    if(!form.name||!form.iqamaNumber){ alert("الاسم ورقم الإقامة إلزاميان"); return; }
    if(editId!==null){ setRecords(records.map(r=>r.id===editId?{...form,id:editId}:r)); setEditId(null); }
    else setRecords([...records,{...form,id:Date.now()}]);
    setForm(emptyForm); setShowForm(false);
  };

  const handleEdit=(r)=>{setForm(r);setEditId(r.id);setShowForm(true);setActiveTab("list");};
  const handleDelete=(id)=>{if(confirm("حذف هذا السجل؟"))setRecords(records.filter(r=>r.id!==id));};

  const filtered = records.filter(r=>{
    const st=getStatus(r);
    return (filterStatus==="الكل"||st===filterStatus)
      &&(filterCompany==="الكل"||r.company===filterCompany)
      &&(filterType==="الكل"||(filterType==="موظف"&&r.type!=="مرافق")||(filterType==="مرافق"&&r.type==="مرافق"))
      &&(filterNationality==="الكل"||r.nationality===filterNationality)
      &&(r.name.includes(search)||r.iqamaNumber.includes(search)||(r.nationality||"").includes(search)||(r.jobTitle||"").includes(search)||(r.company||"").includes(search));
  }).sort((a,b)=>{
    if(sortBy==="expiryDate") return getDaysLeft(a.expiryDate)-getDaysLeft(b.expiryDate);
    if(sortBy==="name") return a.name.localeCompare(b.name,"ar");
    if(sortBy==="company") return (a.company||"").localeCompare(b.company||"","ar");
    return 0;
  });

  const stats = {
    total: records.length,
    employees: records.filter(r=>r.type!=="مرافق").length,
    dependents: records.filter(r=>r.type==="مرافق").length,
    expired: records.filter(r=>getStatus(r)==="منتهية").length,
    soon: records.filter(r=>getStatus(r)==="تنتهي قريباً").length,
    valid: records.filter(r=>getStatus(r)==="سارية").length,
    anjal: records.filter(r=>r.company==="انجال المشاعر").length,
    delta: records.filter(r=>r.company==="دلتا الماسية").length,
  };

  const inp={padding:"9px 12px",border:`1px solid ${darkMode?"#2e3340":"#d1d5db"}`,borderRadius:"8px",fontSize:"13px",width:"100%",boxSizing:"border-box",fontFamily:"inherit",direction:"rtl",background:darkMode?"#1c1f26":"#fff",color:darkMode?"#e8eaf0":"#1f2937",outline:"none"};

  return (
    <div style={{minHeight:"100vh",background:darkMode?"#111318":"#f0f4f8",fontFamily:"'Segoe UI',Tahoma,sans-serif",direction:"rtl",colorScheme:darkMode?"dark":"light"}}>
      {/* Header */}
      <div style={{background:darkMode?"linear-gradient(135deg,#1a0a00 0%,#3d1000 100%)":"linear-gradient(135deg,#6B1A1A 0%,#8B2500 50%,#F5A800 100%)",color:"#fff",padding:"16px 24px",boxShadow:"0 4px 16px rgba(0,0,0,0.25)"}}>
        <div style={{maxWidth:1200,margin:"0 auto",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
          <div style={{display:"flex",alignItems:"center",gap:16}}>
            <img src={LOGO} alt="شعار أنجال المشاعر" style={{height:58,width:58,objectFit:"contain",borderRadius:8,background:"rgba(255,255,255,0.92)",padding:4}}/>
            <div>
              <h1 style={{margin:0,fontSize:20,fontWeight:800,letterSpacing:0.5,textShadow:"0 1px 4px rgba(0,0,0,0.3)"}}>🪪 نظام متابعة الإقامات</h1>
              <div style={{fontSize:11,opacity:.85,marginTop:2,fontWeight:600}}>شركة أنجال المشاعر · سلامة · مصاعد · كاميرات</div>
              <div style={{display:"flex",gap:12,marginTop:4,fontSize:11,opacity:.9,flexWrap:"wrap"}}>
                <span>👤 {stats.employees} موظف</span>
                <span>👨‍👩‍👧 {stats.dependents} مرافق</span>
                <span style={{background:"rgba(255,255,255,0.15)",padding:"1px 8px",borderRadius:8}}>🏢 انجال: {stats.anjal}</span>
                <span style={{background:"rgba(255,255,255,0.15)",padding:"1px 8px",borderRadius:8}}>🏢 دلتا: {stats.delta}</span>
              </div>
            </div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <button onClick={()=>setDarkMode(!darkMode)}
              title={darkMode?"وضع النهار":"الوضع المظلم"}
              style={{background:darkMode?"#F5A800":"rgba(255,255,255,0.15)",color:darkMode?"#1a0a00":"#fff",border:`1.5px solid ${darkMode?"#F5A800":"rgba(255,255,255,0.4)"}`,borderRadius:9,padding:"8px 14px",fontWeight:700,fontSize:16,cursor:"pointer",transition:"all 0.3s"}}>
              {darkMode?"☀️":"🌙"}
            </button>
            <div style={{position:"relative"}}>
              <button onClick={(e)=>{e.stopPropagation();setShowExportMenu(!showExportMenu);}} disabled={records.length===0}
                style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1.5px solid rgba(255,255,255,0.5)",borderRadius:9,padding:"8px 14px",fontWeight:600,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",gap:5}}>
                📤 تصدير ▾
              </button>
              {showExportMenu&&(
                <div style={{position:"absolute",top:"110%",left:0,background:darkMode?"#1c1f26":"#fff",borderRadius:10,boxShadow:darkMode?"0 8px 32px rgba(0,0,0,0.7)":"0 8px 24px rgba(0,0,0,0.15)",overflow:"hidden",minWidth:185,zIndex:100}}>
                  <button onClick={()=>{exportToExcel(records);setShowExportMenu(false);}}
                    style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"11px 15px",border:"none",background:"none",cursor:"pointer",fontSize:13,fontFamily:"inherit",textAlign:"right"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#f0fdf4"} onMouseLeave={e=>e.currentTarget.style.background="none"}>
                    <span style={{fontSize:17}}>📊</span><div><div style={{fontWeight:700,color:"#16a34a"}}>تصدير Excel</div><div style={{fontSize:11,color:"#6b7280"}}>3 أوراق: موظفون، مرافقون، ملخص</div></div>
                  </button>
                  <div style={{height:1,background:darkMode?"#2e3340":"#f3f4f6"}}/>
                  <button onClick={()=>{exportToPDF(records);setShowExportMenu(false);}}
                    style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"11px 15px",border:"none",background:"none",cursor:"pointer",fontSize:13,fontFamily:"inherit",textAlign:"right"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#fef2f2"} onMouseLeave={e=>e.currentTarget.style.background="none"}>
                    <span style={{fontSize:17}}>📄</span><div><div style={{fontWeight:700,color:"#dc2626"}}>تصدير PDF</div><div style={{fontSize:11,color:"#6b7280"}}>تقرير شامل للطباعة</div></div>
                  </button>
                </div>
              )}
            </div>
            <button onClick={()=>{setForm(emptyForm);setEditId(null);setShowForm(true);setActiveTab("list");setShowExportMenu(false);}}
              style={{background:darkMode?"#F5A800":"#fff",color:darkMode?"#1a0a00":"#6B1A1A",border:"none",borderRadius:9,padding:"8px 16px",fontWeight:700,fontSize:14,cursor:"pointer"}}>
              ＋ إضافة
            </button>
          </div>
        </div>
      </div>

      <div style={{maxWidth:1200,margin:"0 auto",padding:"18px 14px",color:darkMode?"#e8eaf0":"inherit"}} onClick={()=>setShowExportMenu(false)}>
        {/* Stats Cards */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:10,marginBottom:18}}>
          {[
            {label:"الإجمالي",value:stats.total,color:"#1e3a5f",icon:"📋",filter:null},
            {label:"موظفون",value:stats.employees,color:"#374151",icon:"👤",filter:"موظف",fKey:"type"},
            {label:"مرافقون",value:stats.dependents,color:"#7c3aed",icon:"👨‍👩‍👧",filter:"مرافق",fKey:"type"},
            {label:"سارية",value:stats.valid,color:"#16a34a",icon:"✅",filter:"سارية",fKey:"status"},
            {label:"تنتهي قريباً",value:stats.soon,color:"#d97706",icon:"⚠️",filter:"تنتهي قريباً",fKey:"status"},
            {label:"منتهية",value:stats.expired,color:"#dc2626",icon:"❌",filter:"منتهية",fKey:"status"},
            {label:"انجال المشاعر",value:stats.anjal,color:"#6B1A1A",icon:"🏢",filter:"انجال المشاعر",fKey:"company"},
            {label:"دلتا الماسية",value:stats.delta,color:"#b45309",icon:"🏢",filter:"دلتا الماسية",fKey:"company"},
          ].map(s=>(
            <div key={s.label}
              onClick={()=>{ if(s.value>0) setModalCard(s); }}
              style={{background:darkMode?"#1c1f26":"#fff",borderRadius:10,padding:"12px 10px",boxShadow:darkMode?"0 2px 6px rgba(0,0,0,0.5)":"0 2px 6px rgba(0,0,0,0.07)",borderTop:`3px solid ${s.color}`,textAlign:"center",cursor:s.value>0?"pointer":"default",transition:"all 0.15s",transform:"scale(1)",userSelect:"none"}}
              onMouseEnter={e=>{if(s.value>0)e.currentTarget.style.transform="scale(1.04) translateY(-2px)";}}
              onMouseLeave={e=>{e.currentTarget.style.transform="scale(1)";}}>
              <div style={{fontSize:20}}>{s.icon}</div>
              <div style={{fontSize:24,fontWeight:800,color:s.color}}>{s.value}</div>
              <div style={{fontSize:11,color:darkMode?"#9299aa":"#6b7280",marginTop:1}}>{s.label}</div>
            </div>
          ))}
        </div>

        {/* Tabs */}
        <div style={{display:"flex",gap:8,marginBottom:16}}>
          {["list","alerts","family","cost"].map(tab=>(
            <button key={tab} onClick={()=>setActiveTab(tab)}
              style={{padding:"7px 18px",borderRadius:8,border:"none",cursor:"pointer",fontWeight:600,fontSize:13,fontFamily:"inherit",
                background:activeTab===tab?"#8B2500":darkMode?"#1c1f26":"#fff",color:activeTab===tab?"#fff":darkMode?"#e8eaf0":"#374151",border:darkMode&&activeTab!==tab?"1px solid #2e3340":"none",boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.5)":"0 2px 6px rgba(0,0,0,0.07)"}}>
              {tab==="list"?"📋 الكل":tab==="alerts"?`🔔 تنبيهات ${stats.expired+stats.soon>0?`(${stats.expired+stats.soon})`:""}`:tab==="family"?`👨‍👩‍👧 عائلات (${stats.dependents})`:"💰 حاسبة التكلفة"}
            </button>
          ))}
        </div>

        {/* ALERTS */}
        {activeTab==="alerts"&&(
          <div style={{background:darkMode?"#1c1f26":"#fff",borderRadius:14,padding:22,boxShadow:darkMode?"0 2px 10px rgba(0,0,0,0.5)":"0 2px 10px rgba(0,0,0,0.07)"}}>
            <h3 style={{margin:"0 0 14px",color:darkMode?"#e8eaf0":"#1e3a5f"}}>🔔 إقامات منتهية أو تنتهي قريباً</h3>
            {records.filter(r=>r.expiryDate&&getDaysLeft(r.expiryDate)<=30).length===0?(
              <div style={{textAlign:"center",color:"#6b7280",padding:30}}><div style={{fontSize:44}}>✅</div><p>لا توجد تنبيهات.</p></div>
            ):records.filter(r=>r.expiryDate&&getDaysLeft(r.expiryDate)<=30).sort((a,b)=>getDaysLeft(a.expiryDate)-getDaysLeft(b.expiryDate)).map(r=>{
              const st=getStatus(r),sc=STATUS_COLORS[st],days=getDaysLeft(r.expiryDate);
              const cc=COMPANY_COLORS[r.company]||{bg:"#f9f9f9",text:"#374151"};
              return (
                <div key={r.id} style={{background:sc.bg,border:`1px solid ${sc.border}`,borderRadius:10,padding:"12px 16px",marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
                  <div>
                    <strong style={{color:sc.text,fontSize:14}}>{r.name}</strong>
                    <span style={{marginRight:8,color:"#6b7280",fontSize:12}}>{r.jobTitle||r.type}</span>
                    <span style={{background:cc.bg,color:cc.text,padding:"1px 8px",borderRadius:8,fontSize:11,marginRight:4}}>{r.company}</span>
                  </div>
                  <div style={{fontWeight:700,color:sc.text,fontSize:15}}>{days<0?`منتهية منذ ${Math.abs(days)} يوم`:`${days} يوم متبقي`}</div>
                  <span style={{background:sc.bg,border:`1px solid ${sc.border}`,color:sc.text,padding:"3px 10px",borderRadius:16,fontSize:12,fontWeight:600}}>{st}</span>
                </div>
              );
            })}
          </div>
        )}

        {/* FAMILY VIEW */}
        {activeTab==="family"&&(
          <div>
            {records.filter(r=>r.type!=="مرافق"&&records.some(d=>d.familyHeadId===r.iqamaNumber)).map(head=>{
              const deps=records.filter(d=>d.familyHeadId===head.iqamaNumber);
              const hSt=getStatus(head),hSc=STATUS_COLORS[hSt];
              const cc=COMPANY_COLORS[head.company]||{bg:"#f9f9f9",text:"#374151",border:"#e5e7eb"};
              return (
                <div key={head.id} style={{background:darkMode?"#1c1f26":"#fff",borderRadius:14,padding:"18px 20px",marginBottom:14,boxShadow:darkMode?"0 2px 10px rgba(0,0,0,0.5)":"0 2px 10px rgba(0,0,0,0.07)",borderRight:`5px solid ${hSc.border}`}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8,marginBottom:12}}>
                    <div>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                        <span style={{fontSize:18}}>👤</span>
                        <strong style={{fontSize:16,color:darkMode?"#e8eaf0":"#1e3a5f"}}>{head.name}</strong>
                        <span style={{background:cc.bg,color:cc.text,border:`1px solid ${cc.border}`,padding:"2px 10px",borderRadius:12,fontSize:11,fontWeight:600}}>{head.company}</span>
                      </div>
                      <div style={{fontSize:12,color:"#6b7280",display:"flex",gap:12}}>
                        <span>🪪 {head.iqamaNumber}</span>
                        <span>💼 {head.jobTitle}</span>
                        <span>🌍 {head.nationality}</span>
                      </div>
                    </div>
                    <span style={{background:hSc.bg,border:`1px solid ${hSc.border}`,color:hSc.text,padding:"4px 12px",borderRadius:16,fontSize:12,fontWeight:700}}>{hSt} · {head.expiryDate?`${getDaysLeft(head.expiryDate)} يوم`:"—"}</span>
                  </div>
                  <div style={{background:darkMode?"#252830":"#f8faff",borderRadius:10,padding:"10px 14px"}}>
                    <div style={{fontSize:12,fontWeight:700,color:darkMode?"#c9cdd6":"#374151",marginBottom:8}}>👨‍👩‍👧 المرافقون ({deps.length})</div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:8}}>
                      {deps.map(d=>(
                        <div key={d.id} style={{background:darkMode?"#1c1f26":"#fff",borderRadius:8,padding:"8px 12px",border:`1px solid ${darkMode?"#2e3340":"#e0e7ff"}`,display:"flex",alignItems:"center",gap:8}}>
                          <span style={{fontSize:16}}>{RELATION_ICONS[d.relation]||"👤"}</span>
                          <div>
                            <div style={{fontSize:13,fontWeight:600,color:darkMode?"#e8eaf0":"#374151"}}>{d.name}</div>
                            <div style={{fontSize:11,color:"#6b7280"}}>{d.relation} · {d.gender} · {d.iqamaNumber}</div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* LIST */}
        {activeTab==="list"&&(
          <>
            <div style={{background:darkMode?"#1c1f26":"#fff",borderRadius:11,padding:"13px 16px",marginBottom:14,boxShadow:darkMode?"0 2px 7px rgba(0,0,0,0.5)":"0 2px 7px rgba(0,0,0,0.06)",display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 بحث بالاسم، الرقم، المهنة..."
                style={{...inp,flex:"1 1 180px",maxWidth:260}}/>
              <select value={filterCompany} onChange={e=>setFilterCompany(e.target.value)} style={{...inp,width:"auto"}}>
                <option>الكل</option><option>انجال المشاعر</option><option>دلتا الماسية</option>
              </select>
              <select value={filterType} onChange={e=>setFilterType(e.target.value)} style={{...inp,width:"auto"}}>
                <option>الكل</option><option>موظف</option><option>مرافق</option>
              </select>
              <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} style={{...inp,width:"auto"}}>
                {["الكل","سارية","تنتهي قريباً","منتهية","قيد التجديد","مرافق"].map(s=><option key={s}>{s}</option>)}
              </select>
              <select value={filterNationality} onChange={e=>setFilterNationality(e.target.value)} style={{...inp,width:"auto"}}>
                {nationalities.map(n=><option key={n}>{n}</option>)}
              </select>
              <select value={sortBy} onChange={e=>setSortBy(e.target.value)} style={{...inp,width:"auto"}}>
                <option value="expiryDate">ترتيب: تاريخ الانتهاء</option>
                <option value="name">ترتيب: الاسم</option>
                <option value="company">ترتيب: الشركة</option>
              </select>
              <span style={{color:"#6b7280",fontSize:12,whiteSpace:"nowrap"}}>{filtered.length} سجل</span>
            </div>

            {showForm&&(
              <div style={{background:darkMode?"#1c1f26":"#fff",borderRadius:14,padding:22,marginBottom:16,boxShadow:darkMode?"0 4px 20px rgba(0,0,0,0.6)":"0 4px 20px rgba(0,0,0,0.12)",border:`2px solid ${darkMode?"#F5A800":"#2563eb"}`}}>
                <h3 style={{margin:"0 0 16px",color:darkMode?"#e8eaf0":"#1e3a5f",fontSize:15}}>{editId?"✏️ تعديل":"➕ إضافة إقامة جديدة"}</h3>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:11}}>
                  {[{l:"الاسم *",k:"name",t:"text"},{l:"الجنسية",k:"nationality",t:"text"},{l:"رقم الإقامة *",k:"iqamaNumber",t:"text"},{l:"تاريخ انتهاء الإقامة",k:"expiryDate",t:"date"},{l:"رقم الجواز",k:"passportNumber",t:"text"},{l:"المهنة",k:"jobTitle",t:"text"},{l:"تكلفة التجديد (ريال)",k:"renewalCost",t:"number"},{l:"رقم إقامة رب الأسرة",k:"familyHeadId",t:"text"}].map(f=>(
                    <div key={f.k}>
                      <label style={{display:"block",marginBottom:3,fontSize:12,fontWeight:600,color:darkMode?"#c9cdd6":"#374151"}}>{f.l}</label>
                      <input type={f.t} value={form[f.k]} onChange={e=>setForm({...form,[f.k]:e.target.value})} style={inp}/>
                    </div>
                  ))}
                  {[{l:"النوع",k:"type",o:["موظف","مرافق"]},{l:"صلة القرابة",k:"relation",o:["","زوجة","ابن","بنت"]},{l:"الجنس",k:"gender",o:["ذكر","أنثى"]},{l:"خارج المملكة",k:"outsideKingdom",o:["لا","نعم"]},{l:"حالة التجديد",k:"renewalStatus",o:["لم يبدأ","قيد التجديد","مكتمل"]},{l:"الشركة",k:"company",o:["انجال المشاعر","دلتا الماسية"]}].map(f=>(
                    <div key={f.k}>
                      <label style={{display:"block",marginBottom:3,fontSize:12,fontWeight:600,color:"#374151"}}>{f.l}</label>
                      <select value={form[f.k]} onChange={e=>setForm({...form,[f.k]:e.target.value})} style={inp}>{f.o.map(o=><option key={o}>{o}</option>)}</select>
                    </div>
                  ))}
                </div>
                <div style={{display:"flex",gap:8,marginTop:14}}>
                  <button onClick={handleSubmit} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:8,padding:"8px 22px",fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>{editId?"💾 حفظ":"✅ إضافة"}</button>
                  <button onClick={()=>{setShowForm(false);setEditId(null);setForm(emptyForm);}} style={{background:darkMode?"#252830":"#f3f4f6",color:darkMode?"#e8eaf0":"#374151",border:darkMode?"1px solid #2e3340":"none",borderRadius:8,padding:"8px 16px",fontWeight:600,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>إلغاء</button>
                </div>
              </div>
            )}

            {filtered.filter(r=>r.type!=="مرافق").length===0?(
              <div style={{textAlign:"center",background:darkMode?"#1c1f26":"#fff",borderRadius:14,padding:50,boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.5)":"0 2px 8px rgba(0,0,0,0.06)"}}>
                <div style={{fontSize:44}}>🔍</div><h3 style={{color:darkMode?"#e8eaf0":"#374151",marginTop:8}}>لا توجد نتائج</h3>
              </div>
            ):(
              <div style={{display:"grid",gap:10}}>
                {filtered.filter(r=>r.type!=="مرافق").map(r=>{
                  const st=getStatus(r),sc=STATUS_COLORS[st],days=r.expiryDate?getDaysLeft(r.expiryDate):null;
                  const cc=COMPANY_COLORS[r.company]||{bg:"#f9f9f9",text:"#374151",border:"#e5e7eb"};
                  const expanded=expandedId===r.id;
                  // المرافقون المرتبطون برب الأسرة هذا
                  const dependents=records.filter(d=>d.familyHeadId===r.iqamaNumber);
                  const hasDependents=dependents.length>0;
                  return (
                    <div key={r.id} style={{background:darkMode?"#1c1f26":"#fff",borderRadius:12,boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.5)":"0 2px 8px rgba(0,0,0,0.07)",borderRight:`5px solid ${sc.border}`,overflow:"hidden"}}>
                      {/* ── صف رب الأسرة ── */}
                      <div style={{padding:"14px 18px"}}>
                        <div style={{display:"flex",flexWrap:"wrap",gap:10,alignItems:"center",justifyContent:"space-between"}}>
                          {/* الاسم والتفاصيل */}
                          <div style={{flex:"1 1 200px",cursor:"pointer"}} onClick={()=>setExpandedId(expanded?null:r.id)}>
                            <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:4,flexWrap:"wrap"}}>
                              <span style={{fontSize:17}}>{r.gender==="أنثى"?"👩":"👤"}</span>
                              <strong style={{fontSize:15,color:darkMode?"#e8eaf0":"#1e3a5f"}}>{r.name}</strong>
                              <span style={{background:cc.bg,color:cc.text,border:`1px solid ${cc.border}`,padding:"1px 8px",borderRadius:10,fontSize:11,fontWeight:600}}>{r.company}</span>
                              {r.outsideKingdom==="نعم"&&<span style={{background:"#fef3c7",color:"#d97706",padding:"1px 7px",borderRadius:10,fontSize:11}}>✈️ خارج</span>}
                              {hasDependents&&(
                                <span style={{background:"#f3e8ff",color:"#7c3aed",padding:"1px 8px",borderRadius:10,fontSize:11,fontWeight:600,display:"flex",alignItems:"center",gap:3}}>
                                  👨‍👩‍👧 {dependents.length} مرافق
                                </span>
                              )}
                            </div>
                            <div style={{display:"flex",gap:12,fontSize:12,color:"#6b7280",flexWrap:"wrap"}}>
                              <span>🪪 {r.iqamaNumber}</span>
                              {r.nationality&&<span>🌍 {r.nationality}</span>}
                              {(r.jobTitle||r.notes)&&<span>💼 {r.jobTitle||r.notes}</span>}
                            </div>
                          </div>
                          {/* الأيام */}
                          <div style={{textAlign:"center",minWidth:80}}>
                            {days!==null?(
                              <>
                                <div style={{fontWeight:800,fontSize:20,color:sc.text}}>{days<0?Math.abs(days):days}</div>
                                <div style={{fontSize:10,color:"#6b7280"}}>{days<0?"يوم منتهي":"يوم متبقي"}</div>
                                <div style={{fontSize:11,color:"#9ca3af"}}>{new Date(r.expiryDate).toLocaleDateString("ar-SA")}</div>
                              </>
                            ):null}
                          </div>
                          {/* الحالة والأزرار */}
                          <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>
                            <span style={{background:sc.bg,border:`1px solid ${sc.border}`,color:sc.text,padding:"4px 11px",borderRadius:16,fontSize:12,fontWeight:700}}>{st}</span>
                            <div style={{display:"flex",gap:5,alignItems:"center"}}>
                              {hasDependents&&(
                                <button onClick={()=>setExpandedId(expanded?null:r.id)}
                                  title="عرض/إخفاء المرافقين"
                                  style={{background:expanded?"#f3e8ff":"#faf5ff",color:"#7c3aed",border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:3}}>
                                  {expanded?"▲":"▼"} مرافقون
                                </button>
                              )}
                              <button onClick={()=>handleEdit(r)} style={{background:"#eff6ff",color:"#2563eb",border:"1px solid #bfdbfe",borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>✏️</button>
                              <button onClick={()=>handleDelete(r.id)} style={{background:"#fef2f2",color:"#dc2626",border:"1px solid #fecaca",borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>🗑️</button>
                            </div>
                          </div>
                        </div>
                        {/* تفاصيل رب الأسرة عند الضغط على الاسم */}
                        {expanded&&(
                          <div style={{marginTop:10,paddingTop:10,borderTop:"1px dashed #e5e7eb",display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:7}}>
                            {[{l:"رقم الإقامة",v:r.iqamaNumber},{l:"الجنسية",v:r.nationality},{l:"الجنس",v:r.gender},{l:"رقم الجواز",v:r.passportNumber},{l:"المهنة",v:r.jobTitle||r.notes},{l:"خارج المملكة",v:r.outsideKingdom},{l:"حالة التجديد",v:r.renewalStatus},{l:"الشركة",v:r.company}].map(f=>(
                              <div key={f.l} style={{background:darkMode?"#252830":"#f0f2f5",borderRadius:7,padding:"7px 10px"}}>
                                <div style={{fontSize:10,color:"#6b7280",marginBottom:1}}>{f.l}</div>
                                <div style={{fontSize:12,fontWeight:600,color:darkMode?"#e8eaf0":"#1f2937"}}>{f.v||"-"}</div>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>

                      {/* ── قسم المرافقين المنسدل ── */}
                      {expanded&&hasDependents&&(
                        <div style={{background:darkMode?"#1a1128":"#faf5ff",borderTop:`2px dashed ${darkMode?"#5b21b6":"#c4b5fd"}`,padding:"14px 18px"}}>
                          <div style={{fontSize:13,fontWeight:700,color:"#7c3aed",marginBottom:10,display:"flex",alignItems:"center",gap:6}}>
                            👨‍👩‍👧 المرافقون ({dependents.length})
                          </div>
                          <div style={{display:"grid",gap:8}}>
                            {dependents.map(d=>{
                              const dcc=COMPANY_COLORS[d.company]||{bg:"#f9f9f9",text:"#374151",border:"#e5e7eb"};
                              return (
                                <div key={d.id} style={{background:darkMode?"#1c1f26":"#fff",borderRadius:10,padding:"12px 16px",border:`1px solid ${darkMode?"#3b1f6e":"#e9d5ff"}`,display:"flex",flexWrap:"wrap",alignItems:"center",justifyContent:"space-between",gap:8}}>
                                  <div style={{flex:"1 1 180px"}}>
                                    <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:3,flexWrap:"wrap"}}>
                                      <span style={{fontSize:16}}>{RELATION_ICONS[d.relation]||"👤"}</span>
                                      <strong style={{fontSize:14,color:darkMode?"#c4b5fd":"#4c1d95"}}>{d.name}</strong>
                                      <span style={{background:"#f3e8ff",color:"#7c3aed",padding:"1px 8px",borderRadius:8,fontSize:11,fontWeight:600}}>{d.relation||"مرافق"}</span>
                                      {d.outsideKingdom==="نعم"&&<span style={{background:"#fef3c7",color:"#d97706",padding:"1px 7px",borderRadius:8,fontSize:11}}>✈️ خارج</span>}
                                    </div>
                                    <div style={{display:"flex",gap:10,fontSize:11,color:"#6b7280",flexWrap:"wrap"}}>
                                      <span>🪪 {d.iqamaNumber}</span>
                                      <span>🌍 {d.nationality}</span>
                                      <span>{d.gender==="أنثى"?"♀️ أنثى":"♂️ ذكر"}</span>
                                      {d.passportNumber&&<span>📘 {d.passportNumber}</span>}
                                    </div>
                                  </div>
                                  <div style={{display:"flex",gap:5}}>
                                    <button onClick={()=>handleEdit(d)} style={{background:"#eff6ff",color:"#2563eb",border:"1px solid #bfdbfe",borderRadius:6,padding:"4px 9px",fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>✏️</button>
                                    <button onClick={()=>handleDelete(d.id)} style={{background:"#fef2f2",color:"#dc2626",border:"1px solid #fecaca",borderRadius:6,padding:"4px 9px",fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>🗑️</button>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </>
        )}

        {/* ── حاسبة التكلفة ── */}
        {activeTab==="cost"&&(()=>{
          const WORK_PERMIT_PER_Q     = 2425;
          const PASSPORT_NORMAL_PER_Q = 163;
          const PASSPORT_LATE_PER_Q   = 413;
          const PASSPORT_DEP_PER_Q    = 1200;

          const quarters = Math.ceil(calcMonths / 3);

          const companyFiltered = records.filter(r => {
            if (calcTarget === "company1") return r.company === "انجال المشاعر";
            if (calcTarget === "company2") return r.company === "دلتا الماسية";
            return true;
          });

          const allEmployees = companyFiltered.filter(r => r.type !== "مرافق");

          // ترتيب الأولوية: منتهية (0) → تنتهي قريباً (1) → سارية (2) → مرافق (3)
          const priorityOrder = emp => {
            const st = getStatus(emp);
            if (st === "منتهية") return 0;
            if (st === "تنتهي قريباً") return 1;
            if (st === "سارية") return 2;
            return 3;
          };

          // فلترة الموظفين حسب الحالة
          const statusFilteredEmployees = calcStatusFilter === "all"
            ? allEmployees
            : allEmployees.filter(emp => {
                const st = getStatus(emp);
                if (calcStatusFilter === "expired") return st === "منتهية";
                if (calcStatusFilter === "soon")    return st === "تنتهي قريباً";
                if (calcStatusFilter === "valid")   return st === "سارية";
                return true;
              });

          // إحصائيات الحالات للأزرار
          const countExpired = allEmployees.filter(e => getStatus(e) === "منتهية").length;
          const countSoon    = allEmployees.filter(e => getStatus(e) === "تنتهي قريباً").length; // 30 يوم
          const countValid   = allEmployees.filter(e => getStatus(e) === "سارية").length;

          // ترتيب الموظفين بالأولوية
          const sortedEmployees = [...statusFilteredEmployees].sort((a,b) => priorityOrder(a) - priorityOrder(b) || getDaysLeft(a.expiryDate) - getDaysLeft(b.expiryDate));

          // Set فارغ = الكل محدد، Set بقيم = المحدد فقط
          const hasSelection = calcSelectedIds.size > 0;
          const isEmpSelected = emp => !hasSelection || calcSelectedIds.has(emp.id);
          const targetEmployees = allEmployees.filter(isEmpSelected);

          // المرافقون: كل موظف محدد له خيار مستقل شمل/استثناء مرافقيه
          // calcIncludeDeps[empId] === false → استثنِ مرافقيه، غير ذلك → شملهم
          const targetDependents = companyFiltered.filter(r => {
            if (r.type !== "مرافق") return false;
            const head = targetEmployees.find(e => e.iqamaNumber === r.familyHeadId);
            if (!head) return false;
            return calcIncludeDeps[head.id] !== false; // افتراضي: مشمول
          });

          const targetRecords = [...targetEmployees, ...targetDependents];

          const calcRecord = r => {
            const isEmployee = r.type !== "مرافق";
            const days = r.expiryDate ? getDaysLeft(r.expiryDate) : null;
            const isExpired = days !== null && days < 0;
            const expiredDays = isExpired ? Math.abs(days) : 0;
            const expiredQuarters = isExpired ? Math.ceil(expiredDays / 90) : 0;
            const useLateRate = isExpired && expiredDays > 3;
            let backlogPassport=0,backlogWork=0,renewPassport=0,renewWork=0;
            if (isEmployee) {
              const rateQ = useLateRate ? PASSPORT_LATE_PER_Q : PASSPORT_NORMAL_PER_Q;
              backlogPassport = expiredQuarters * rateQ;
              backlogWork     = expiredQuarters * WORK_PERMIT_PER_Q;
              renewPassport   = quarters * rateQ;
              renewWork       = quarters * WORK_PERMIT_PER_Q;
            } else {
              renewPassport = quarters * PASSPORT_DEP_PER_Q;
            }
            const totalBacklog = backlogPassport + backlogWork;
            const totalRenew   = renewPassport + renewWork;
            const total        = totalBacklog + totalRenew;
            return {r,isEmployee,isExpired,expiredDays,expiredQuarters,useLateRate,
                    backlogPassport,backlogWork,totalBacklog,renewPassport,renewWork,totalRenew,total};
          };

          const breakdown    = targetRecords.map(calcRecord);
          const grandTotal   = breakdown.reduce((s,b)=>s+b.total,0);
          const grandBacklog = breakdown.reduce((s,b)=>s+b.totalBacklog,0);
          const grandRenew   = breakdown.reduce((s,b)=>s+b.totalRenew,0);
          const expiredCount = breakdown.filter(b=>b.isExpired).length;
          const fmt = n => n.toLocaleString("ar-SA") + " ريال";

          // helpers للاختيار
          const toggleEmp = id => {
            const next = new Set(hasSelection ? calcSelectedIds : new Set(allEmployees.map(e=>e.id)));
            if(next.has(id)) next.delete(id); else next.add(id);
            setCalcSelectedIds(next);
          };
          const selectAll   = () => setCalcSelectedIds(new Set());            // فارغ = الكل
          const deselectAll = () => setCalcSelectedIds(new Set([-1]));        // قيمة وهمية = لا أحد
          const toggleDeps  = id => setCalcIncludeDeps(p=>({...p,[id]:p[id]===false?true:false}));

          return (
            <div>
              {/* ── شريط الإعدادات ── */}
              <div style={{background:darkMode?"#1c1f26":"#fff",borderRadius:12,padding:"16px 20px",marginBottom:14,boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.5)":"0 2px 8px rgba(0,0,0,0.07)",display:"flex",gap:16,flexWrap:"wrap",alignItems:"flex-start"}}>
                <div>
                  <label style={{display:"block",fontSize:12,fontWeight:700,color:darkMode?"#c9cdd6":"#374151",marginBottom:6}}>📅 مدة التجديد</label>
                  <div style={{display:"flex",gap:7}}>
                    {[3,6,9,12].map(m=>(
                      <button key={m} onClick={()=>setCalcMonths(m)}
                        style={{padding:"6px 14px",borderRadius:8,border:`2px solid ${calcMonths===m?"#2563eb":"#d1d5db"}`,background:calcMonths===m?"#2563eb":"#fff",color:calcMonths===m?"#fff":"#374151",fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>
                        {m} شهر
                      </button>
                    ))}
                  </div>
                </div>
                <div>
                  <label style={{display:"block",fontSize:12,fontWeight:700,color:darkMode?"#c9cdd6":"#374151",marginBottom:6}}>🏢 الشركة</label>
                  <div style={{display:"flex",gap:7}}>
                    {[{k:"all",l:"الكل"},{k:"company1",l:"انجال المشاعر"},{k:"company2",l:"دلتا الماسية"}].map(c=>(
                      <button key={c.k} onClick={()=>{setCalcTarget(c.k);setCalcSelectedIds(new Set());setCalcIncludeDeps({});setCalcStatusFilter('all');}}
                        style={{padding:"6px 12px",borderRadius:8,border:`2px solid ${calcTarget===c.k?"#2563eb":"#d1d5db"}`,background:calcTarget===c.k?"#2563eb":"#fff",color:calcTarget===c.k?"#fff":"#374151",fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>
                        {c.l}
                      </button>
                    ))}
                  </div>
                </div>
                <div style={{marginRight:"auto",background:darkMode?"#0a2218":"#f0fdf4",border:`1px solid ${darkMode?"#166534":"#86efac"}`,borderRadius:10,padding:"10px 18px",textAlign:"center",minWidth:190}}>
                  <div style={{fontSize:11,color:darkMode?"#4ade80":"#15803d",fontWeight:600}}>إجمالي التكلفة · {calcMonths} شهر</div>
                  <div style={{fontSize:24,fontWeight:900,color:darkMode?"#4ade80":"#15803d"}}>{fmt(grandTotal)}</div>
                  <div style={{fontSize:11,color:darkMode?"#9299aa":"#6b7280",marginTop:2}}>{targetRecords.length} سجل ({targetEmployees.length} موظف + {targetDependents.length} مرافق)</div>
                </div>
              </div>

              {/* ── اختيار الموظفين مع خيار المرافقين ── */}
              <div style={{background:darkMode?"#1c1f26":"#fff",borderRadius:12,padding:"16px 20px",marginBottom:14,boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.5)":"0 2px 8px rgba(0,0,0,0.06)"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
                  <div style={{fontWeight:700,fontSize:14,color:darkMode?"#e8eaf0":"#1e3a5f",display:"flex",alignItems:"center",gap:8}}>
                    👤 اختيار الموظفين
                    <span style={{background:darkMode?"#1e2d45":"#dbeafe",color:darkMode?"#93c5fd":"#2563eb",padding:"2px 10px",borderRadius:10,fontSize:12}}>
                      {targetEmployees.length} من {allEmployees.length} محدد
                    </span>
                  </div>
                  <div style={{display:"flex",gap:7}}>
                    <button onClick={selectAll}
                      style={{padding:"5px 12px",borderRadius:7,border:"1px solid #86efac",background:"#f0fdf4",color:"#15803d",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>
                      ✅ تحديد الكل
                    </button>
                    <button onClick={deselectAll}
                      style={{padding:"5px 12px",borderRadius:7,border:"1px solid #fca5a5",background:"#fef2f2",color:"#dc2626",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>
                      ✗ إلغاء الكل
                    </button>
                  </div>
                </div>

                {/* ── شريط فلتر الحالة ── */}
                <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
                  {[
                    {k:"all",    l:`الكل (${allEmployees.length})`,          bg:"#1c1f26", activeBg:"#374151",  border:"#4b5563",  activeBorder:"#9ca3af",  text:darkMode?"#9299aa":"#6b7280"},
                    {k:"expired",l:`❌ منتهية (${countExpired})`,             bg:darkMode?"#2c0f0f":"#fee2e2", activeBg:darkMode?"#7f1d1d":"#fca5a5", border:darkMode?"#7f1d1d":"#fca5a5", activeBorder:"#dc2626", text:"#dc2626"},
                    {k:"soon",   l:`⚠️ تنتهي قريباً (${countSoon})`,         bg:darkMode?"#2c2008":"#fef3c7", activeBg:darkMode?"#78350f":"#fcd34d", border:darkMode?"#78350f":"#fcd34d", activeBorder:"#d97706", text:"#d97706"},
                    {k:"valid",  l:`✅ سارية (${countValid})`,                bg:darkMode?"#0a2218":"#dcfce7", activeBg:darkMode?"#14532d":"#86efac", border:darkMode?"#14532d":"#86efac", activeBorder:"#16a34a", text:"#16a34a"},
                  ].map(f=>{
                    const isActive = calcStatusFilter === f.k;
                    return (
                      <button key={f.k}
                        onClick={()=>setCalcStatusFilter(f.k)}
                        style={{padding:"6px 14px",borderRadius:8,border:`2px solid ${isActive?f.activeBorder:f.border}`,background:isActive?f.activeBg:f.bg,color:f.text,fontWeight:isActive?700:500,fontSize:12,cursor:"pointer",fontFamily:"inherit",transition:"all 0.15s",opacity:isActive?1:0.8}}>
                        {f.l}
                      </button>
                    );
                  })}
                </div>

                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:8,maxHeight:320,overflowY:"auto"}}>
                  {sortedEmployees.map(emp=>{
                    const empSelected = isEmpSelected(emp);
                    const deps = companyFiltered.filter(d=>d.familyHeadId===emp.iqamaNumber);
                    const depsIncluded = calcIncludeDeps[emp.id] !== false;
                    const st=getStatus(emp), sc=STATUS_COLORS[st];
                    const rcc=COMPANY_COLORS[emp.company]||{bg:"#f9f9f9",text:"#374151"};
                    return (
                      <div key={emp.id} style={{borderRadius:10,border:`2px solid ${empSelected?"#F5A800":darkMode?"#2e3340":"#e5e7eb"}`,background:empSelected?(darkMode?"#1f1200":"#f0f9ff"):(darkMode?"#1c1f26":"#fafafa"),overflow:"hidden",position:"relative"}}>
                        {/* شريط الأولوية */}
                        <div style={{position:"absolute",right:0,top:0,bottom:0,width:4,borderRadius:"0 10px 10px 0",background:priorityOrder(emp)===0?"#dc2626":priorityOrder(emp)===1?"#d97706":"#16a34a"}}/>
                        {/* صف الموظف */}
                        <div onClick={()=>toggleEmp(emp.id)}
                          style={{display:"flex",alignItems:"center",gap:10,padding:"9px 12px",cursor:"pointer"}}>
                          <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${empSelected?"#2563eb":"#d1d5db"}`,background:empSelected?"#2563eb":"#fff",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
                            {empSelected&&<span style={{color:"#fff",fontSize:11,fontWeight:900}}>✓</span>}
                          </div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontWeight:600,fontSize:12,color:darkMode?"#e8eaf0":"#1e3a5f",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{emp.name}</div>
                            <div style={{display:"flex",gap:5,marginTop:2,flexWrap:"wrap"}}>
                              <span style={{background:rcc.bg,color:rcc.text,padding:"0 5px",borderRadius:5,fontSize:10}}>{emp.company}</span>
                              <span style={{background:sc.bg,color:sc.text,padding:"0 5px",borderRadius:5,fontSize:10}}>{st}</span>
                            </div>
                          </div>
                        </div>
                        {/* خيار المرافقين — يظهر فقط إذا الموظف محدد وله مرافقون */}
                        {empSelected && deps.length > 0 && (
                          <div
                            onClick={e=>{e.stopPropagation();toggleDeps(emp.id);}}
                            style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"7px 12px",background:depsIncluded?(darkMode?"#1a1128":"#f3e8ff"):(darkMode?"#252830":"#f9fafb"),borderTop:`1px dashed ${darkMode?"#5b21b6":"#e9d5ff"}`,cursor:"pointer",gap:8}}>
                            <div style={{display:"flex",alignItems:"center",gap:6}}>
                              <div style={{width:15,height:15,borderRadius:3,border:`2px solid ${depsIncluded?"#7c3aed":"#d1d5db"}`,background:depsIncluded?"#7c3aed":"#fff",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
                                {depsIncluded&&<span style={{color:"#fff",fontSize:9,fontWeight:900}}>✓</span>}
                              </div>
                              <span style={{fontSize:11,color:"#7c3aed",fontWeight:600}}>👨‍👩‍👧 شمل المرافقين ({deps.length})</span>
                            </div>
                            <div style={{display:"flex",gap:4}}>
                              {deps.map(d=>(
                                <span key={d.id} style={{background:"#ede9fe",color:"#6d28d9",padding:"1px 6px",borderRadius:6,fontSize:10}}>{d.relation||"مرافق"}</span>
                              ))}
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* ── بطاقات الملخص ── */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(148px,1fr))",gap:10,marginBottom:14}}>
                {[
                  {label:"الإجمالي الكلي",val:grandTotal,color:"#15803d",bg:"#f0fdf4",border:"#86efac",icon:"💰"},
                  {label:"تكلفة التجديد",val:grandRenew,color:"#2563eb",bg:"#eff6ff",border:"#bfdbfe",icon:"🔄"},
                  {label:"متأخرات منتهية",val:grandBacklog,color:"#dc2626",bg:"#fee2e2",border:"#fca5a5",icon:"⚠️"},
                  {label:"موظفون محددون",val:targetEmployees.length,color:"#374151",bg:"#f9fafb",border:"#e5e7eb",icon:"👤",isCount:true},
                  {label:"مرافقون مشمولون",val:targetDependents.length,color:"#7c3aed",bg:"#f3e8ff",border:"#c4b5fd",icon:"👨‍👩‍👧",isCount:true},
                  {label:"إقامات منتهية",val:expiredCount,color:"#dc2626",bg:"#fef2f2",border:"#fecaca",icon:"❌",isCount:true},
                ].map(c=>(
                  <div key={c.label} style={{background:c.bg,border:`1px solid ${c.border}`,borderRadius:10,padding:"12px",textAlign:"center"}}>
                    <div style={{fontSize:20}}>{c.icon}</div>
                    <div style={{fontSize:c.isCount?26:15,fontWeight:800,color:c.color,marginTop:2}}>{c.isCount?c.val:fmt(c.val)}</div>
                    <div style={{fontSize:10,color:"#6b7280",marginTop:2}}>{c.label}</div>
                  </div>
                ))}
              </div>

              {/* ── معادلة الحساب ── */}
              <div style={{background:darkMode?"#1a1500":"#fefce8",border:`1px solid ${darkMode?"#92400e":"#fcd34d"}`,borderRadius:10,padding:"10px 16px",marginBottom:14,fontSize:11,color:darkMode?"#fcd34d":"#78350f",lineHeight:1.9}}>
                <strong>📌 معادلة الحساب ({calcMonths} شهر = {quarters} ربع):</strong><br/>
                <span>👤 موظف سارية: رخصة عمل ({quarters}×2,425) + جواز ({quarters}×163) = {fmt(quarters*2425+quarters*163)}</span><br/>
                <span>👤 موظف منتهية &gt;3 أيام: نفس التجديد لكن الجواز بمعدل 413 ريال/ربع + متأخرات بنفس المعدل</span><br/>
                <span>👨‍👩‍👧 مرافق: رسوم جواز فقط ({quarters}×1,200={fmt(quarters*1200)})</span>
              </div>

              {/* ── جدول التفاصيل ── */}
              <div style={{background:darkMode?"#1c1f26":"#fff",borderRadius:12,boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.5)":"0 2px 8px rgba(0,0,0,0.07)",overflow:"hidden"}}>
                <div style={{background:darkMode?"#3d1000":"#6B1A1A",padding:"10px 16px",display:"grid",gridTemplateColumns:"2.2fr 0.8fr 0.9fr 1.1fr 1.1fr 1.1fr",gap:6,color:"#fff",fontSize:11,fontWeight:700}}>
                  <span>الاسم</span>
                  <span style={{textAlign:"center"}}>النوع</span>
                  <span style={{textAlign:"center"}}>الحالة</span>
                  <span style={{textAlign:"center"}}>متأخرات</span>
                  <span style={{textAlign:"center"}}>تجديد ({calcMonths}ش)</span>
                  <span style={{textAlign:"center"}}>الإجمالي</span>
                </div>
                {breakdown.sort((a,b)=>priorityOrder(a.r)-priorityOrder(b.r)||getDaysLeft(a.r.expiryDate)-getDaysLeft(b.r.expiryDate)).map(({r,isEmployee,isExpired,expiredDays,useLateRate,totalBacklog,totalRenew,total},i)=>{
                  const st=getStatus(r),sc=STATUS_COLORS[st];
                  const rcc=COMPANY_COLORS[r.company]||{bg:"#f9f9f9",text:"#374151"};
                  return (
                    <div key={r.id} style={{padding:"9px 16px",borderBottom:`1px solid ${darkMode?"#2e3340":"#f3f4f6"}`,background:i%2===0?(darkMode?'#1c1f26':'#fff'):(darkMode?'#191c23':'#fafafa'),display:"grid",gridTemplateColumns:"2.2fr 0.8fr 0.9fr 1.1fr 1.1fr 1.1fr",gap:6,alignItems:"center"}}>
                      <div>
                        <div style={{fontWeight:600,fontSize:12,color:isEmployee?(darkMode?"#e8eaf0":"#1e3a5f"):(darkMode?"#c4b5fd":"#5b21b6")}}>{r.name}</div>
                        <div style={{fontSize:10,color:"#6b7280",display:"flex",gap:5,marginTop:2,flexWrap:"wrap"}}>
                          <span style={{background:rcc.bg,color:rcc.text,padding:"0 5px",borderRadius:5}}>{r.company}</span>
                          {isExpired&&<span style={{color:"#dc2626"}}>⚠️ {expiredDays} يوم {useLateRate?"غرامة":""}</span>}
                          {!isEmployee&&<span style={{color:"#7c3aed"}}>{r.relation}</span>}
                        </div>
                      </div>
                      <div style={{textAlign:"center"}}>
                        <span style={{background:isEmployee?"#eff6ff":"#f3e8ff",color:isEmployee?"#2563eb":"#7c3aed",padding:"2px 7px",borderRadius:7,fontSize:10,fontWeight:600}}>{isEmployee?"موظف":"مرافق"}</span>
                      </div>
                      <div style={{textAlign:"center"}}>
                        <span style={{background:sc.bg,color:sc.text,padding:"2px 7px",borderRadius:7,fontSize:10,fontWeight:600}}>{st}</span>
                      </div>
                      <div style={{textAlign:"center",fontWeight:700,color:totalBacklog>0?"#dc2626":"#9ca3af",fontSize:12}}>
                        {totalBacklog>0?fmt(totalBacklog):"—"}
                      </div>
                      <div style={{textAlign:"center",fontWeight:600,color:"#2563eb",fontSize:12}}>{fmt(totalRenew)}</div>
                      <div style={{textAlign:"center",fontWeight:800,color:"#15803d",fontSize:13}}>{fmt(total)}</div>
                    </div>
                  );
                })}
                <div style={{padding:"11px 16px",background:darkMode?"#0a2218":"#f0fdf4",borderTop:`2px solid ${darkMode?"#166534":"#86efac"}`,display:"grid",gridTemplateColumns:"2.2fr 0.8fr 0.9fr 1.1fr 1.1fr 1.1fr",gap:6,alignItems:"center"}}>
                  <div style={{fontWeight:800,color:darkMode?"#4ade80":"#15803d",fontSize:13}}>الإجمالي ({targetRecords.length} سجل)</div>
                  <div/><div/>
                  <div style={{textAlign:"center",fontWeight:800,color:"#dc2626",fontSize:13}}>{grandBacklog>0?fmt(grandBacklog):"—"}</div>
                  <div style={{textAlign:"center",fontWeight:800,color:"#2563eb",fontSize:13}}>{fmt(grandRenew)}</div>
                  <div style={{textAlign:"center",fontWeight:900,color:"#15803d",fontSize:15}}>{fmt(grandTotal)}</div>
                </div>
              </div>
            </div>
          );
        })()}


        {/* ══════ النافذة المنبثقة ══════ */}
        {modalCard && (() => {
          // جلب السجلات المناسبة للبطاقة
          const getModalRecords = () => {
            if (!modalCard.fKey) return records; // الإجمالي
            if (modalCard.fKey === "status") return records.filter(r => getStatus(r) === modalCard.filter);
            if (modalCard.fKey === "company") return records.filter(r => r.company === modalCard.filter);
            if (modalCard.fKey === "type") {
              if (modalCard.filter === "موظف") return records.filter(r => r.type !== "مرافق");
              return records.filter(r => r.type === "مرافق");
            }
            return [];
          };

          const modalRecords = getModalRecords().sort((a,b) => {
            const pa = getStatus(a)==="منتهية"?0:getStatus(a)==="تنتهي قريباً"?1:2;
            const pb = getStatus(b)==="منتهية"?0:getStatus(b)==="تنتهي قريباً"?1:2;
            return pa-pb || getDaysLeft(a.expiryDate)-getDaysLeft(b.expiryDate);
          });

          const dm = darkMode;

          return (
            <div
              onClick={()=>setModalCard(null)}
              style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16,backdropFilter:"blur(3px)"}}>
              <div
                onClick={e=>e.stopPropagation()}
                style={{background:dm?"#1c1f26":"#fff",borderRadius:16,width:"100%",maxWidth:820,maxHeight:"88vh",display:"flex",flexDirection:"column",boxShadow:"0 24px 64px rgba(0,0,0,0.4)",overflow:"hidden"}}>

                {/* هيدر المودال */}
                <div style={{background:dm?"linear-gradient(135deg,#3d1000,#6B1A1A)":"linear-gradient(135deg,#6B1A1A,#F5A800)",padding:"18px 24px",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
                  <div style={{display:"flex",alignItems:"center",gap:12}}>
                    <span style={{fontSize:28}}>{modalCard.icon}</span>
                    <div>
                      <h2 style={{margin:0,fontSize:18,fontWeight:800,color:"#fff"}}>{modalCard.label}</h2>
                      <div style={{fontSize:12,color:"rgba(255,255,255,0.8)",marginTop:2}}>{modalRecords.length} سجل</div>
                    </div>
                  </div>
                  <button onClick={()=>setModalCard(null)}
                    style={{background:"rgba(255,255,255,0.15)",border:"1.5px solid rgba(255,255,255,0.4)",color:"#fff",borderRadius:8,width:34,height:34,fontSize:18,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700}}>
                    ✕
                  </button>
                </div>

                {/* المحتوى */}
                <div style={{overflowY:"auto",padding:"16px 20px",flex:1}}>
                  {modalRecords.length === 0 ? (
                    <div style={{textAlign:"center",padding:40,color:dm?"#9299aa":"#6b7280"}}>
                      <div style={{fontSize:44}}>📭</div>
                      <p style={{marginTop:8}}>لا توجد سجلات</p>
                    </div>
                  ) : modalRecords.map((r,i) => {
                    const st = getStatus(r);
                    const sc = STATUS_COLORS[st];
                    const days = r.expiryDate ? getDaysLeft(r.expiryDate) : null;
                    const cc = COMPANY_COLORS[r.company]||{bg:"#f9f9f9",text:"#374151",border:"#e5e7eb"};
                    const isDependent = r.type === "مرافق";
                    // رب الأسرة إذا كان مرافقاً
                    const headRecord = isDependent ? records.find(e=>e.iqamaNumber===r.familyHeadId) : null;

                    return (
                      <div key={r.id}
                        style={{background:dm?"#252830":"#f9fafb",borderRadius:12,padding:"14px 16px",marginBottom:10,borderRight:`4px solid ${sc.border}`,display:"flex",flexWrap:"wrap",gap:12,alignItems:"center",justifyContent:"space-between"}}>

                        {/* الاسم والتفاصيل */}
                        <div style={{flex:"1 1 200px"}}>
                          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:5}}>
                            <span style={{fontSize:16}}>{isDependent?(RELATION_ICONS[r.relation]||"👤"):r.gender==="أنثى"?"👩":"👤"}</span>
                            <strong style={{fontSize:14,color:dm?"#e8eaf0":"#1e3a5f"}}>{r.name}</strong>
                            {isDependent&&<span style={{background:"#f3e8ff",color:"#7c3aed",padding:"1px 8px",borderRadius:8,fontSize:11,fontWeight:600}}>{r.relation}</span>}
                            <span style={{background:cc.bg,color:cc.text,border:`1px solid ${cc.border}`,padding:"1px 7px",borderRadius:8,fontSize:11,fontWeight:600}}>{r.company}</span>
                            {r.outsideKingdom==="نعم"&&<span style={{background:"#fef3c7",color:"#d97706",padding:"1px 7px",borderRadius:8,fontSize:11}}>✈️ خارج</span>}
                          </div>
                          <div style={{display:"flex",gap:12,fontSize:12,color:dm?"#9299aa":"#6b7280",flexWrap:"wrap"}}>
                            <span>🪪 {r.iqamaNumber}</span>
                            {r.nationality&&<span>🌍 {r.nationality}</span>}
                            {r.jobTitle&&!isDependent&&<span>💼 {r.jobTitle}</span>}
                            {headRecord&&<span style={{color:dm?"#c4b5fd":"#7c3aed"}}>👤 رب الأسرة: {headRecord.name}</span>}
                          </div>
                        </div>

                        {/* الأيام والحالة */}
                        <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6,textAlign:"center"}}>
                          {days!==null && (
                            <div style={{textAlign:"center"}}>
                              <div style={{fontWeight:800,fontSize:20,color:sc.text,lineHeight:1}}>{days<0?Math.abs(days):days}</div>
                              <div style={{fontSize:10,color:dm?"#9299aa":"#6b7280"}}>{days<0?"يوم منتهي":"يوم متبقي"}</div>
                              <div style={{fontSize:11,color:dm?"#9299aa":"#9ca3af"}}>{new Date(r.expiryDate).toLocaleDateString("ar-SA")}</div>
                            </div>
                          )}
                          <span style={{background:sc.bg,border:`1px solid ${sc.border}`,color:sc.text,padding:"3px 10px",borderRadius:14,fontSize:12,fontWeight:700}}>{st}</span>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* فوتر */}
                <div style={{padding:"12px 20px",borderTop:`1px solid ${dm?"#2e3340":"#e5e7eb"}`,display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0,background:dm?"#1c1f26":"#f9fafb"}}>
                  <span style={{fontSize:12,color:dm?"#9299aa":"#6b7280"}}>اضغط خارج النافذة للإغلاق</span>
                  <button onClick={()=>setModalCard(null)}
                    style={{background:"#8B2500",color:"#fff",border:"none",borderRadius:8,padding:"7px 20px",fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>
                    إغلاق
                  </button>
                </div>
              </div>
            </div>
          );
        })()}

      </div>
    </div>
  );
}
