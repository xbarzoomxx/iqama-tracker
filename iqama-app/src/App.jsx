import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { ref, onValue, set } from "firebase/database";
import {
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signOut,
  onAuthStateChanged,
  sendPasswordResetEmail,
  updateProfile
} from "firebase/auth";
import { db, auth, DB_PATH as DB_PATH_CONST } from "./firebase.js";

const STORAGE_KEY = "iqama_tracker_v3";

// ── متطلبات التنصيف ──
const CLASSIFICATION_REQUIREMENTS = {
  "انجال المشاعر": {
    activity:"تصنيف أ — إنذار + إطفاء", icon:"🔥", color:"#6B1A1A",
    jobs:[
      {title:"سباك",required:4},
      {title:"مهندس كهرباء",required:3},
      {title:"مهندس ميكانيكا",required:3},
      {title:"مهندس إلكترونيات",required:2},
      {title:"فني كهرباء",required:9},
      {title:"فني إلكترونيات",required:3},
      {title:"لحام",required:4},
      {title:"فني ميكانيكا",required:2},
    ],
  },
  "دلتا الماسية": {
    activity:"مصاعد", icon:"🛗", color:"#b45309",
    jobs:[
      {title:"مهندس كهربائي",required:1},
      {title:"فني كهربائي",required:2},
      {title:"فني ميكانيكي",required:2},
      {title:"فني هندسة إلكترونيات",required:1},
      {title:"عمالة",required:4},
    ],
  },
};

const matchJob = (empJob, reqTitle) => {
  if (!empJob) return false;
  const j = empJob.trim().replace(/\s+/g," ").toLowerCase();
  const r = reqTitle.trim().toLowerCase();
  if (j === r) return true;
  if (r.includes("سباك") && j.includes("سباك")) return true;
  if (r.includes("مهندس كهرب") && j.includes("مهندس") && j.includes("كهرب")) return true;
  if (r.includes("مهندس ميكان") && j.includes("مهندس") && j.includes("ميكان")) return true;
  if (r.includes("مهندس إلكترون") && j.includes("مهندس") && (j.includes("إلكترون")||j.includes("الكترون"))) return true;
  if (r.includes("فني كهرب") && j.includes("فني") && j.includes("كهرب")) return true;
  if (r.includes("فني إلكترون") && j.includes("فني") && (j.includes("إلكترون")||j.includes("الكترون"))) return true;
  if (r.includes("فني ميكان") && j.includes("فني") && j.includes("ميكان")) return true;
  if (r.includes("لحام") && j.includes("لحام")) return true;
  if (r.includes("عمالة") && (j.includes("عمال")||j.includes("عامل"))) return true;
  // ميكانيكي مصاعد يُحتسب ضمن العمالة
  if (r.includes("عمالة") && j.includes("ميكانيكي") && j.includes("مصاعد")) return true;
  if (r.includes("مهندس كهربائي") && j.includes("مهندس") && j.includes("كهرب")) return true;
  if (r.includes("فني كهربائي") && j.includes("فني") && j.includes("كهرب")) return true;
  if (r.includes("فني ميكانيكي") && j.includes("فني") && j.includes("ميكان")) return true;
  if (r.includes("فني هندسة إلكترونيات") && j.includes("فني") && (j.includes("إلكترون")||j.includes("الكترون"))) return true;
  return false;
};

// ── بيانات السجلات التجارية ──
const COMPANY_CR = {
  "انجال المشاعر": {
    name:"شركة انجال المشاعر لاجهزة السلامة",
    crNumber:"4031081110", unifiedNumber:"7006539477",
    establishNumber:"13-1044947", type:"شركة ذات مسؤولية محدودة",
    status:"نشط", address:"مكة المكرمة - عبدالله خياط",
    email:"info@anjal.cc", manager:"عمار حسن علي الحسن",
    facilityManager:"عمار حسن علي الحسن",
  },
  "دلتا الماسية": {
    name:"شركة دلتا الماسية للمصاعد",
    crNumber:"", unifiedNumber:"7016055357",
    type:"شركة ذات مسؤولية محدودة",
    status:"نشط", issueDate:"2019-10-10",
    facilityManager:"عمار حسن علي الحسن",
    manager:"عمار حسن علي الحسن",
  },
  "البيوت الذكية": {
    name:"شركة البيوت الذكية للمراقبة الأمنية",
    crNumber:"", unifiedNumber:"7002772726",
    type:"شركة ذات مسؤولية محدودة",
    status:"نشط", issueDate:"2018-03-22",
    facilityManager:"عمار حسن علي الحسن",
    note:"لا يوجد موظفون مسجلون",
  },
};

const INITIAL_DATA = [{"id":2161651126,"name":"JOYNUL ABEDIN - - ABDUL MONAF","nationality":"بنجلاديش","iqamaNumber":"2161651126","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-18","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل صيانة أجهزة كهربائية","jobTitle":"عامل صيانة أجهزة كهربائية","passportNumber":"EM0468275","outsideKingdom":"نعم","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2168878490,"name":"عمار حسن علي الحسن","contractNumber":"2181161","contractStart":"2022-03-21","contractEnd":"2025-03-26","phone":"0564334224","email":"basha@anjal.cc","nationality":"السودان","iqamaNumber":"2168878490","contractDate":"2022-03-21","contractNum":"2181161","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-26","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مدير عام","jobTitle":"مدير عام","passportNumber":"P08151865","passportExpiry":"2026-06-28","passportIssue":"2021-06-29","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2499287072,"name":"ليلى عمار حسن علي","nationality":"السودان","iqamaNumber":"2499287072","type":"مرافق","relation":"بنت","gender":"أنثى","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"P08179155","passportExpiry":"2026-07-06","passportIssue":"2021-07-07","outsideKingdom":"لا","familyHeadId":"2168878490","employerId":"7006539477","company":"انجال المشاعر"},{"id":2374056725,"name":"محمد عمار حسن الحسن","nationality":"السودان","iqamaNumber":"2374056725","type":"مرافق","relation":"ابن","gender":"ذكر","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"P08228098","passportExpiry":"2026-07-21","passportIssue":"2021-07-22","outsideKingdom":"لا","familyHeadId":"2168878490","employerId":"7006539477","company":"انجال المشاعر"},{"id":2570076238,"name":"يزن عمار حسن علي","nationality":"السودان","iqamaNumber":"2570076238","type":"مرافق","relation":"ابن","gender":"ذكر","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"P11918293","passportExpiry":"2029-04-29","passportIssue":"2024-04-30","outsideKingdom":"لا","familyHeadId":"2168878490","employerId":"7006539477","company":"انجال المشاعر"},{"id":2374056485,"name":"هسه عمار حسن الحسن","nationality":"السودان","iqamaNumber":"2374056485","type":"مرافق","relation":"بنت","gender":"أنثى","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"P08153195","passportExpiry":"2026-06-29","passportIssue":"2021-06-30","outsideKingdom":"لا","familyHeadId":"2168878490","employerId":"7006539477","company":"انجال المشاعر"},{"id":2312831726,"name":"هبه محمداحمد محمد صديق","nationality":"السودان","iqamaNumber":"2312831726","type":"مرافق","relation":"زوجة","gender":"أنثى","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"P08151939","passportExpiry":"2026-06-28","passportIssue":"2021-06-29","outsideKingdom":"لا","familyHeadId":"2168878490","employerId":"7006539477","company":"انجال المشاعر"},{"id":2207144896,"name":"ZAFOR ULLAH ABDUL LATIB","nationality":"بنجلاديش","iqamaNumber":"2207144896","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-08","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني هندسة كهربائية","jobTitle":"فني هندسة كهربائية","passportNumber":"EM0405033","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2253325704,"name":"MOHAMED SHAHALAM MOHAMED SHAMSULHAQUE","nationality":"بنجلاديش","iqamaNumber":"2253325704","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-10-09","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني صيانة ميكانيكية","jobTitle":"فني صيانة ميكانيكية","passportNumber":"EH0582394","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2270598135,"name":"حماده منصور محمد احمد","nationality":"مصر","iqamaNumber":"2270598135","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-12-06","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني ميكانيكي تمديدات صحية","jobTitle":"فني ميكانيكي تمديدات صحية","passportNumber":"A28507403","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2285035610,"name":"حازم حشمت توفيق البسطويسي","nationality":"مصر","iqamaNumber":"2285035610","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-27","renewalStatus":"لم يبدأ","renewalCost":"","notes":"محاسب","jobTitle":"محاسب","passportNumber":"A29095351","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2316800917,"name":"محمد فتحي علي الفقي","nationality":"مصر","iqamaNumber":"2316800917","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-23","renewalStatus":"لم يبدأ","renewalCost":"","notes":"لحام","jobTitle":"لحام","passportNumber":"A04219484","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2368280851,"name":"عمار فضل المولي محمد الزبير","nationality":"السودان","iqamaNumber":"2368280851","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-02","renewalStatus":"لم يبدأ","renewalCost":"","notes":"سباك","jobTitle":"سباك","passportNumber":"P12124172","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2377607409,"name":"احمد حمدى احمد هلال","nationality":"مصر","iqamaNumber":"2377607409","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-11-19","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مهندس  ميكانيكي","jobTitle":"مهندس  ميكانيكي","passportNumber":"A28249488","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2391783970,"name":"بلال احمد حمدي هلال","nationality":"مصر","iqamaNumber":"2391783970","type":"مرافق","relation":"ابن","gender":"ذكر","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"A28249489","outsideKingdom":"نعم","familyHeadId":"2377607409","employerId":"7006539477","company":"انجال المشاعر"},{"id":2422905527,"name":"اروى احمد حمدي هلال","nationality":"مصر","iqamaNumber":"2422905527","type":"مرافق","relation":"بنت","gender":"أنثى","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"A32896771","outsideKingdom":"نعم","familyHeadId":"2377607409","employerId":"7006539477","company":"انجال المشاعر"},{"id":2545044733,"name":"زيد احمد حمدي هلال","nationality":"مصر","iqamaNumber":"2545044733","type":"مرافق","relation":"ابن","gender":"ذكر","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"A32896837","outsideKingdom":"نعم","familyHeadId":"2377607409","employerId":"7006539477","company":"انجال المشاعر"},{"id":2391783988,"name":"ساره محمود خليل محمد","nationality":"مصر","iqamaNumber":"2391783988","type":"مرافق","relation":"زوجة","gender":"أنثى","expiryDate":"","renewalStatus":"لم يبدأ","renewalCost":"","notes":"","jobTitle":"","passportNumber":"A28249643","outsideKingdom":"نعم","familyHeadId":"2377607409","employerId":"7006539477","company":"انجال المشاعر"},{"id":2383433394,"name":"NAFEES AHMAD MOHD MUSTFA","nationality":"الهند","iqamaNumber":"2383433394","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-28","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني كهربائي أنظمةحمايةكهربائية","jobTitle":"فني كهربائي أنظمةحمايةكهربائية","passportNumber":"V5262084","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2408364129,"name":"احمد جمال البسيوني ابراهيم","nationality":"مصر","iqamaNumber":"2408364129","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-04","renewalStatus":"لم يبدأ","renewalCost":"","notes":"سباك","jobTitle":"سباك","passportNumber":"A23626418","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2424076574,"name":"محمد عبدالله عبدالحليم عبدالله","nationality":"السودان","iqamaNumber":"2424076574","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-02","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مهندس  الكترونيات","jobTitle":"مهندس  الكترونيات","passportNumber":"P09775301","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2426377111,"name":"GAYYUR AHMAD MURSLEEN AHMAD","nationality":"الهند","iqamaNumber":"2426377111","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-20","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني كهربائي تمديدات كهربائية","jobTitle":"فني كهربائي تمديدات كهربائية","passportNumber":"U0688192","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2438640225,"name":"MOHAMMAD RAJU MOJUMDAR HAFEJA","nationality":"بنجلاديش","iqamaNumber":"2438640225","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2024-11-07","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل بناء","jobTitle":"عامل بناء","passportNumber":"EJ0383352","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2444383471,"name":"عبدالعليم قائد ناصر قاسم","nationality":"اليمن","iqamaNumber":"2444383471","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-09-25","renewalStatus":"لم يبدأ","renewalCost":"","notes":"سباك","jobTitle":"سباك","passportNumber":"10928633","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2464717319,"name":"MOHAMMED JOYNUL ABEDIN","nationality":"بنجلاديش","iqamaNumber":"2464717319","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-18","renewalStatus":"لم يبدأ","renewalCost":"","notes":"ميكانيكي مصاعد","jobTitle":"ميكانيكي مصاعد","passportNumber":"EH0762631","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2466025505,"name":"MANJOOR AHMED HAFIZ FAZRUDDIN","nationality":"الهند","iqamaNumber":"2466025505","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-26","renewalStatus":"لم يبدأ","renewalCost":"","notes":"لحام","jobTitle":"لحام","passportNumber":"Y6166044","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2479489730,"name":"FIROZ AHMED SAEED HUSSAIN","nationality":"الهند","iqamaNumber":"2479489730","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-19","renewalStatus":"لم يبدأ","renewalCost":"","notes":"سباك","jobTitle":"سباك","passportNumber":"T3702870","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2495781763,"name":"غسان على جلال احمد","contractNumber":"24432207","contractStart":"2024-12-24","contractEnd":"2025-12-23","phone":"0545580861","email":"ghassan@anjal.cc","nationality":"السودان","iqamaNumber":"2495781763","contractDate":"2024-12-22","contractNum":"24432207","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2025-10-01","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل صيانة أجهزة كهربائية","jobTitle":"عامل صيانة أجهزة كهربائية","passportNumber":"P07034750","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2508579659,"name":"FARHAN MANNAN ABDUL MANNAN","nationality":"باكستان","iqamaNumber":"2508579659","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-05-08","renewalStatus":"لم يبدأ","renewalCost":"","notes":"ميكانيكي مصاعد","jobTitle":"ميكانيكي مصاعد","passportNumber":"AR0704682","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2508580137,"name":"HABIBUR RAHMAN","nationality":"بنجلاديش","iqamaNumber":"2508580137","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-24","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني هندسة الكترونية","jobTitle":"فني هندسة الكترونية","passportNumber":"A18031168","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2510535806,"name":"اسد الله محمد مساعد سليمان","nationality":"السودان","iqamaNumber":"2510535806","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-14","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني صيانة أجهزة الكترونية","jobTitle":"فني صيانة أجهزة الكترونية","passportNumber":"P05665810","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2513490108,"name":"MD FOQRUL ISLAM","nationality":"بنجلاديش","iqamaNumber":"2513490108","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2024-02-20","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل صيانة أجهزة كهربائية","jobTitle":"عامل صيانة أجهزة كهربائية","passportNumber":"EN0083184","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2515822423,"name":"IRSHAD AHMED","nationality":"الهند","iqamaNumber":"2515822423","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-16","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني كهربائي تمديدات كهربائية","jobTitle":"فني كهربائي تمديدات كهربائية","passportNumber":"U0136451","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2525086464,"name":"ABDULLAH MUHAMMAD ASLAM KHAN","nationality":"باكستان","iqamaNumber":"2525086464","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-10","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني هندسة كهربائية","jobTitle":"فني هندسة كهربائية","passportNumber":"HB1073993","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2526249459,"name":"ابوبكر حبيب الله محمد فضل الله","nationality":"السودان","iqamaNumber":"2526249459","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2024-09-04","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مهندس  كهربائي","jobTitle":"مهندس  كهربائي","passportNumber":"P08667135","outsideKingdom":"لا","familyHeadId":"","employerId":"7016055357","company":"دلتا الماسية"},{"id":2526824905,"name":"BABU MANJUR AHMED","nationality":"الهند","iqamaNumber":"2526824905","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-11","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني صيانة ميكانيكية","jobTitle":"فني صيانة ميكانيكية","passportNumber":"U2024984","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2549937999,"name":"MOHAMMED ELIAS","nationality":"بنجلاديش","iqamaNumber":"2549937999","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-11","renewalStatus":"لم يبدأ","renewalCost":"","notes":"سباك","jobTitle":"سباك","passportNumber":"EG0771953","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2566347197,"name":"محمد احمد محمد قسم السيد","contractNumber":"24436149","contractStart":"2024-12-23","contractEnd":"2025-12-22","phone":"0535949982","email":"cfo@anjal.cc","nationality":"السودان","iqamaNumber":"2566347197","contractDate":"2024-12-22","contractNum":"24436149","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-06-30","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مدير مالي","jobTitle":"مدير مالي","passportNumber":"P09336349","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2567584533,"name":"حسن عطا المنان الحسن محمد","nationality":"السودان","iqamaNumber":"2567584533","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-07-20","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل بناء","jobTitle":"عامل بناء","passportNumber":"P11068420","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2571845128,"name":"المصطفي محمد علي الضو","contractNumber":"24433429","contractStart":"2024-12-23","contractEnd":"2025-12-22","phone":"0531606798","email":"mustafa.eldaw@gmail.com","nationality":"السودان","iqamaNumber":"2571845128","contractDate":"2024-12-22","contractNum":"24433429","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-21","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني كهربائي تمديدات كهربائية","jobTitle":"فني كهربائي تمديدات كهربائية","passportNumber":"P10162929","outsideKingdom":"نعم","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2575431966,"name":"ايات جاد الله محمد احمد البلوله","contractNumber":"24433747","contractStart":"2024-12-23","contractEnd":"2025-12-22","phone":"0574038225","email":"ayatjadallah3000@gmail.com","nationality":"السودان","iqamaNumber":"2575431966","contractDate":"2024-12-22","contractNum":"24433747","type":"موظف","relation":"","gender":"أنثى","expiryDate":"2026-06-12","renewalStatus":"لم يبدأ","renewalCost":"","notes":"مهندس  ميكانيكي","jobTitle":"مهندس  ميكانيكي","passportNumber":"P13192846","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2578801967,"name":"MD JONY MIA","nationality":"بنجلاديش","iqamaNumber":"2578801967","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2027-07-10","renewalStatus":"لم يبدأ","renewalCost":"","notes":"ميكانيكي معدات الكرتونية","jobTitle":"ميكانيكي معدات الكرتونية","passportNumber":"A11176721","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2590187510,"name":"محمد سر الختم عوض محمد","nationality":"السودان","iqamaNumber":"2590187510","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-18","renewalStatus":"لم يبدأ","renewalCost":"","notes":"فني كهربائي تمديدات كهربائية","jobTitle":"فني كهربائي تمديدات كهربائية","passportNumber":"P11400698","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2628316743,"name":"عبدالحليم محمد عثمان سليمان","nationality":"السودان","iqamaNumber":"2628316743","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2026-08-26","renewalStatus":"لم يبدأ","renewalCost":"","notes":"عامل صيانة أجهزة كهربائية","jobTitle":"عامل صيانة أجهزة كهربائية","passportNumber":"P13512941","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2388580298,"name":"الجميل محمد داؤد علوان","contractNumber":"24431968","contractStart":"2024-12-23","contractEnd":"2025-12-22","phone":"0534460589","email":"algamil@anjal.cc","nationality":"السودان","iqamaNumber":"2388580298","contractDate":"2024-12-22","contractNum":"24431968","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2025-01-11","renewalStatus":"قيد التجديد","renewalCost":"","notes":"مهندس ميكانيكي","jobTitle":"مهندس ميكانيكي","passportNumber":"","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2566969032,"name":"داليا عمر عوض الله صالح","contractNumber":"24433245","contractStart":"2024-12-23","contractEnd":"2025-12-22","phone":"0509367163","email":"dalia.omer@hotmail.com","nationality":"السودان","iqamaNumber":"2566969032","contractDate":"2024-12-22","contractNum":"24433245","type":"موظف","relation":"","gender":"أنثى","expiryDate":"2024-07-31","renewalStatus":"قيد التجديد","renewalCost":"","notes":"مهندس إلكترونيات","jobTitle":"مهندس إلكترونيات","passportNumber":"","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2505071395,"name":"عمر فاروق احمد حسين","contractNumber":"24432287","contractStart":"2024-12-23","contractEnd":"2025-12-22","phone":"0546216077","email":"omarfaroug910@gmail.com","nationality":"السودان","iqamaNumber":"2505071395","contractDate":"2024-12-22","contractNum":"24432287","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2024-08-03","renewalStatus":"قيد التجديد","renewalCost":"","notes":"عامل صيانة أجهزة كهربائية","jobTitle":"عامل صيانة أجهزة كهربائية","passportNumber":"","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"},{"id":2513005807,"name":"عبدالله بقاري محمد محجوب","contractNumber":"26000907","contractStart":"2025-03-03","contractEnd":"2026-03-02","phone":"0573437944","email":"tooomsa71993@gmail.com","nationality":"السودان","iqamaNumber":"2513005807","contractDate":"2025-03-03","contractNum":"26000907","type":"موظف","relation":"","gender":"ذكر","expiryDate":"2024-02-09","renewalStatus":"قيد التجديد","renewalCost":"","notes":"فني هندسة ميكانيكية","jobTitle":"فني هندسة ميكانيكية","passportNumber":"","outsideKingdom":"لا","familyHeadId":"","employerId":"7006539477","company":"انجال المشاعر"}];

const STATUS_COLORS = {
  منتهية:        { bg:"#fee2e2", text:"#dc2626", border:"#fca5a5" },
  "تنتهي قريباً":{ bg:"#fef3c7", text:"#d97706", border:"#fcd34d" },
  سارية:         { bg:"#dcfce7", text:"#16a34a", border:"#86efac" },
  "قيد التجديد": { bg:"#dbeafe", text:"#2563eb", border:"#93c5fd" },
  مرافق:         { bg:"#f3e8ff", text:"#7c3aed", border:"#c4b5fd" },
};

const COMPANY_INFO = {
  "انجال المشاعر": {
    crNumber: "7006539477",
    crCommercial: "4031081110",
    crDate: "2013-11-25",
    type: "شركة ذات مسؤولية محدودة",
    status: "نشط",
    email: "info@anjal.cc",
    address: "مكة المكرمة، عبدالله خياط",
  },
  "دلتا الماسية": {
    crNumber: "7016055357",
    crDate: "2019-10-10",
    type: "شركة ذات مسؤولية محدودة",
    status: "نشط",
  },
};

const COMPANY_COLORS = {
  "انجال المشاعر": { bg:"#fdf0f0", text:"#6B1A1A", border:"#e8b4b4" },
  "دلتا الماسية":  { bg:"#fffbeb", text:"#b45309", border:"#fcd34d" },
  "البيوت الذكية": { bg:"#f0f9ff", text:"#0369a1", border:"#bae6fd" },
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

// ══════════════════════════════
// صفحة تسجيل الدخول
// ══════════════════════════════
function AuthScreen({ darkMode }) {
  const [mode, setMode] = useState("login"); // login | register | reset
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [name, setName] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [showPass, setShowPass] = useState(false);

  const dm = darkMode;

  const errMap = {
    "auth/invalid-email": "البريد الإلكتروني غير صحيح",
    "auth/user-not-found": "البريد الإلكتروني غير مسجل",
    "auth/wrong-password": "كلمة المرور غير صحيحة",
    "auth/email-already-in-use": "هذا البريد مسجل مسبقاً",
    "auth/weak-password": "كلمة المرور يجب أن تكون 6 أحرف على الأقل",
    "auth/too-many-requests": "محاولات كثيرة، حاول لاحقاً",
    "auth/invalid-credential": "البريد أو كلمة المرور غير صحيحة",
  };

  const handleSubmit = async () => {
    setError(""); setSuccess(""); setLoading(true);
    try {
      if (mode === "login") {
        await signInWithEmailAndPassword(auth, email, password);
      } else if (mode === "register") {
        if (!name.trim()) { setError("يرجى إدخال الاسم"); setLoading(false); return; }
        const cred = await createUserWithEmailAndPassword(auth, email, password);
        await updateProfile(cred.user, { displayName: name });
      } else {
        await sendPasswordResetEmail(auth, email);
        setSuccess("تم إرسال رابط إعادة التعيين على بريدك الإلكتروني ✅");
      }
    } catch (err) {
      setError(errMap[err.code] || "حدث خطأ، حاول مرة أخرى");
    }
    setLoading(false);
  };

  const inp = {
    width:"100%", padding:"12px 16px", borderRadius:10,
    border:`1.5px solid ${dm?"#2a2f3d":"#d1d5db"}`,
    background:dm?"#1e222b":"#f9fafb", color:dm?"#f0f2f7":"#1f2937",
    fontSize:14, fontFamily:"inherit", direction:"rtl", outline:"none",
    boxSizing:"border-box", transition:"border 0.2s",
  };

  return (
    <div style={{minHeight:"100vh",background:dm?"#0d0f13":"#f5f0eb",display:"flex",alignItems:"center",justifyContent:"center",padding:16,fontFamily:"'Segoe UI',Tahoma,sans-serif",direction:"rtl"}}>
      <div style={{width:"100%",maxWidth:420}}>

        {/* شعار وعنوان */}
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{width:80,height:80,borderRadius:20,background:"linear-gradient(135deg,#6B1A1A,#F5A800)",margin:"0 auto 16px",display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"0 8px 24px rgba(107,26,26,0.3)"}}>
            <span style={{fontSize:36}}>🪪</span>
          </div>
          <h1 style={{fontSize:22,fontWeight:800,color:dm?"#f0f2f7":"#6B1A1A",margin:"0 0 6px"}}>نظام متابعة الإقامات</h1>
          <p style={{fontSize:13,color:dm?"#a0a8bb":"#6b7280",margin:0}}>شركة أنجال المشاعر · سلامة · مصاعد · كاميرات</p>
        </div>

        {/* البطاقة */}
        <div style={{background:dm?"#161920":"#fff",borderRadius:20,padding:"32px 28px",boxShadow:dm?"0 8px 32px rgba(0,0,0,0.5)":"0 8px 32px rgba(0,0,0,0.1)"}}>

          {/* عناوين الأوضاع */}
          <div style={{marginBottom:24}}>
            <h2 style={{fontSize:18,fontWeight:800,color:dm?"#f0f2f7":"#1e3a5f",margin:"0 0 4px"}}>
              {mode==="login"?"تسجيل الدخول":mode==="register"?"إنشاء حساب جديد":"استعادة كلمة المرور"}
            </h2>
            <p style={{fontSize:13,color:dm?"#a0a8bb":"#6b7280",margin:0}}>
              {mode==="login"?"أدخل بياناتك للدخول":mode==="register"?"أنشئ حسابك للوصول":"سنرسل لك رابط الاستعادة"}
            </p>
          </div>

          {/* حقل الاسم - للتسجيل فقط */}
          {mode==="register"&&(
            <div style={{marginBottom:14}}>
              <label style={{display:"block",fontSize:12,fontWeight:700,color:dm?"#a0a8bb":"#374151",marginBottom:6}}>الاسم الكامل</label>
              <input value={name} onChange={e=>setName(e.target.value)} placeholder="أدخل اسمك" style={inp}
                onFocus={e=>e.target.style.borderColor="#6B1A1A"} onBlur={e=>e.target.style.borderColor=dm?"#2a2f3d":"#d1d5db"}/>
            </div>
          )}

          {/* البريد الإلكتروني */}
          <div style={{marginBottom:14}}>
            <label style={{display:"block",fontSize:12,fontWeight:700,color:dm?"#a0a8bb":"#374151",marginBottom:6}}>البريد الإلكتروني</label>
            <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="example@email.com"
              style={{...inp,direction:"ltr",textAlign:"right"}}
              onFocus={e=>e.target.style.borderColor="#6B1A1A"} onBlur={e=>e.target.style.borderColor=dm?"#2a2f3d":"#d1d5db"}
              onKeyDown={e=>e.key==="Enter"&&handleSubmit()}/>
          </div>

          {/* كلمة المرور */}
          {mode!=="reset"&&(
            <div style={{marginBottom:20}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                <label style={{fontSize:12,fontWeight:700,color:dm?"#a0a8bb":"#374151"}}>كلمة المرور</label>
                {mode==="login"&&(
                  <button onClick={()=>{setMode("reset");setError("");setSuccess("");}}
                    style={{background:"none",border:"none",fontSize:12,color:"#6B1A1A",cursor:"pointer",fontFamily:"inherit"}}>
                    نسيت كلمة المرور؟
                  </button>
                )}
              </div>
              <div style={{position:"relative"}}>
                <input type={showPass?"text":"password"} value={password} onChange={e=>setPassword(e.target.value)}
                  placeholder="••••••••" style={{...inp,paddingLeft:44,direction:"ltr",textAlign:"right"}}
                  onFocus={e=>e.target.style.borderColor="#6B1A1A"} onBlur={e=>e.target.style.borderColor=dm?"#2a2f3d":"#d1d5db"}
                  onKeyDown={e=>e.key==="Enter"&&handleSubmit()}/>
                <button onClick={()=>setShowPass(!showPass)}
                  style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:16,color:dm?"#a0a8bb":"#6b7280"}}>
                  {showPass?"🙈":"👁️"}
                </button>
              </div>
            </div>
          )}

          {/* رسالة خطأ أو نجاح */}
          {error&&(
            <div style={{background:dm?"#2a1515":"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"10px 14px",marginBottom:16,fontSize:13,color:"#dc2626",display:"flex",alignItems:"center",gap:8}}>
              ⚠️ {error}
            </div>
          )}
          {success&&(
            <div style={{background:dm?"#081a12":"#f0fdf4",border:"1px solid #86efac",borderRadius:8,padding:"10px 14px",marginBottom:16,fontSize:13,color:"#16a34a",display:"flex",alignItems:"center",gap:8}}>
              ✅ {success}
            </div>
          )}

          {/* زر الإجراء */}
          <button onClick={handleSubmit} disabled={loading}
            style={{width:"100%",background:loading?"#9ca3af":"linear-gradient(135deg,#6B1A1A,#8B2500)",color:"#fff",border:"none",borderRadius:12,padding:"13px",fontWeight:800,fontSize:15,cursor:loading?"not-allowed":"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:8,transition:"opacity 0.2s"}}>
            {loading?"⏳ جارٍ التحميل...":mode==="login"?"🔐 دخول":mode==="register"?"✅ إنشاء الحساب":"📧 إرسال رابط الاستعادة"}
          </button>

          {/* روابط التبديل */}
          <div style={{textAlign:"center",marginTop:20,fontSize:13,color:dm?"#a0a8bb":"#6b7280"}}>
            {mode==="login"&&(
              <>ليس لديك حساب؟{" "}
                <button onClick={()=>{setMode("register");setError("");setSuccess("");}}
                  style={{background:"none",border:"none",color:"#6B1A1A",fontWeight:700,cursor:"pointer",fontFamily:"inherit",fontSize:13}}>
                  أنشئ حساباً
                </button>
              </>
            )}
            {mode==="register"&&(
              <>لديك حساب بالفعل؟{" "}
                <button onClick={()=>{setMode("login");setError("");setSuccess("");}}
                  style={{background:"none",border:"none",color:"#6B1A1A",fontWeight:700,cursor:"pointer",fontFamily:"inherit",fontSize:13}}>
                  سجّل دخول
                </button>
              </>
            )}
            {mode==="reset"&&(
              <button onClick={()=>{setMode("login");setError("");setSuccess("");}}
                style={{background:"none",border:"none",color:"#6B1A1A",fontWeight:700,cursor:"pointer",fontFamily:"inherit",fontSize:13}}>
                ← العودة لتسجيل الدخول
              </button>
            )}
          </div>
        </div>

        <p style={{textAlign:"center",marginTop:20,fontSize:11,color:dm?"#6b7280":"#9ca3af"}}>
          🔒 محمي بـ Firebase Authentication
        </p>
      </div>
    </div>
  );
}

export default function App() {
  const [user, setUser] = useState(undefined); // undefined=loading, null=لم يسجل, object=مسجل
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
  const [calcBacklogOnly, setCalcBacklogOnly] = useState(false);
  const [showFilterHelp, setShowFilterHelp] = useState(false);
  const [editCrModal, setEditCrModal]   = useState(null); // {compName} - تعديل السجل التجاري
  const [editCrData, setEditCrData]     = useState({});
  const [customCr, setCustomCr]         = useState(() => { try { return JSON.parse(localStorage.getItem("custom_cr")||"{}"); } catch{return{};} });
  const [editLinkModal, setEditLinkModal] = useState(null); // {id, linkIdx}
  const [editLinkData, setEditLinkData]   = useState({label:"", url:"", expiryDate:""});
  const [showExpiringDocs, setShowExpiringDocs] = useState(false);
  const [selectedReports, setSelectedReports] = useState(new Set(["status","company","nationality","jobs","location","calendar","costs","renewed"]));
  const [showExportReports, setShowExportReports] = useState(false); // وضع المتأخر فقط
  const [darkMode, setDarkMode] = useState(false);
  const [modalCard, setModalCard] = useState(null);
  const [renewModal, setRenewModal] = useState(null); // {record} أو null
  const [renewDate, setRenewDate] = useState("");
  const [renewNote, setRenewNote] = useState("");
  const [driveLinks, setDriveLinks] = useState(() => {
    try { return JSON.parse(localStorage.getItem("drive_links")||"{}"); } catch { return {}; }
  });
  const [driveModal, setDriveModal] = useState(null); // {type:"company"|"employee", id, name}
  const [driveLinkInput, setDriveLinkInput] = useState("");
  const [driveLinkExpiry, setDriveLinkExpiry] = useState("");
  const [driveLinkLabel, setDriveLinkLabel] = useState("");
  const [importModal, setImportModal] = useState(false);
  const [importResult, setImportResult] = useState(null); // {updated, notFound, rows}
  const [notifModal, setNotifModal] = useState(false);
  const [notifDays, setNotifDays] = useState(30); // null | {label, filter, fKey, color, icon}
  const LOGO = 'data:image/png;base64,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'; // {empId: true/false}

  // الاستماع لحالة تسجيل الدخول
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => {
      setUser(u || null);
    });
    return () => unsub();
  }, []);

  // ── تحميل البيانات من Firebase (تتزامن بين جميع الأجهزة) ──
  const [recordsLoaded, setRecordsLoaded]   = useState(false);
  const [driveLoaded, setDriveLoaded]       = useState(false);
  const [crLoaded, setCrLoaded]             = useState(false);
  const isFirstSync = useRef({records:true, drive:true, cr:true});

  useEffect(() => {
    if (!user) return;
    const recRef = ref(db, DB_PATH);
    const unsub = onValue(recRef, (snap) => {
      const data = snap.val();
      if (data && Array.isArray(data) && data.length > 0) {
        setRecords(data);
      } else if (data === null) {
        // أول استخدام - رفع البيانات الافتراضية
        try {
          const local = localStorage.getItem(STORAGE_KEY);
          setRecords(local ? JSON.parse(local) : INITIAL_DATA);
        } catch { setRecords(INITIAL_DATA); }
      }
      isFirstSync.current.records = false;
      setRecordsLoaded(true);
    });
    return () => unsub();
  }, [user]);

  useEffect(() => {
    if (!user) return;
    const dlRef = ref(db, "drive_links");
    const unsub = onValue(dlRef, (snap) => {
      const data = snap.val();
      if (data) setDriveLinks(data);
      isFirstSync.current.drive = false;
      setDriveLoaded(true);
    });
    return () => unsub();
  }, [user]);

  useEffect(() => {
    if (!user) return;
    const crRef = ref(db, "custom_cr");
    const unsub = onValue(crRef, (snap) => {
      const data = snap.val();
      if (data) setCustomCr(data);
      isFirstSync.current.cr = false;
      setCrLoaded(true);
    });
    return () => unsub();
  }, [user]);

  // ── حفظ التغييرات في Firebase + احتياطي محلي ──
  useEffect(() => {
    if (!recordsLoaded || records.length === 0) return;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(records));
    set(ref(db, DB_PATH), records).catch(()=>{});
  }, [records, recordsLoaded]);

  useEffect(() => {
    if (!driveLoaded) return;
    localStorage.setItem("drive_links", JSON.stringify(driveLinks));
    set(ref(db, "drive_links"), driveLinks).catch(()=>{});
  }, [driveLinks, driveLoaded]);

  useEffect(() => {
    if (!crLoaded) return;
    localStorage.setItem("custom_cr", JSON.stringify(customCr));
    set(ref(db, "custom_cr"), customCr).catch(()=>{});
  }, [customCr, crLoaded]);

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

  // ── استيراد من Excel (مقيم) ──
  const handleImportExcel = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, {type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {defval:""});
        let updated=0, notFound=[];
        const updatedRecords = [...records];

        rows.forEach(row => {
          // البحث عن رقم الإقامة في الأعمدة المحتملة
          const iqamaNum = String(
            row["رقم الاقامة"] || row["رقم الإقامة"] || row["iqama"] || row["ID"] || ""
          ).trim().replace(/\s/g,"");

          // تاريخ الانتهاء في الأعمدة المحتملة
          const rawDate = row["تاريخ انتهاء الاقامة"] || row["تاريخ انتهاء الإقامة"] ||
                          row["تاريخ الانتهاء"] || row["ExpiryDate"] || "";

          if (!iqamaNum) return;

          // تحويل التاريخ
          let expiryDate = "";
          if (rawDate) {
            if (typeof rawDate === "number") {
              // تاريخ Excel رقمي
              const d = new Date((rawDate - 25569) * 86400 * 1000);
              expiryDate = d.toISOString().slice(0,10);
            } else {
              const d = new Date(rawDate);
              if (!isNaN(d)) expiryDate = d.toISOString().slice(0,10);
            }
          }

          const idx = updatedRecords.findIndex(r => r.iqamaNumber === iqamaNum);
          if (idx !== -1 && expiryDate) {
            const oldDate = updatedRecords[idx].expiryDate;
            if (oldDate !== expiryDate) {
              updatedRecords[idx] = {
                ...updatedRecords[idx],
                expiryDate,
                renewalStatus: "مكتمل",
                lastRenewalDate: new Date().toISOString().slice(0,10),
                lastRenewalNote: "محدّث من ملف Excel",
              };
              updated++;
            }
          } else if (idx === -1) {
            notFound.push(iqamaNum);
          }
        });

        setRecords(updatedRecords);
        setImportResult({updated, notFound, total: rows.length});
      } catch(err) {
        alert("خطأ في قراءة الملف: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // ── طلب إذن الإشعارات ──
  const requestNotifications = async () => {
    if (!("Notification" in window)) {
      alert("متصفحك لا يدعم الإشعارات"); return;
    }
    const perm = await Notification.requestPermission();
    if (perm === "granted") {
      // جدولة إشعارات للإقامات القريبة من الانتهاء
      const soon = records.filter(r => {
        const d = getDaysLeft(r.expiryDate);
        return d >= 0 && d <= notifDays;
      });
      if (soon.length === 0) {
        new Notification("✅ متابعة الإقامات", {
          body: `لا توجد إقامات تنتهي خلال ${notifDays} يوم`,
          icon: "/icons/icon-192.png"
        });
      } else {
        soon.forEach(r => {
          new Notification("⚠️ إقامة تنتهي قريباً", {
            body: `${r.name} — تنتهي بعد ${getDaysLeft(r.expiryDate)} يوم`,
            icon: "/icons/icon-192.png"
          });
        });
      }
      setNotifModal(false);
      alert(`✅ تم تفعيل الإشعارات! سيتم إشعارك بـ ${soon.length} إقامة`);
    } else {
      alert("تم رفض إذن الإشعارات. يمكنك تفعيلها من إعدادات المتصفح.");
    }
  };

  // ── دوال تعديل السجل التجاري ──
  const getCrData = (compName) => ({ ...COMPANY_CR[compName], ...(customCr[compName]||{}) });

  const saveCrData = (compName, data) => {
    setCustomCr(prev => ({ ...prev, [compName]: data }));
    setEditCrModal(null);
  };

  // ── تعديل رابط Drive (اسم + تاريخ انتهاء) ──
  const updateDriveLink = (id, idx, newData) => {
    setDriveLinks(prev => ({
      ...prev,
      [id]: (prev[id]||[]).map((l,i) => i===idx ? {...l,...newData} : l)
    }));
    setEditLinkModal(null);
  };

  // ── المستندات التي تنتهي قريباً (خلال 90 يوم) ──
  const getExpiringDocs = () => {
    const results = [];
    const today = new Date();
    Object.entries(driveLinks).forEach(([id, links]) => {
      links.forEach((link, idx) => {
        if (!link.expiryDate) return;
        const exp = new Date(link.expiryDate);
        const days = Math.ceil((exp - today) / 86400000);
        if (days <= 90) {
          // إيجاد صاحب الملف
          const rec = records.find(r => r.id === Number(id));
          const compName = Object.keys({company_anjal:"انجال المشاعر",company_delta:"دلتا الماسية",company_smart:"البيوت الذكية"}).find(k => k===id);
          const ownerName = rec ? rec.name : compName ? {company_anjal:"انجال المشاعر",company_delta:"دلتا الماسية",company_smart:"البيوت الذكية"}[id] : id;
          results.push({ id, idx, link, days, ownerName });
        }
      });
    });
    return results.sort((a,b) => a.days - b.days);
  };

  // ── إشعار المستندات المنتهية ──
  const notifyExpiringDocs = async () => {
    const docs = getExpiringDocs();
    if (!("Notification" in window)) { alert("متصفحك لا يدعم الإشعارات"); return; }
    const perm = await Notification.requestPermission();
    if (perm === "granted") {
      if (docs.length === 0) {
        new Notification("✅ المستندات", { body: "لا توجد مستندات تنتهي خلال 90 يوم" });
      } else {
        docs.forEach(d => {
          new Notification(d.days < 0 ? "❌ مستند منتهي" : "⚠️ مستند قريب الانتهاء", {
            body: `${d.link.label} — ${d.ownerName} — ${d.days < 0 ? "منتهي منذ "+Math.abs(d.days)+" يوم" : "ينتهي بعد "+d.days+" يوم"}`,
          });
        });
      }
    }
  };

  // ── دوال Google Drive ──
  const getDriveLinks = (id) => driveLinks[id] || [];
  const addDriveLink  = (id) => {
    if (!driveLinkInput.trim()) return;
    const link = { url: driveLinkInput.trim(), label: driveLinkLabel.trim()||"ملف Drive", addedAt: new Date().toISOString().slice(0,10), expiryDate: driveLinkExpiry||"" };
    setDriveLinks(prev => ({ ...prev, [id]: [...(prev[id]||[]), link] }));
    setDriveLinkInput(""); setDriveLinkLabel(""); setDriveLinkExpiry("");
  };
  const removeDriveLink = (id, idx) => {
    setDriveLinks(prev => ({ ...prev, [id]: (prev[id]||[]).filter((_,i)=>i!==idx) }));
  };



  // ── متغيرات مشتركة ──
  const inp={padding:"9px 12px",border:`1px solid ${darkMode?"#2a2f3d":"#d1d5db"}`,borderRadius:"8px",fontSize:"13px",width:"100%",boxSizing:"border-box",fontFamily:"inherit",direction:"rtl",background:darkMode?"#161920":"#fff",color:darkMode?"#f0f2f7":"#1f2937",outline:"none"};

  // شاشة التحميل
  if (user === undefined) return (
    <div style={{minHeight:"100vh",background:darkMode?"#0d0f13":"#f5f0eb",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Segoe UI',Tahoma,sans-serif"}}>
      <div style={{textAlign:"center"}}><div style={{fontSize:48,marginBottom:12}}>🪪</div>
        <div style={{fontSize:16,fontWeight:700,color:darkMode?"#f0f2f7":"#6B1A1A"}}>جارٍ التحميل...</div>
      </div>
    </div>
  );

  // صفحة تسجيل الدخول
  if (!user) return <AuthScreen darkMode={darkMode}/>;

  // انتظار تحميل البيانات من Firebase
  if (!recordsLoaded) return (
    <div style={{minHeight:"100vh",background:darkMode?"#0d0f13":"#f5f0eb",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Segoe UI',Tahoma,sans-serif"}}>
      <div style={{textAlign:"center"}}>
        <div style={{fontSize:48,marginBottom:12}}>🪪</div>
        <div style={{fontSize:16,fontWeight:700,color:darkMode?"#f0f2f7":"#6B1A1A"}}>جارٍ تحميل البيانات...</div>
      </div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:darkMode?"#0d0f13":"#f0f4f8",fontFamily:"'Segoe UI',Tahoma,sans-serif",direction:"rtl",colorScheme:darkMode?"dark":"light"}}>
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
            {/* معلومات المستخدم + تسجيل الخروج */}
            <div style={{display:"flex",alignItems:"center",gap:8,background:"rgba(255,255,255,0.1)",borderRadius:10,padding:"6px 12px"}}>
              <div style={{textAlign:"right"}}>
                <div style={{fontSize:12,fontWeight:700,color:"#fff"}}>{user?.displayName||user?.email?.split("@")[0]}</div>
                <div style={{fontSize:10,color:"rgba(255,255,255,0.7)"}}>{user?.email}</div>
              </div>
              <button onClick={()=>signOut(auth)} title="تسجيل الخروج"
                style={{background:"rgba(255,255,255,0.15)",border:"1.5px solid rgba(255,255,255,0.4)",color:"#fff",borderRadius:8,width:32,height:32,fontSize:14,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>
                🚪
              </button>
            </div>

            {/* زر الاستيراد من Excel */}
            <label title="استيراد بيانات من ملف Excel (مقيم)"
              style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1.5px solid rgba(255,255,255,0.5)",borderRadius:9,padding:"8px 14px",fontWeight:600,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",gap:5}}>
              📥 استيراد
              <input type="file" accept=".xlsx,.xls" style={{display:"none"}}
                onChange={e=>{if(e.target.files[0]){handleImportExcel(e.target.files[0]);setImportModal(true);e.target.value="";}}}/>
            </label>

            {/* زر الإشعارات */}
            <button onClick={()=>setNotifModal(true)} title="تفعيل إشعارات انتهاء الإقامة"
              style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1.5px solid rgba(255,255,255,0.5)",borderRadius:9,padding:"8px 14px",fontWeight:600,fontSize:13,cursor:"pointer"}}>
              🔔
            </button>

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
                <div style={{position:"absolute",top:"110%",left:0,background:darkMode?"#161920":"#fff",borderRadius:10,boxShadow:darkMode?"0 8px 32px rgba(0,0,0,0.8)":"0 8px 24px rgba(0,0,0,0.15)",overflow:"hidden",minWidth:185,zIndex:100}}>
                  <button onClick={()=>{exportToExcel(records);setShowExportMenu(false);}}
                    style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"11px 15px",border:"none",background:"none",cursor:"pointer",fontSize:13,fontFamily:"inherit",textAlign:"right"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#f0fdf4"} onMouseLeave={e=>e.currentTarget.style.background="none"}>
                    <span style={{fontSize:17}}>📊</span><div><div style={{fontWeight:700,color:"#16a34a"}}>تصدير Excel</div><div style={{fontSize:11,color:"#6b7280"}}>3 أوراق: موظفون، مرافقون، ملخص</div></div>
                  </button>
                  <div style={{height:1,background:darkMode?"#2a2f3d":"#f3f4f6"}}/>
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

      <div style={{maxWidth:1200,margin:"0 auto",padding:"18px 14px",color:darkMode?"#f0f2f7":"inherit"}} onClick={()=>setShowExportMenu(false)}>
        {/* Stats Cards */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:10,marginBottom:18}}>
          {[
            {label:"الإجمالي",value:stats.total,color:darkMode?"#f0f2f7":"#1e3a5f",icon:"📋",filter:null},
            {label:"موظفون",value:stats.employees,color:darkMode?"#f0f2f7":"#374151",icon:"👤",filter:"موظف",fKey:"type"},
            {label:"مرافقون",value:stats.dependents,color:"#7c3aed",icon:"👨‍👩‍👧",filter:"مرافق",fKey:"type"},
            {label:"سارية",value:stats.valid,color:"#16a34a",icon:"✅",filter:"سارية",fKey:"status"},
            {label:"تنتهي قريباً",value:stats.soon,color:"#d97706",icon:"⚠️",filter:"تنتهي قريباً",fKey:"status"},
            {label:"منتهية",value:stats.expired,color:"#dc2626",icon:"❌",filter:"منتهية",fKey:"status"},
            {label:"انجال المشاعر",value:stats.anjal,color:"#6B1A1A",icon:"🏢",filter:"انجال المشاعر",fKey:"company"},
            {label:"دلتا الماسية",value:stats.delta,color:"#b45309",icon:"🏢",filter:"دلتا الماسية",fKey:"company"},
          ].map(s=>(
            <div key={s.label}
              onClick={()=>{ if(s.value>0) setModalCard(s); }}
              style={{background:darkMode?"#161920":"#fff",borderRadius:10,padding:"12px 10px",boxShadow:darkMode?"0 2px 6px rgba(0,0,0,0.6)":"0 2px 6px rgba(0,0,0,0.07)",borderTop:`3px solid ${s.color}`,textAlign:"center",cursor:s.value>0?"pointer":"default",transition:"all 0.15s",transform:"scale(1)",userSelect:"none"}}
              onMouseEnter={e=>{if(s.value>0)e.currentTarget.style.transform="scale(1.04) translateY(-2px)";}}
              onMouseLeave={e=>{e.currentTarget.style.transform="scale(1)";}}>
              <div style={{fontSize:20}}>{s.icon}</div>
              <div style={{fontSize:24,fontWeight:800,color:s.color}}>{s.value}</div>
              <div style={{fontSize:11,color:darkMode?"#a0a8bb":"#6b7280",marginTop:1}}>{s.label}</div>
            </div>
          ))}
        </div>

        {/* Tabs */}
        <div style={{display:"flex",gap:8,marginBottom:16}}>
          {["list","alerts","family","cost","reports","files","classify"].map(tab=>(
            <button key={tab} onClick={()=>setActiveTab(tab)}
              style={{padding:"7px 18px",borderRadius:8,border:"none",cursor:"pointer",fontWeight:600,fontSize:13,fontFamily:"inherit",
                background:activeTab===tab?"#8B2500":darkMode?"#161920":"#fff",color:activeTab===tab?"#fff":darkMode?"#f0f2f7":"#374151",border:darkMode&&activeTab!==tab?"1px solid #2a2f3d":"none",boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.6)":"0 2px 6px rgba(0,0,0,0.07)"}}>
              {tab==="list"?"📋 الكل":tab==="alerts"?`🔔 تنبيهات ${stats.expired+stats.soon>0?`(${stats.expired+stats.soon})`:""}`:tab==="family"?`👨‍👩‍👧 عائلات (${stats.dependents})`:tab==="cost"?"💰 حاسبة التكلفة":tab==="reports"?"📊 التقارير":tab==="files"?"📁 الملفات":"⚖️ التنصيف"}
            </button>
          ))}
        </div>

        {/* ALERTS */}
        {activeTab==="alerts"&&(
          <div style={{background:darkMode?"#161920":"#fff",borderRadius:14,padding:22,boxShadow:darkMode?"0 2px 10px rgba(0,0,0,0.6)":"0 2px 10px rgba(0,0,0,0.07)"}}>
            <h3 style={{margin:"0 0 14px",color:darkMode?"#f0f2f7":"#1e3a5f"}}>🔔 إقامات منتهية أو تنتهي قريباً</h3>
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
                <div key={head.id} style={{background:darkMode?"#161920":"#fff",borderRadius:14,padding:"18px 20px",marginBottom:14,boxShadow:darkMode?"0 2px 10px rgba(0,0,0,0.6)":"0 2px 10px rgba(0,0,0,0.07)",borderRight:`5px solid ${hSc.border}`}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8,marginBottom:12}}>
                    <div>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                        <span style={{fontSize:18}}>👤</span>
                        <strong style={{fontSize:16,color:darkMode?"#f0f2f7":"#1e3a5f"}}>{head.name}</strong>
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
                  <div style={{background:darkMode?"#1e222b":"#f8faff",borderRadius:10,padding:"10px 14px"}}>
                    <div style={{fontSize:12,fontWeight:700,color:darkMode?"#f0f2f7":"#374151",marginBottom:8}}>👨‍👩‍👧 المرافقون ({deps.length})</div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:8}}>
                      {deps.map(d=>(
                        <div key={d.id} style={{background:darkMode?"#161920":"#fff",borderRadius:8,padding:"8px 12px",border:`1px solid ${darkMode?"#2a2f3d":"#e0e7ff"}`,display:"flex",alignItems:"center",gap:8}}>
                          <span style={{fontSize:16}}>{RELATION_ICONS[d.relation]||"👤"}</span>
                          <div>
                            <div style={{fontSize:13,fontWeight:600,color:darkMode?"#f0f2f7":"#374151"}}>{d.name}</div>
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
            <div style={{background:darkMode?"#161920":"#fff",borderRadius:11,padding:"13px 16px",marginBottom:14,boxShadow:darkMode?"0 2px 7px rgba(0,0,0,0.6)":"0 2px 7px rgba(0,0,0,0.06)",display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
              {/* بحث */}
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 بحث بالاسم، الرقم، المهنة..."
                style={{...inp,flex:"1 1 180px",maxWidth:260}}/>

              {/* فلتر الشركة */}
              {(()=>{
                const isActive = filterCompany!=="الكل";
                return (
                  <select value={filterCompany} onChange={e=>setFilterCompany(e.target.value)}
                    style={{...inp,width:"auto",fontWeight:isActive?700:400,
                      border:`1.5px solid ${isActive?"#6B1A1A":(darkMode?"#2a2f3d":"#d1d5db")}`,
                      color:isActive?"#6B1A1A":(darkMode?"#f0f2f7":"#374151"),
                      background:isActive?(darkMode?"#2d0f0f":"#fdf0f0"):(darkMode?"#161920":"#fff")}}>
                    <option value="الكل">🏢 الشركة</option>
                    <option>انجال المشاعر</option>
                    <option>دلتا الماسية</option>
                    <option>البيوت الذكية</option>
                  </select>
                );
              })()}

              {/* فلتر النوع */}
              {(()=>{
                const isActive = filterType!=="الكل";
                return (
                  <select value={filterType} onChange={e=>setFilterType(e.target.value)}
                    style={{...inp,width:"auto",fontWeight:isActive?700:400,
                      border:`1.5px solid ${isActive?"#2563eb":(darkMode?"#2a2f3d":"#d1d5db")}`,
                      color:isActive?"#2563eb":(darkMode?"#f0f2f7":"#374151"),
                      background:isActive?(darkMode?"#0f1a2e":"#eff6ff"):(darkMode?"#161920":"#fff")}}>
                    <option value="الكل">👤 نوع السجل</option>
                    <option>موظف</option>
                    <option>مرافق</option>
                  </select>
                );
              })()}

              {/* فلتر الحالة */}
              {(()=>{
                const isActive = filterStatus!=="الكل";
                const statusColors = {"سارية":"#16a34a","تنتهي قريباً":"#d97706","منتهية":"#dc2626","قيد التجديد":"#2563eb","مرافق":"#7c3aed"};
                const c = statusColors[filterStatus]||"#16a34a";
                return (
                  <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)}
                    style={{...inp,width:"auto",fontWeight:isActive?700:400,
                      border:`1.5px solid ${isActive?c:(darkMode?"#2a2f3d":"#d1d5db")}`,
                      color:isActive?c:(darkMode?"#f0f2f7":"#374151"),
                      background:isActive?(darkMode?"#1a1a1a":"#f9fafb"):(darkMode?"#161920":"#fff")}}>
                    <option value="الكل">⏳ حالة الإقامة</option>
                    <option>سارية</option>
                    <option>تنتهي قريباً</option>
                    <option>منتهية</option>
                    <option>قيد التجديد</option>
                    <option>مرافق</option>
                  </select>
                );
              })()}

              {/* فلتر الجنسية */}
              {(()=>{
                const isActive = filterNationality!=="الكل";
                return (
                  <select value={filterNationality} onChange={e=>setFilterNationality(e.target.value)}
                    style={{...inp,width:"auto",fontWeight:isActive?700:400,
                      border:`1.5px solid ${isActive?"#0891b2":(darkMode?"#2a2f3d":"#d1d5db")}`,
                      color:isActive?"#0891b2":(darkMode?"#f0f2f7":"#374151"),
                      background:isActive?(darkMode?"#0a1f2e":"#f0f9ff"):(darkMode?"#161920":"#fff")}}>
                    <option value="الكل">🌍 الجنسية</option>
                    {nationalities.filter(n=>n!=="الكل").map(n=><option key={n}>{n}</option>)}
                  </select>
                );
              })()}

              {/* الترتيب */}
              <select value={sortBy} onChange={e=>setSortBy(e.target.value)}
                style={{...inp,width:"auto",color:darkMode?"#f0f2f7":"#374151"}}>
                <option value="expiryDate">↕️ تاريخ الانتهاء</option>
                <option value="name">↕️ الاسم</option>
                <option value="company">↕️ الشركة</option>
              </select>

              {/* إحصاء + إعادة ضبط */}
              <span style={{color:darkMode?"#a0a8bb":"#6b7280",fontSize:12,whiteSpace:"nowrap"}}>{filtered.length} سجل</span>
              {(filterCompany!=="الكل"||filterType!=="الكل"||filterStatus!=="الكل"||filterNationality!=="الكل"||search)&&(
                <button onClick={()=>{setFilterCompany("الكل");setFilterType("الكل");setFilterStatus("الكل");setFilterNationality("الكل");setSearch("");}}
                  style={{background:"#fee2e2",color:"#dc2626",border:"1px solid #fca5a5",borderRadius:7,padding:"3px 10px",fontSize:11,cursor:"pointer",fontFamily:"inherit",flexShrink:0,fontWeight:700}}>
                  ✕ مسح
                </button>
              )}
            </div>



            {showForm&&(
              <div style={{background:darkMode?"#161920":"#fff",borderRadius:14,padding:22,marginBottom:16,boxShadow:darkMode?"0 4px 20px rgba(0,0,0,0.7)":"0 4px 20px rgba(0,0,0,0.12)",border:`2px solid ${darkMode?"#F5A800":"#2563eb"}`}}>
                <h3 style={{margin:"0 0 16px",color:darkMode?"#f0f2f7":"#1e3a5f",fontSize:15}}>{editId?"✏️ تعديل":"➕ إضافة إقامة جديدة"}</h3>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:11}}>
                  {[{l:"الاسم *",k:"name",t:"text"},{l:"الجنسية",k:"nationality",t:"text"},{l:"رقم الإقامة *",k:"iqamaNumber",t:"text"},{l:"تاريخ انتهاء الإقامة",k:"expiryDate",t:"date"},{l:"رقم الجواز",k:"passportNumber",t:"text"},{l:"المهنة",k:"jobTitle",t:"text"},{l:"تكلفة التجديد (ريال)",k:"renewalCost",t:"number"},{l:"رقم إقامة رب الأسرة",k:"familyHeadId",t:"text"}].map(f=>(
                    <div key={f.k}>
                      <label style={{display:"block",marginBottom:3,fontSize:12,fontWeight:600,color:darkMode?"#d1d5db":"#374151"}}>{f.l}</label>
                      <input type={f.t} value={form[f.k]} onChange={e=>setForm({...form,[f.k]:e.target.value})} style={inp}/>
                    </div>
                  ))}
                  {[{l:"النوع",k:"type",o:["موظف","مرافق"]},{l:"صلة القرابة",k:"relation",o:["","زوجة","ابن","بنت"]},{l:"الجنس",k:"gender",o:["ذكر","أنثى"]},{l:"خارج المملكة",k:"outsideKingdom",o:["لا","نعم"]},{l:"حالة التجديد",k:"renewalStatus",o:["لم يبدأ","قيد التجديد","مكتمل"]},{l:"الشركة",k:"company",o:["انجال المشاعر","دلتا الماسية"]}].map(f=>(
                    <div key={f.k}>
                      <label style={{display:"block",marginBottom:3,fontSize:12,fontWeight:600,color:darkMode?"#f0f2f7":"#374151"}}>{f.l}</label>
                      <select value={form[f.k]} onChange={e=>setForm({...form,[f.k]:e.target.value})} style={inp}>{f.o.map(o=><option key={o}>{o}</option>)}</select>
                    </div>
                  ))}
                </div>
                <div style={{display:"flex",gap:8,marginTop:14}}>
                  <button onClick={handleSubmit} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:8,padding:"8px 22px",fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>{editId?"💾 حفظ":"✅ إضافة"}</button>
                  <button onClick={()=>{setShowForm(false);setEditId(null);setForm(emptyForm);}} style={{background:darkMode?"#1e222b":"#f3f4f6",color:darkMode?"#f0f2f7":"#374151",border:darkMode?"1px solid #2a2f3d":"none",borderRadius:8,padding:"8px 16px",fontWeight:600,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>إلغاء</button>
                </div>
              </div>
            )}

            {filtered.filter(r=>r.type!=="مرافق").length===0?(
              <div style={{textAlign:"center",background:darkMode?"#161920":"#fff",borderRadius:14,padding:50,boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.6)":"0 2px 8px rgba(0,0,0,0.06)"}}>
                <div style={{fontSize:44}}>🔍</div><h3 style={{color:darkMode?"#f0f2f7":"#374151",marginTop:8}}>لا توجد نتائج</h3>
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
                    <div key={r.id} style={{background:darkMode?"#161920":"#fff",borderRadius:12,boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.6)":"0 2px 8px rgba(0,0,0,0.07)",borderRight:`5px solid ${sc.border}`,overflow:"hidden"}}>
                      {/* ── صف رب الأسرة ── */}
                      <div style={{padding:"14px 18px"}}>
                        <div style={{display:"flex",flexWrap:"wrap",gap:10,alignItems:"center",justifyContent:"space-between"}}>
                          {/* الاسم والتفاصيل */}
                          <div style={{flex:"1 1 200px",cursor:"pointer"}} onClick={()=>setExpandedId(expanded?null:r.id)}>
                            <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:4,flexWrap:"wrap"}}>
                              <span style={{fontSize:17}}>{r.gender==="أنثى"?"👩":"👤"}</span>
                              <strong style={{fontSize:15,color:darkMode?"#f0f2f7":"#1e3a5f"}}>{r.name}</strong>
                              <span style={{background:cc.bg,color:cc.text,border:`1px solid ${cc.border}`,padding:"1px 8px",borderRadius:10,fontSize:11,fontWeight:600}}>{r.company}</span>
                              {r.outsideKingdom==="نعم"&&<span style={{background:"#fef3c7",color:"#d97706",padding:"1px 7px",borderRadius:10,fontSize:11}}>✈️ خارج</span>}
                            {r.passportExpiry&&(()=>{
                              const pd=Math.ceil((new Date(r.passportExpiry)-new Date())/86400000);
                              return pd<180?<span style={{background:pd<0?"#fee2e2":pd<90?"#fef3c7":"#dbeafe",color:pd<0?"#dc2626":pd<90?"#d97706":"#2563eb",padding:"1px 7px",borderRadius:10,fontSize:11}}>🛂 {pd<0?`جواز منتهي ${Math.abs(pd)} يوم`:`جواز ${pd} يوم`}</span>:null;
                            })()}
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
                                <div style={{fontSize:11,color:darkMode?"#6b7585":"#9ca3af"}}>{new Date(r.expiryDate).toLocaleDateString("ar-SA")}</div>
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
                              <button onClick={()=>{setRenewModal(r);setRenewDate("");setRenewNote("");}} title="تجديد الإقامة" style={{background:darkMode?"#1a2a1a":"#f0fdf4",color:"#16a34a",border:"1px solid #86efac",borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>🔄</button>
                              <button onClick={()=>{setDriveModal({type:"employee",id:r.id,name:r.name});setDriveLinkInput("");setDriveLinkLabel("");setDriveLinkExpiry("");}} title="ملفات Google Drive"
                                style={{background:darkMode?"#1a1a2e":"#f0f4ff",color:"#4285f4",border:"1px solid #93c5fd",borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontFamily:"inherit",position:"relative",display:"inline-flex",alignItems:"center",gap:3}}>
                                📁{getDriveLinks(r.id).length>0&&<span style={{background:"#4285f4",color:"#fff",borderRadius:10,padding:"0 5px",fontSize:9,fontWeight:800}}>{getDriveLinks(r.id).length}</span>}
                              </button>
                              <button onClick={()=>handleEdit(r)} style={{background:darkMode?"#1a1f2e":"#eff6ff",color:"#2563eb",border:"1px solid #bfdbfe",borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>✏️</button>
                              <button onClick={()=>handleDelete(r.id)} style={{background:darkMode?"#2a1515":"#fef2f2",color:"#dc2626",border:"1px solid #fecaca",borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>🗑️</button>
                            </div>
                          </div>
                        </div>
                        {/* تفاصيل رب الأسرة عند الضغط على الاسم */}
                        {expanded&&(
                          <div style={{marginTop:10,paddingTop:10,borderTop:"1px dashed #e5e7eb",display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:7}}>
                            {[{l:"رقم الإقامة",v:r.iqamaNumber},{l:"الجنسية",v:r.nationality},{l:"الجنس",v:r.gender},{l:"رقم الجواز",v:r.passportNumber},{l:"انتهاء الجواز",v:r.passportExpiry?new Date(r.passportExpiry).toLocaleDateString("ar-SA",{year:"numeric",month:"long",day:"numeric"}):"-"},{l:"تاريخ العقد",v:r.contractDate?new Date(r.contractDate).toLocaleDateString("ar-SA"):"-"},{l:"رقم العقد",v:r.contractNum||"-"},{l:"المهنة",v:r.jobTitle||r.notes},{l:"خارج المملكة",v:r.outsideKingdom},{l:"حالة التجديد",v:r.renewalStatus},{l:"الشركة",v:r.company},{l:"آخر تجديد",v:r.lastRenewalDate?new Date(r.lastRenewalDate).toLocaleDateString("ar-SA"):"-"},{l:"ملاحظة التجديد",v:r.lastRenewalNote||"-"}].map(f=>(
                              <div key={f.l} style={{background:darkMode?"#1e222b":"#f0f2f5",borderRadius:7,padding:"7px 10px"}}>
                                <div style={{fontSize:10,color:darkMode?"#a0a8bb":"#6b7280",marginBottom:1}}>{f.l}</div>
                                <div style={{fontSize:12,fontWeight:600,color:darkMode?"#f0f2f7":"#1f2937"}}>{f.v||"-"}</div>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>

                      {/* ── قسم المرافقين المنسدل ── */}
                      {expanded&&hasDependents&&(
                        <div style={{background:darkMode?"#13102a":"#faf5ff",borderTop:`2px dashed ${darkMode?"#5b21b6":"#c4b5fd"}`,padding:"14px 18px"}}>
                          <div style={{fontSize:13,fontWeight:700,color:"#7c3aed",marginBottom:10,display:"flex",alignItems:"center",gap:6}}>
                            👨‍👩‍👧 المرافقون ({dependents.length})
                          </div>
                          <div style={{display:"grid",gap:8}}>
                            {dependents.map(d=>{
                              const dcc=COMPANY_COLORS[d.company]||{bg:"#f9f9f9",text:"#374151",border:"#e5e7eb"};
                              return (
                                <div key={d.id} style={{background:darkMode?"#161920":"#fff",borderRadius:10,padding:"12px 16px",border:`1px solid ${darkMode?"#3b1f6e":"#e9d5ff"}`,display:"flex",flexWrap:"wrap",alignItems:"center",justifyContent:"space-between",gap:8}}>
                                  <div style={{flex:"1 1 180px"}}>
                                    <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:3,flexWrap:"wrap"}}>
                                      <span style={{fontSize:16}}>{RELATION_ICONS[d.relation]||"👤"}</span>
                                      <strong style={{fontSize:14,color:darkMode?"#d8b4fe":"#4c1d95"}}>{d.name}</strong>
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
                                    <button onClick={()=>{setRenewModal(d);setRenewDate("");setRenewNote("");}} title="تجديد الإقامة" style={{background:darkMode?"#1a2a1a":"#f0fdf4",color:"#16a34a",border:"1px solid #86efac",borderRadius:6,padding:"4px 9px",fontSize:11,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>🔄</button>
                                    <button onClick={()=>handleEdit(d)} style={{background:darkMode?"#1a1f2e":"#eff6ff",color:"#2563eb",border:"1px solid #bfdbfe",borderRadius:6,padding:"4px 9px",fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>✏️</button>
                                    <button onClick={()=>handleDelete(d.id)} style={{background:darkMode?"#2a1515":"#fef2f2",color:"#dc2626",border:"1px solid #fecaca",borderRadius:6,padding:"4px 9px",fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>🗑️</button>
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
        {activeTab==="cost"&&records.length>0&&(()=>{
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
            const isExpired  = days !== null && days < 0;
            const expiredDays = isExpired ? Math.abs(days) : 0;
            // أيام تبقت قبل الانتهاء (أقل من 90 يوم = داخل ربع حالي)
            const remainingDays = (!isExpired && days !== null) ? days : 0;
            const isWithinQuarter = !isExpired && remainingDays < 90; // تنتهي قبل نهاية الربع
            const useLateRate = isExpired && expiredDays > 3;

            // ── حساب المتأخرات (منتهية) ──
            const expiredFullQuarters  = Math.floor(expiredDays / 90);
            const remainingExpiredDays = expiredDays % 90;
            const partialExpiredCost   = remainingExpiredDays / 90;

            // ── حساب الربع الجاري (سارية لكن تنتهي قبل 90 يوم) ──
            // ندفع فقط الأيام المتبقية من الربع الحالي كـ"استكمال"
            const currentPartialCost = isWithinQuarter ? (remainingDays / 90) : 0;

            let backlogPassport=0, backlogWork=0, currentPassport=0, currentWork=0;
            let renewPassport=0, renewWork=0;

            if (isEmployee) {
              const rateQ = useLateRate ? PASSPORT_LATE_PER_Q : PASSPORT_NORMAL_PER_Q;

              // متأخرات (للمنتهية) - بحد أدنى ثابت
              if (isExpired) {
                // كل ربع (90 يوم) أو جزء منه = ربع كامل
                // 1-90 يوم = 2838، 91-180 = 5676، وهكذا
                const totalQuarters = Math.ceil(expiredDays / 90);
                backlogPassport = totalQuarters * PASSPORT_LATE_PER_Q;
                backlogWork     = totalQuarters * WORK_PERMIT_PER_Q;
              }

              if (calcBacklogOnly) {
                // ── وضع تلقائي ──
                if (isExpired) {
                  // منتهية → متأخرات فقط، لا تجديد
                } else {
                  // سارية أو تنتهي قريباً → تجديد 3 أشهر فقط بسعر عادي (بدون استكمال)
                  renewPassport = 1 * PASSPORT_NORMAL_PER_Q;
                  renewWork     = 1 * WORK_PERMIT_PER_Q;
                }
              } else {
                // ── وضع يدوي: تجديد بالمدة المختارة فقط (بدون استكمال) ──
                renewPassport = quarters * (isExpired ? rateQ : PASSPORT_NORMAL_PER_Q);
                renewWork     = quarters * WORK_PERMIT_PER_Q;
              }

            } else {
              // مرافق
              if (calcBacklogOnly) {
                if (!isExpired) {
                  renewPassport = 1 * PASSPORT_DEP_PER_Q;
                }
              } else {
                // تجديد بالمدة المختارة فقط (بدون استكمال)
                renewPassport = quarters * PASSPORT_DEP_PER_Q;
              }
            }

            const totalBacklog  = backlogPassport + backlogWork;
            const totalCurrent  = currentPassport + currentWork;
            const totalRenew    = renewPassport + renewWork;
            const total         = totalBacklog + totalCurrent + totalRenew;

            // نص توضيحي للعرض
            const expiredQuarterDisplay = isExpired
              ? (expiredFullQuarters > 0 ? `${expiredFullQuarters} ربع` : "") +
                (remainingExpiredDays > 0 ? `${expiredFullQuarters > 0 ? " + " : ""}${remainingExpiredDays} يوم` : "")
              : isWithinQuarter ? `${remainingDays} يوم متبقي` : "";

            return {r, isEmployee, isExpired, isWithinQuarter, expiredDays, remainingDays,
                    expiredFullQuarters, remainingExpiredDays, useLateRate,
                    backlogPassport, backlogWork, totalBacklog,
                    currentPassport, currentWork, totalCurrent,
                    renewPassport, renewWork, totalRenew,
                    total, expiredQuarterDisplay};
          };

          const breakdown      = targetRecords.map(calcRecord);
          const grandTotal     = breakdown.reduce((s,b)=>s+b.total,0);
          const grandBacklog   = breakdown.reduce((s,b)=>s+b.totalBacklog,0);
          const grandCurrent   = breakdown.reduce((s,b)=>s+b.totalCurrent,0);
          const grandRenew     = breakdown.reduce((s,b)=>s+b.totalRenew,0);
          const expiredCount   = breakdown.filter(b=>b.isExpired).length;
          const soonCount      = breakdown.filter(b=>b.isWithinQuarter).length;
          const fmt = n => n.toLocaleString("ar-SA") + " ريال";

          // helpers للاختيار
          const toggleEmp = id => {
            const next = new Set(hasSelection ? calcSelectedIds : new Set(allEmployees.map(e=>e.id)));
            if(next.has(id)) next.delete(id); else next.add(id);
            setCalcSelectedIds(next);
          };
          // تحديد الكل = فقط من يظهرون في القائمة الحالية (بعد الفلتر)
          const selectAll   = () => setCalcSelectedIds(new Set(sortedEmployees.map(e=>e.id)));
          // إلغاء الكل = إلغاء جميع الموظفين في كل القوائم
          const deselectAll = () => setCalcSelectedIds(new Set([-1]));
          const toggleDeps  = id => setCalcIncludeDeps(p=>({...p,[id]:p[id]===false?true:false}));

          return (
            <div>
              {/* ── شريط الإعدادات ── */}
              <div style={{background:darkMode?"#161920":"#fff",borderRadius:12,padding:"16px 20px",marginBottom:14,boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.6)":"0 2px 8px rgba(0,0,0,0.07)",display:"flex",gap:16,flexWrap:"wrap",alignItems:"flex-start"}}>
                <div>
                  <label style={{display:"block",fontSize:12,fontWeight:700,color:darkMode?"#a0a8bb":"#374151",marginBottom:6}}>📅 مدة التجديد</label>
                  <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
                    {[3,6,9,12].map(m=>(
                      <button key={m}
                        onClick={()=>{setCalcMonths(m);setCalcBacklogOnly(false);}}
                        style={{padding:"6px 14px",borderRadius:8,border:`2px solid ${!calcBacklogOnly&&calcMonths===m?"#2563eb":"#d1d5db"}`,background:!calcBacklogOnly&&calcMonths===m?"#2563eb":darkMode?"#1c1f26":"#fff",color:!calcBacklogOnly&&calcMonths===m?"#fff":darkMode?"#e8eaf0":"#374151",fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>
                        {m} شهر
                      </button>
                    ))}
                    <button
                      onClick={()=>setCalcBacklogOnly(!calcBacklogOnly)}
                      style={{padding:"6px 14px",borderRadius:8,border:`2px solid ${calcBacklogOnly?"#dc2626":"#fca5a5"}`,background:calcBacklogOnly?"#dc2626":darkMode?"#1c1f26":"#fff",color:calcBacklogOnly?"#fff":darkMode?"#e8eaf0":"#dc2626",fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:5}}>
                      🔄 تلقائي
                    </button>
                  </div>
                </div>
                <div>
                  <label style={{display:"block",fontSize:12,fontWeight:700,color:darkMode?"#a0a8bb":"#374151",marginBottom:6}}>🏢 الشركة</label>
                  <div style={{display:"flex",gap:7}}>
                    {[{k:"all",l:"الكل"},{k:"company1",l:"انجال المشاعر"},{k:"company2",l:"دلتا الماسية"}].map(c=>(
                      <button key={c.k} onClick={()=>{setCalcTarget(c.k);setCalcSelectedIds(new Set());setCalcIncludeDeps({});setCalcStatusFilter('all');}}
                        style={{padding:"6px 12px",borderRadius:8,border:`2px solid ${calcTarget===c.k?"#2563eb":"#d1d5db"}`,background:calcTarget===c.k?"#2563eb":"#fff",color:calcTarget===c.k?"#fff":"#374151",fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>
                        {c.l}
                      </button>
                    ))}
                  </div>
                </div>
                <div style={{marginRight:"auto",background:darkMode?"#081a12":"#f0fdf4",border:`1px solid ${darkMode?"#166534":"#86efac"}`,borderRadius:10,padding:"10px 18px",textAlign:"center",minWidth:190}}>
                  <div style={{fontSize:11,color:darkMode?"#6ee7b7":"#15803d",fontWeight:600}}>{calcBacklogOnly?"🔄 حساب تلقائي":`إجمالي التكلفة · ${calcMonths} شهر`}</div>
                  <div style={{fontSize:24,fontWeight:900,color:darkMode?"#6ee7b7":"#15803d"}}>{fmt(grandTotal)}</div>
                  <div style={{fontSize:11,color:darkMode?"#a0a8bb":"#6b7280",marginTop:2}}>{targetRecords.length} سجل ({targetEmployees.length} موظف + {targetDependents.length} مرافق)</div>
                </div>
              </div>

              {/* ── اختيار الموظفين مع خيار المرافقين ── */}
              <div style={{background:darkMode?"#161920":"#fff",borderRadius:12,padding:"16px 20px",marginBottom:14,boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.6)":"0 2px 8px rgba(0,0,0,0.06)"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
                  <div style={{fontWeight:700,fontSize:14,color:darkMode?"#f0f2f7":"#1e3a5f",display:"flex",alignItems:"center",gap:8}}>
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
                    {k:"all",    l:`الكل (${allEmployees.length})`,          bg:"#1c1f26", activeBg:"#374151",  border:"#4b5563",  activeBorder:"#9ca3af",  text:darkMode?"#a0a8bb":"#6b7280"},
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
                      <div key={emp.id} style={{borderRadius:10,border:`2px solid ${empSelected?"#F5A800":darkMode?"#2a2f3d":"#e5e7eb"}`,background:empSelected?(darkMode?"#1a1000":"#f0f9ff"):(darkMode?"#161920":"#fafafa"),overflow:"hidden",position:"relative"}}>
                        {/* شريط الأولوية */}
                        <div style={{position:"absolute",right:0,top:0,bottom:0,width:4,borderRadius:"0 10px 10px 0",background:priorityOrder(emp)===0?"#dc2626":priorityOrder(emp)===1?"#d97706":"#16a34a"}}/>
                        {/* صف الموظف */}
                        <div onClick={()=>toggleEmp(emp.id)}
                          style={{display:"flex",alignItems:"center",gap:10,padding:"9px 12px",cursor:"pointer"}}>
                          <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${empSelected?"#2563eb":"#d1d5db"}`,background:empSelected?"#2563eb":"#fff",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
                            {empSelected&&<span style={{color:"#fff",fontSize:11,fontWeight:900}}>✓</span>}
                          </div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontWeight:600,fontSize:12,color:darkMode?"#f0f2f7":"#1e3a5f",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{emp.name}</div>
                            <div style={{display:"flex",gap:5,marginTop:2,flexWrap:"wrap"}}>
                              <span style={{background:rcc.bg,color:rcc.text,padding:"0 5px",borderRadius:5,fontSize:10}}>{emp.company}</span>
                              <span style={{background:sc.bg,color:sc.text,padding:"0 5px",borderRadius:5,fontSize:10}}>{st}</span>
                              {emp.expiryDate&&(()=>{
                                const d=getDaysLeft(emp.expiryDate);
                                const col=d<0?"#dc2626":d<=30?"#d97706":"#16a34a";
                                const bg=d<0?"#fee2e2":d<=30?"#fef3c7":"#dcfce7";
                                return <span style={{background:bg,color:col,padding:"0 5px",borderRadius:5,fontSize:10,fontWeight:700}}>
                                  {d<0?`منتهية ${Math.abs(d)} يوم`:d===0?"ينتهي اليوم":`${d} يوم`}
                                </span>;
                              })()}
                            </div>
                          </div>
                        </div>
                        {/* خيار المرافقين — يظهر فقط إذا الموظف محدد وله مرافقون */}
                        {empSelected && deps.length > 0 && (
                          <div
                            onClick={e=>{e.stopPropagation();toggleDeps(emp.id);}}
                            style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"7px 12px",background:depsIncluded?(darkMode?"#13102a":"#f3e8ff"):(darkMode?"#1e222b":"#f9fafb"),borderTop:`1px dashed ${darkMode?"#5b21b6":"#e9d5ff"}`,cursor:"pointer",gap:8}}>
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
                  {label:"استكمال الربع الحالي",val:grandCurrent,color:"#d97706",bg:"#fef3c7",border:"#fcd34d",icon:"🕐"},
                  {label:"موظفون محددون",val:targetEmployees.length,color:darkMode?"#f0f2f7":"#374151",bg:"#f9fafb",border:"#e5e7eb",icon:"👤",isCount:true},
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
              <div style={{background:darkMode?"#161000":"#fefce8",border:`1px solid ${darkMode?"#92400e":"#fcd34d"}`,borderRadius:10,padding:"10px 16px",marginBottom:14,fontSize:11,color:darkMode?"#fde68a":"#78350f",lineHeight:1.9}}>
                <strong>📌 معادلة الحساب ({calcMonths} شهر = {quarters} ربع):</strong><br/>
                <span>👤 موظف سارية: رخصة عمل ({quarters}×2,425) + جواز ({quarters}×163) = {fmt(quarters*2425+quarters*163)}</span><br/>
                <span>👤 موظف منتهية &gt;3 أيام: نفس التجديد لكن الجواز بمعدل 413 ريال/ربع + متأخرات بنفس المعدل</span><br/>
                <span>👨‍👩‍👧 مرافق: رسوم جواز فقط ({quarters}×1,200={fmt(quarters*1200)})</span>
              </div>

              {/* ── جدول التفاصيل ── */}
              <div style={{background:darkMode?"#161920":"#fff",borderRadius:12,boxShadow:darkMode?"0 2px 8px rgba(0,0,0,0.6)":"0 2px 8px rgba(0,0,0,0.07)",overflow:"hidden"}}>
                <div style={{background:darkMode?"#3d1000":"#6B1A1A",padding:"10px 16px",display:"grid",gridTemplateColumns:"2fr 0.7fr 0.8fr 0.9fr 0.9fr 0.9fr 1fr",gap:6,color:"#fff",fontSize:11,fontWeight:700}}>
                  <span>الاسم</span>
                  <span style={{textAlign:"center"}}>النوع</span>
                  <span style={{textAlign:"center"}}>الحالة</span>
                  <span style={{textAlign:"center"}}>متأخرات</span>
                  <span style={{textAlign:"center"}}>استكمال</span>
                  <span style={{textAlign:"center"}}>{calcBacklogOnly?"تجديد (3ش تلقائي)":"تجديد ("+calcMonths+"ش)"}</span>
                  <span style={{textAlign:"center"}}>الإجمالي</span>
                </div>
                {breakdown.sort((a,b)=>priorityOrder(a.r)-priorityOrder(b.r)||getDaysLeft(a.r.expiryDate)-getDaysLeft(b.r.expiryDate)).map(({r,isEmployee,isExpired,isWithinQuarter,expiredDays,useLateRate,totalBacklog,totalCurrent,totalRenew,total,expiredQuarterDisplay},i)=>{
                  const st=getStatus(r),sc=STATUS_COLORS[st];
                  const rcc=COMPANY_COLORS[r.company]||{bg:"#f9f9f9",text:"#374151"};
                  return (
                    <div key={r.id} style={{padding:"9px 16px",borderBottom:`1px solid ${darkMode?"#2a2f3d":"#f3f4f6"}`,background:i%2===0?(darkMode?'#161920':'#fff'):(darkMode?'#111419':'#fafafa'),display:"grid",gridTemplateColumns:"2fr 0.7fr 0.8fr 0.9fr 0.9fr 0.9fr 1fr",gap:6,alignItems:"center"}}>
                      <div>
                        <div style={{fontWeight:600,fontSize:12,color:isEmployee?(darkMode?"#f0f2f7":"#1e3a5f"):(darkMode?"#e9d5ff":"#5b21b6")}}>{r.name}</div>
                        <div style={{fontSize:10,color:darkMode?"#a0a8bb":"#6b7280",display:"flex",gap:5,marginTop:2,flexWrap:"wrap"}}>
                          <span style={{background:rcc.bg,color:rcc.text,padding:"0 5px",borderRadius:5}}>{r.company}</span>
                          {isExpired&&<span style={{color:"#dc2626"}}>⚠️ {expiredQuarterDisplay} {useLateRate?"· غرامة":""}</span>}
                          {isWithinQuarter&&!isExpired&&<span style={{color:"#d97706"}}>🕐 {expiredQuarterDisplay}</span>}
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
                      <div style={{textAlign:"center",fontWeight:600,color:totalCurrent>0?"#d97706":"#9ca3af",fontSize:12}}>
                        {totalCurrent>0?fmt(totalCurrent):"—"}
                      </div>
                      <div style={{textAlign:"center",fontWeight:600,color:"#2563eb",fontSize:12}}>{fmt(totalRenew)}</div>
                      <div style={{textAlign:"center",fontWeight:800,color:"#15803d",fontSize:13}}>{fmt(total)}</div>
                    </div>
                  );
                })}
                <div style={{padding:"11px 16px",background:darkMode?"#081a12":"#f0fdf4",borderTop:`2px solid ${darkMode?"#166534":"#86efac"}`,display:"grid",gridTemplateColumns:"2.2fr 0.8fr 0.9fr 1.1fr 1.1fr 1.1fr",gap:6,alignItems:"center"}}>
                  <div style={{fontWeight:800,color:darkMode?"#6ee7b7":"#15803d",fontSize:13}}>الإجمالي ({targetRecords.length} سجل)</div>
                  <div/><div/>
                  <div style={{textAlign:"center",fontWeight:800,color:"#dc2626",fontSize:13}}>{grandBacklog>0?fmt(grandBacklog):"—"}</div>
                  <div style={{textAlign:"center",fontWeight:800,color:"#2563eb",fontSize:13}}>{fmt(grandRenew)}</div>
                  <div style={{textAlign:"center",fontWeight:900,color:"#15803d",fontSize:15}}>{fmt(grandTotal)}</div>
                </div>
              </div>
            </div>
          );
        })()}






        {/* ══════ تبويب التنصيف ══════ */}
        {activeTab==="classify"&&(()=>{
          const dm = darkMode;
          const card = {background:dm?"#161920":"#fff",borderRadius:14,padding:"20px 22px",boxShadow:dm?"0 2px 10px rgba(0,0,0,0.5)":"0 2px 10px rgba(0,0,0,0.07)",marginBottom:18};

          return (
            <div>
              {/* ── عنوان ── */}
              <div style={{background:dm?"#161920":"#fff",borderRadius:14,padding:"18px 22px",marginBottom:18,boxShadow:dm?"0 2px 10px rgba(0,0,0,0.5)":"0 2px 10px rgba(0,0,0,0.07)"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:10}}>
                  <div>
                    <div style={{fontWeight:800,fontSize:16,color:dm?"#f0f2f7":"#1e3a5f",marginBottom:6,display:"flex",alignItems:"center",gap:8}}>
                      ⚖️ تقرير الالتزام بالتنصيف
                    </div>
                    <p style={{fontSize:13,color:dm?"#a0a8bb":"#6b7280",margin:0}}>
                      مقارنة المهن الموجودة حالياً مع المتطلبات الرسمية لكل نشاط تجاري
                    </p>
                  </div>
                  <button onClick={()=>{
                    const today=new Date().toLocaleDateString("ar-SA",{year:"numeric",month:"long",day:"numeric"});
                    let sections="";
                    Object.entries(CLASSIFICATION_REQUIREMENTS).forEach(([compName,req])=>{
                      const compEmps=records.filter(r=>r.company===compName&&r.type!=="مرافق");
                      const jobCounts=req.jobs.map(j=>{
                        const count=compEmps.filter(e=>matchJob(e.jobTitle||e.notes,j.title)).length;
                        return {...j,count,diff:count-j.required};
                      });
                      const totalReq=jobCounts.reduce((s,j)=>s+j.required,0);
                      const totalFound=jobCounts.reduce((s,j)=>s+Math.min(j.count,j.required),0);
                      const compliance=totalReq>0?Math.round(totalFound/totalReq*100):0;
                      const statusColor=compliance===100?"#16a34a":compliance>=70?"#d97706":"#dc2626";
                      sections+=`
                        <div style="margin-bottom:28px;page-break-inside:avoid">
                          <div style="background:${req.color};color:#fff;padding:12px 18px;border-radius:10px 10px 0 0;display:flex;justify-content:space-between;align-items:center">
                            <div><strong style="font-size:16px">${req.icon} ${compName}</strong><div style="font-size:12px;opacity:.85;margin-top:2px">${req.activity}</div></div>
                            <div style="background:rgba(255,255,255,0.15);border-radius:8px;padding:8px 16px;text-align:center">
                              <div style="font-size:22px;font-weight:800">${compliance}%</div>
                              <div style="font-size:11px">${compliance===100?"✅ مكتمل":compliance>=70?"⚠️ جزئي":"❌ ناقص"}</div>
                            </div>
                          </div>
                          <table style="width:100%;border-collapse:collapse;border:1px solid #e5e7eb;border-top:none">
                            <thead><tr style="background:#f9fafb">
                              <th style="padding:9px 14px;text-align:right;font-size:12px;border-bottom:1px solid #e5e7eb">المهنة</th>
                              <th style="padding:9px 14px;text-align:center;font-size:12px;border-bottom:1px solid #e5e7eb">المطلوب</th>
                              <th style="padding:9px 14px;text-align:center;font-size:12px;border-bottom:1px solid #e5e7eb">الموجود</th>
                              <th style="padding:9px 14px;text-align:center;font-size:12px;border-bottom:1px solid #e5e7eb">الفرق</th>
                              <th style="padding:9px 14px;text-align:center;font-size:12px;border-bottom:1px solid #e5e7eb">الحالة</th>
                              <th style="padding:9px 14px;text-align:right;font-size:12px;border-bottom:1px solid #e5e7eb">الموظفون</th>
                            </tr></thead>
                            <tbody>
                            ${jobCounts.map((j,i)=>{
                              const emps=compEmps.filter(e=>matchJob(e.jobTitle||e.notes,j.title));
                              const c=j.diff>=0?"#16a34a":j.diff>=-1?"#d97706":"#dc2626";
                              const bg=i%2===0?"#fff":"#f9fafb";
                              return `<tr style="background:${bg}">
                                <td style="padding:8px 14px;font-weight:600;border-bottom:1px solid #f3f4f6;border-right:3px solid ${c}">${j.title}</td>
                                <td style="padding:8px 14px;text-align:center;font-weight:700;border-bottom:1px solid #f3f4f6">${j.required}</td>
                                <td style="padding:8px 14px;text-align:center;font-weight:800;color:${j.count>=j.required?"#16a34a":j.count>0?"#d97706":"#dc2626"};border-bottom:1px solid #f3f4f6">${j.count}</td>
                                <td style="padding:8px 14px;text-align:center;color:${c};font-weight:700;border-bottom:1px solid #f3f4f6">${j.diff>0?"+"+j.diff:j.diff}</td>
                                <td style="padding:8px 14px;text-align:center;border-bottom:1px solid #f3f4f6"><span style="background:${j.diff>=0?"#f0fdf4":j.diff>=-1?"#fefce8":"#fef2f2"};color:${c};padding:2px 10px;border-radius:10px;font-size:11px;font-weight:700">${j.diff>=0?"✅ مكتمل":j.diff>=-1?"⚠️ ناقص 1":"❌ ناقص "+Math.abs(j.diff)}</span></td>
                                <td style="padding:8px 14px;font-size:11px;color:#6b7280;border-bottom:1px solid #f3f4f6">${emps.map(e=>e.name).join(" · ")||"—"}</td>
                              </tr>`;
                            }).join("")}
                            <tr style="background:#f0fdf4;font-weight:800">
                              <td style="padding:9px 14px;color:#15803d">الإجمالي</td>
                              <td style="padding:9px 14px;text-align:center;color:#15803d">${req.jobs.reduce((s,j)=>s+j.required,0)}</td>
                              <td style="padding:9px 14px;text-align:center;color:#15803d">${jobCounts.reduce((s,j)=>s+j.count,0)}</td>
                              <td colspan="3" style="padding:9px 14px;text-align:center;color:${statusColor};font-size:13px">${compliance}% التزام</td>
                            </tr>
                            </tbody>
                          </table>
                        </div>`;
                    });
                    const html=`<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="UTF-8"><title>تقرير التنصيف</title>
                    <style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Segoe UI',Tahoma,sans-serif;direction:rtl;padding:24px;color:#1f2937;font-size:13px}
                    @media print{body{padding:12px}}</style></head><body>
                    <div style="background:linear-gradient(135deg,#6B1A1A,#F5A800);color:#fff;padding:20px 24px;border-radius:12px;margin-bottom:24px;display:flex;justify-content:space-between;align-items:center">
                      <div><h1 style="font-size:18px;font-weight:800">⚖️ تقرير الالتزام بالتنصيف</h1><p style="opacity:.85;margin-top:4px;font-size:12px">تاريخ التقرير: ${today}</p></div>
                      <div style="text-align:left;font-size:12px">إجمالي الموظفين: <strong>${records.filter(r=>r.type!=="مرافق").length}</strong></div>
                    </div>
                    ${sections}
                    <script>window.onload=()=>window.print();</script></body></html>`;
                    const win=window.open(URL.createObjectURL(new Blob([html],{type:"text/html;charset=utf-8"})),"_blank");
                    if(!win)alert("يرجى السماح بفتح النوافذ المنبثقة");
                  }}
                    style={{background:"#dc2626",color:"#fff",border:"none",borderRadius:10,padding:"9px 18px",fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
                    📄 تصدير PDF
                  </button>
                </div>
              </div>

              {/* ── بطاقة لكل شركة ── */}
              {Object.entries(CLASSIFICATION_REQUIREMENTS).map(([compName, req])=>{
                const compEmps = records.filter(r=>r.company===compName && r.type!=="مرافق");
                const cc = COMPANY_COLORS[compName]||{bg:"#f9f9f9",text:"#374151",border:"#e5e7eb"};

                // حساب الموجود من كل مهنة
                const jobCounts = req.jobs.map(j=>{
                  const count = compEmps.filter(e=>matchJob(e.jobTitle||e.notes, j.title)).length;
                  const diff  = count - j.required;
                  const status = diff >= 0 ? "مكتمل" : diff >= -1 ? "ناقص" : "ناقص جداً";
                  return {...j, count, diff, status};
                });

                const totalRequired = req.jobs.reduce((s,j)=>s+j.required,0);
                const totalFound    = jobCounts.reduce((s,j)=>s+j.count,0);
                const compliantJobs = jobCounts.filter(j=>j.diff>=0).length;
                const totalJobs     = jobCounts.length;
                // النسبة بناءً على عدد الموظفين الفعليين vs المطلوبين (أكثر دقة)
                const totalRequiredCount = jobCounts.reduce((s,j)=>s+j.required,0);
                const totalFoundCount    = jobCounts.reduce((s,j)=>s+Math.min(j.count,j.required),0);
                const compliance    = totalRequiredCount>0?Math.round(totalFoundCount/totalRequiredCount*100):0;

                // الموظفون غير المصنفين (مهنتهم لا تطابق أي متطلب)
                const classifiedIds = new Set();
                jobCounts.forEach(j=>{
                  compEmps.filter(e=>matchJob(e.jobTitle||e.notes, j.title)).forEach(e=>classifiedIds.add(e.id));
                });
                const unclassified = compEmps.filter(e=>!classifiedIds.has(e.id));

                const statusColor = compliance===100?"#16a34a":compliance>=70?"#d97706":"#dc2626";
                const statusBg    = dm?(compliance===100?"#081a12":compliance>=70?"#1a1500":"#2a1515"):(compliance===100?"#f0fdf4":compliance>=70?"#fefce8":"#fef2f2");

                return (
                  <div key={compName} style={card}>
                    {/* هيدر الشركة */}
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20,flexWrap:"wrap",gap:12}}>
                      <div style={{display:"flex",alignItems:"center",gap:12}}>
                        <div style={{width:44,height:44,borderRadius:12,background:cc.bg,border:`2px solid ${cc.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>
                          {req.icon}
                        </div>
                        <div>
                          <div style={{fontWeight:800,fontSize:16,color:dm?"#f0f2f7":req.color}}>{compName}</div>
                          <div style={{fontSize:12,color:dm?"#a0a8bb":"#6b7280",marginTop:2}}>{req.activity}</div>
                        </div>
                      </div>

                      {/* مؤشر الامتثال الدائري */}
                      <div style={{background:statusBg,border:`2px solid ${statusColor}`,borderRadius:12,padding:"12px 20px",textAlign:"center",minWidth:120}}>
                        <div style={{fontSize:28,fontWeight:900,color:statusColor}}>{compliance}%</div>
                        <div style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280",marginTop:2}}>نسبة الالتزام</div>
                        <div style={{fontSize:11,fontWeight:700,color:statusColor}}>
                          {compliance===100?"✅ مكتمل":compliance>=70?"⚠️ جزئي":"❌ ناقص"}
                        </div>
                      </div>
                    </div>

                    {/* شريط التقدم الكلي */}
                    <div style={{marginBottom:20}}>
                      <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:dm?"#a0a8bb":"#6b7280",marginBottom:6}}>
                        <span>التقدم الكلي ({compliantJobs}/{totalJobs} مهنة مكتملة)</span>
                        <span style={{fontWeight:700,color:statusColor}}>{compliance}%</span>
                      </div>
                      <div style={{background:dm?"#1e222b":"#f3f4f6",borderRadius:20,height:10,overflow:"hidden"}}>
                        <div style={{width:`${compliance}%`,height:"100%",background:statusColor,borderRadius:20,transition:"width 0.6s",minWidth:compliance>0?4:0}}/>
                      </div>
                    </div>

                    {/* جدول المهن */}
                    <div style={{borderRadius:12,overflow:"hidden",border:`1px solid ${dm?"#2a2f3d":"#e5e7eb"}`,marginBottom:unclassified.length>0?16:0}}>
                      {/* رأس الجدول */}
                      <div style={{background:req.color,padding:"10px 16px",display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 1.5fr",gap:8,color:"#fff",fontSize:12,fontWeight:700}}>
                        <span>المهنة المطلوبة</span>
                        <span style={{textAlign:"center"}}>المطلوب</span>
                        <span style={{textAlign:"center"}}>الموجود</span>
                        <span style={{textAlign:"center"}}>الفرق</span>
                        <span style={{textAlign:"center"}}>الحالة</span>
                      </div>

                      {/* صفوف المهن */}
                      {jobCounts.map((j,i)=>{
                        const rowBg = j.diff>=0
                          ? (dm?"#0a2218":"#f0fdf4")
                          : j.diff>=-1
                          ? (dm?"#1a1500":"#fefce8")
                          : (dm?"#2a1515":"#fef2f2");
                        const rowBorder = j.diff>=0?"#86efac":j.diff>=-1?"#fcd34d":"#fca5a5";
                        const textColor = j.diff>=0?"#16a34a":j.diff>=-1?"#d97706":"#dc2626";

                        return (
                          <div key={j.title} style={{
                            padding:"11px 16px",
                            background:i%2===0?(dm?"#161920":"#fff"):(dm?"#111419":"#fafafa"),
                            display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 1.5fr",gap:8,
                            alignItems:"center",
                            borderBottom:`1px solid ${dm?"#2a2f3d":"#f3f4f6"}`,
                            borderRight:`3px solid ${rowBorder}`,
                          }}>
                            <span style={{fontWeight:600,fontSize:13,color:dm?"#ffffff":"#374151"}}>{j.title}</span>
                            <span style={{textAlign:"center",fontWeight:700,fontSize:14,color:dm?"#a0a8bb":"#6b7280"}}>{j.required}</span>
                            <span style={{textAlign:"center",fontWeight:800,fontSize:14,color:j.count>=j.required?"#16a34a":j.count>0?"#d97706":"#dc2626"}}>{j.count}</span>
                            <span style={{textAlign:"center",fontWeight:700,fontSize:13,color:textColor}}>
                              {j.diff>0?`+${j.diff}`:j.diff}
                            </span>
                            <div style={{textAlign:"center"}}>
                              <span style={{background:rowBg,border:`1px solid ${rowBorder}`,color:textColor,padding:"3px 10px",borderRadius:12,fontSize:11,fontWeight:700}}>
                                {j.diff>=0?"✅ مكتمل":j.diff>=-1?"⚠️ ناقص 1":"❌ ناقص "+Math.abs(j.diff)}
                              </span>
                            </div>
                          </div>
                        );
                      })}

                      {/* صف الإجمالي */}
                      <div style={{padding:"11px 16px",background:dm?"#0a2218":"#f0fdf4",borderTop:`2px solid ${dm?"#166534":"#86efac"}`,display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 1.5fr",gap:8,alignItems:"center"}}>
                        <span style={{fontWeight:800,fontSize:13,color:dm?"#6ee7b7":"#15803d"}}>الإجمالي</span>
                        <span style={{textAlign:"center",fontWeight:800,fontSize:14,color:dm?"#6ee7b7":"#15803d"}}>{totalRequired}</span>
                        <span style={{textAlign:"center",fontWeight:800,fontSize:14,color:totalFound>=totalRequired?"#16a34a":"#d97706"}}>{totalFound}</span>
                        <span style={{textAlign:"center",fontWeight:700,fontSize:13,color:totalFound>=totalRequired?"#16a34a":"#dc2626"}}>
                          {totalFound-totalRequired>0?`+${totalFound-totalRequired}`:totalFound-totalRequired}
                        </span>
                        <div style={{textAlign:"center"}}>
                          <span style={{fontWeight:700,fontSize:12,color:statusColor}}>{compliance}% التزام</span>
                        </div>
                      </div>
                    </div>

                    {/* الموظفون المصنفون لكل مهنة */}
                    <div style={{marginBottom:unclassified.length>0?16:0}}>
                      <div style={{fontSize:13,fontWeight:700,color:dm?"#f0f2f7":"#374151",marginBottom:10}}>👥 توزيع الموظفين على المهن</div>
                      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(260px,1fr))",gap:8}}>
                        {jobCounts.map(j=>{
                          const emps = compEmps.filter(e=>matchJob(e.jobTitle||e.notes, j.title));
                          const isOk = j.diff>=0;
                          return (
                            <div key={j.title} style={{background:dm?"#1e222b":"#f9fafb",borderRadius:10,border:`1px solid ${isOk?(dm?"#166534":"#86efac"):(dm?"#7f1d1d":"#fca5a5")}`,padding:"10px 12px"}}>
                              <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
                                <span style={{fontSize:12,fontWeight:700,color:dm?"#ffffff":"#374151"}}>{j.title}</span>
                                <span style={{fontSize:11,fontWeight:700,color:isOk?"#16a34a":"#dc2626"}}>{j.count}/{j.required}</span>
                              </div>
                              {emps.length===0?(
                                <div style={{fontSize:11,color:dm?"#a0a8bb":"#9ca3af",fontStyle:"italic"}}>لا يوجد موظفون</div>
                              ):(
                                <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                                  {emps.map(e=>{
                                    const st=getStatus(e);
                                    const sc=STATUS_COLORS[st];
                                    return (
                                      <span key={e.id} style={{background:sc.bg,color:sc.text,border:`1px solid ${sc.border}`,padding:"1px 8px",borderRadius:10,fontSize:10,fontWeight:600}}>
                                        {e.name}
                                      </span>
                                    );
                                  })}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    {/* الموظفون غير المصنفين */}
                    {unclassified.length>0&&(
                      <div style={{background:dm?"#1a1500":"#fefce8",border:`1px solid ${dm?"#92400e":"#fcd34d"}`,borderRadius:10,padding:"12px 16px"}}>
                        <div style={{fontSize:13,fontWeight:700,color:dm?"#fde68a":"#78350f",marginBottom:8,display:"flex",alignItems:"center",gap:6}}>
                          ⚠️ موظفون خارج متطلبات التنصيف ({unclassified.length})
                        </div>
                        <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                          {unclassified.map(e=>(
                            <div key={e.id} style={{background:dm?"#161920":"#fff",border:`1px solid ${dm?"#92400e":"#fcd34d"}`,borderRadius:8,padding:"4px 10px",fontSize:11}}>
                              <span style={{fontWeight:600,color:dm?"#f0f2f7":"#374151"}}>{e.name}</span>
                              <span style={{color:dm?"#a0a8bb":"#6b7280",marginRight:4}}> — {e.jobTitle||e.notes||"بدون مهنة"}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          );
        })()}

        
        {/* ══════ النافذة المنبثقة للبطاقات ══════ */}
        {modalCard && (() => {
          const getModalRecords = () => {
            if (!modalCard.fKey) return records;
            if (modalCard.fKey === "status")  return records.filter(r => getStatus(r) === modalCard.filter);
            if (modalCard.fKey === "company") return records.filter(r => r.company === modalCard.filter);
            if (modalCard.fKey === "type")    return records.filter(r => modalCard.filter === "مرافق" ? r.type === "مرافق" : r.type !== "مرافق");
            return [];
          };
          const modalRecords = getModalRecords().sort((a,b) => {
            const pa = getStatus(a)==="منتهية"?0:getStatus(a)==="تنتهي قريباً"?1:2;
            const pb = getStatus(b)==="منتهية"?0:getStatus(b)==="تنتهي قريباً"?1:2;
            return pa-pb || getDaysLeft(a.expiryDate)-getDaysLeft(b.expiryDate);
          });
          const dm = darkMode;
          return (
            <div onClick={()=>setModalCard(null)}
              style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16,backdropFilter:"blur(3px)"}}>
              <div onClick={e=>e.stopPropagation()}
                style={{background:dm?"#161920":"#fff",borderRadius:16,width:"100%",maxWidth:820,maxHeight:"88vh",display:"flex",flexDirection:"column",boxShadow:"0 24px 64px rgba(0,0,0,0.4)",overflow:"hidden"}}>
                <div style={{background:"linear-gradient(135deg,#3d1000,#6B1A1A)",padding:"18px 24px",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
                  <div style={{display:"flex",alignItems:"center",gap:12}}>
                    <span style={{fontSize:28}}>{modalCard.icon}</span>
                    <div>
                      <h2 style={{margin:0,fontSize:18,fontWeight:800,color:"#fff"}}>{modalCard.label}</h2>
                      <div style={{fontSize:12,color:"rgba(255,255,255,0.8)",marginTop:2}}>{modalRecords.length} سجل</div>
                    </div>
                  </div>
                  <button onClick={()=>setModalCard(null)}
                    style={{background:"rgba(255,255,255,0.15)",border:"1.5px solid rgba(255,255,255,0.4)",color:"#fff",borderRadius:8,width:34,height:34,fontSize:18,cursor:"pointer",fontWeight:700}}>✕</button>
                </div>
                <div style={{overflowY:"auto",padding:"16px 20px",flex:1}}>
                  {modalRecords.length === 0 ? (
                    <div style={{textAlign:"center",padding:40,color:dm?"#a0a8bb":"#6b7280"}}>
                      <div style={{fontSize:44}}>📭</div><p style={{marginTop:8}}>لا توجد سجلات</p>
                    </div>
                  ) : modalRecords.map(r => {
                    const st=getStatus(r), sc=STATUS_COLORS[st];
                    const days=r.expiryDate?getDaysLeft(r.expiryDate):null;
                    const cc=COMPANY_COLORS[r.company]||{bg:"#f9f9f9",text:"#374151",border:"#e5e7eb"};
                    const headRecord=r.type==="مرافق"?records.find(e=>e.iqamaNumber===r.familyHeadId):null;
                    return (
                      <div key={r.id} style={{background:dm?"#252830":"#f9fafb",borderRadius:12,padding:"14px 16px",marginBottom:10,borderRight:`4px solid ${sc.border}`,display:"flex",flexWrap:"wrap",gap:12,alignItems:"center",justifyContent:"space-between"}}>
                        <div>
                          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:5}}>
                            <strong style={{fontSize:14,color:dm?"#ffffff":"#1e3a5f"}}>{r.name}</strong>
                            {r.type==="مرافق"&&<span style={{background:"#f3e8ff",color:"#7c3aed",padding:"1px 8px",borderRadius:8,fontSize:11,fontWeight:600}}>{r.relation}</span>}
                            <span style={{background:cc.bg,color:cc.text,border:`1px solid ${cc.border}`,padding:"1px 7px",borderRadius:8,fontSize:11,fontWeight:600}}>{r.company}</span>
                            {r.outsideKingdom==="نعم"&&<span style={{background:"#fef3c7",color:"#d97706",padding:"1px 7px",borderRadius:8,fontSize:11}}>✈️ خارج</span>}
                          </div>
                          <div style={{display:"flex",gap:12,fontSize:12,color:dm?"#a0a8bb":"#6b7280",flexWrap:"wrap"}}>
                            <span>🪪 {r.iqamaNumber}</span>
                            {r.nationality&&<span>🌍 {r.nationality}</span>}
                            {r.jobTitle&&r.type!=="مرافق"&&<span>💼 {r.jobTitle}</span>}
                            {headRecord&&<span style={{color:dm?"#c4b5fd":"#7c3aed"}}>👤 {headRecord.name}</span>}
                          </div>
                        </div>
                        <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>
                          {days!==null&&(
                            <div style={{textAlign:"center"}}>
                              <div style={{fontWeight:800,fontSize:20,color:sc.text,lineHeight:1}}>{days<0?Math.abs(days):days}</div>
                              <div style={{fontSize:10,color:dm?"#a0a8bb":"#6b7280"}}>{days<0?"يوم منتهي":"يوم متبقي"}</div>
                              <div style={{fontSize:11,color:dm?"#a0a8bb":"#9ca3af"}}>{new Date(r.expiryDate).toLocaleDateString("ar-SA")}</div>
                            </div>
                          )}
                          <span style={{background:sc.bg,border:`1px solid ${sc.border}`,color:sc.text,padding:"3px 10px",borderRadius:14,fontSize:12,fontWeight:700}}>{st}</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
                <div style={{padding:"12px 20px",borderTop:`1px solid ${dm?"#2a2f3d":"#e5e7eb"}`,display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0,background:dm?"#161920":"#f9fafb"}}>
                  <span style={{fontSize:12,color:dm?"#a0a8bb":"#6b7280"}}>اضغط خارج النافذة للإغلاق</span>
                  <button onClick={()=>setModalCard(null)}
                    style={{background:"#8B2500",color:"#fff",border:"none",borderRadius:8,padding:"7px 20px",fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>إغلاق</button>
                </div>
              </div>
            </div>
          );
        })()}


        {/* ══════ نافذة تعديل الرابط ══════ */}
        {editLinkModal&&(()=>{
          const inpStyle={padding:"9px 12px",border:`1px solid ${darkMode?"#2a2f3d":"#d1d5db"}`,borderRadius:"8px",fontSize:"13px",width:"100%",boxSizing:"border-box",fontFamily:"inherit",direction:"rtl",background:darkMode?"#161920":"#fff",color:darkMode?"#f0f2f7":"#1f2937",outline:"none"};
          return (
          <div onClick={()=>setEditLinkModal(null)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",zIndex:1300,display:"flex",alignItems:"center",justifyContent:"center",padding:16,backdropFilter:"blur(3px)"}}>
            <div onClick={e=>e.stopPropagation()} style={{background:darkMode?"#161920":"#fff",borderRadius:16,width:"100%",maxWidth:460,boxShadow:"0 24px 64px rgba(0,0,0,0.4)",overflow:"hidden"}}>
              <div style={{background:"linear-gradient(135deg,#1a73e8,#4285f4)",padding:"16px 22px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div style={{color:"#fff",fontWeight:800,fontSize:16}}>✏️ تعديل الرابط</div>
                <button onClick={()=>setEditLinkModal(null)} style={{background:"rgba(255,255,255,0.15)",border:"1.5px solid rgba(255,255,255,0.4)",color:"#fff",borderRadius:8,width:32,height:32,fontSize:16,cursor:"pointer",fontWeight:700}}>✕</button>
              </div>
              <div style={{padding:"20px 22px"}}>
                <div style={{marginBottom:12}}>
                  <label style={{display:"block",fontSize:12,fontWeight:600,color:darkMode?"#a0a8bb":"#6b7280",marginBottom:5}}>اسم الملف</label>
                  <input value={editLinkModal.link.label} onChange={e=>setEditLinkModal(p=>({...p,link:{...p.link,label:e.target.value}}))}
                    style={{...inpStyle}}/>
                </div>
                <div style={{marginBottom:12}}>
                  <label style={{display:"block",fontSize:12,fontWeight:600,color:darkMode?"#a0a8bb":"#6b7280",marginBottom:5}}>رابط Google Drive</label>
                  <input value={editLinkModal.link.url} onChange={e=>setEditLinkModal(p=>({...p,link:{...p.link,url:e.target.value}}))}
                    style={{...inpStyle,direction:"ltr",textAlign:"right"}}/>
                </div>
                <div style={{marginBottom:20}}>
                  <label style={{display:"block",fontSize:12,fontWeight:600,color:darkMode?"#a0a8bb":"#6b7280",marginBottom:5}}>📅 تاريخ انتهاء المستند</label>
                  <input type="date" value={editLinkModal.link.expiryDate||""} onChange={e=>setEditLinkModal(p=>({...p,link:{...p.link,expiryDate:e.target.value}}))}
                    style={{...inpStyle,direction:"ltr"}}/>
                </div>
                <div style={{display:"flex",gap:10}}>
                  <button onClick={()=>{updateDriveLink(editLinkModal.id,editLinkModal.index,editLinkModal.link);setEditLinkModal(null);}}
                    style={{flex:1,background:"#1a73e8",color:"#fff",border:"none",borderRadius:10,padding:"10px",fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>
                    💾 حفظ التعديل
                  </button>
                  <button onClick={()=>setEditLinkModal(null)}
                    style={{background:darkMode?"#1e222b":"#f3f4f6",color:darkMode?"#f0f2f7":"#374151",border:`1px solid ${darkMode?"#2a2f3d":"#e5e7eb"}`,borderRadius:10,padding:"10px 16px",fontWeight:600,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>
                    إلغاء
                  </button>
                </div>
              </div>
            </div>
          </div>
          );
        })()}

        {/* ══════ نافذة تعديل السجل التجاري ══════ */}
        {editCrModal&&(()=>{
          const inpStyle={padding:"9px 12px",border:`1px solid ${darkMode?"#2a2f3d":"#d1d5db"}`,borderRadius:"8px",fontSize:"13px",width:"100%",boxSizing:"border-box",fontFamily:"inherit",direction:"rtl",background:darkMode?"#161920":"#fff",color:darkMode?"#f0f2f7":"#1f2937",outline:"none"};
          return (
          <div onClick={()=>setEditCrModal(null)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",zIndex:1300,display:"flex",alignItems:"center",justifyContent:"center",padding:16,backdropFilter:"blur(3px)"}}>
            <div onClick={e=>e.stopPropagation()} style={{background:darkMode?"#161920":"#fff",borderRadius:16,width:"100%",maxWidth:520,boxShadow:"0 24px 64px rgba(0,0,0,0.4)",overflow:"hidden",maxHeight:"85vh",display:"flex",flexDirection:"column"}}>
              <div style={{background:"linear-gradient(135deg,#6B1A1A,#F5A800)",padding:"16px 22px",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
                <div style={{color:"#fff"}}>
                  <div style={{fontWeight:800,fontSize:16}}>✏️ تعديل السجل التجاري</div>
                  <div style={{fontSize:12,opacity:.85,marginTop:2}}>{editCrModal.compName}</div>
                </div>
                <button onClick={()=>setEditCrModal(null)} style={{background:"rgba(255,255,255,0.15)",border:"1.5px solid rgba(255,255,255,0.4)",color:"#fff",borderRadius:8,width:32,height:32,fontSize:16,cursor:"pointer",fontWeight:700}}>✕</button>
              </div>
              <div style={{padding:"20px 22px",overflowY:"auto",flex:1}}>
                {[
                  {k:"name",          l:"الاسم الرسمي للشركة"},
                  {k:"crNumber",       l:"رقم السجل التجاري"},
                  {k:"unifiedNumber",  l:"الرقم الوطني الموحد"},
                  {k:"establishNumber",l:"رقم المنشأة"},
                  {k:"type",           l:"نوع الكيان"},
                  {k:"status",         l:"حالة السجل"},
                  {k:"address",        l:"العنوان"},
                  {k:"email",          l:"البريد الإلكتروني"},
                  {k:"manager",        l:"المدير العام"},
                  {k:"facilityManager",l:"مدير المنشأة"},
                  {k:"phone",          l:"رقم الهاتف"},
                  {k:"issueDate",      l:"تاريخ الإصدار",  type:"date"},
                  {k:"expiryDate",     l:"تاريخ انتهاء السجل", type:"date"},
                ].map(f=>(
                  <div key={f.k} style={{marginBottom:12}}>
                    <label style={{display:"block",fontSize:12,fontWeight:600,color:darkMode?"#a0a8bb":"#6b7280",marginBottom:4}}>{f.l}</label>
                    <input type={f.type||"text"} value={editCrModal.data[f.k]||""} onChange={e=>setEditCrModal(p=>({...p,data:{...p.data,[f.k]:e.target.value}}))}
                      style={{...inpStyle,direction:f.type==="date"?"ltr":"rtl"}}/>
                  </div>
                ))}
              </div>
              <div style={{padding:"14px 22px",borderTop:`1px solid ${darkMode?"#2a2f3d":"#e5e7eb"}`,display:"flex",gap:10,flexShrink:0}}>
                <button onClick={()=>{saveCrData(editCrModal.compName,editCrModal.data);setEditCrModal(null);}}
                  style={{flex:1,background:"linear-gradient(135deg,#6B1A1A,#8B2500)",color:"#fff",border:"none",borderRadius:10,padding:"11px",fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>
                  💾 حفظ التعديلات
                </button>
                <button onClick={()=>setEditCrModal(null)}
                  style={{background:darkMode?"#1e222b":"#f3f4f6",color:darkMode?"#f0f2f7":"#374151",border:`1px solid ${darkMode?"#2a2f3d":"#e5e7eb"}`,borderRadius:10,padding:"11px 16px",fontWeight:600,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>
                  إلغاء
                </button>
              </div>
            </div>
          </div>
          );
        })()}

{/* ══════ نافذة Google Drive ══════ */}
        {driveModal && (
          <div onClick={()=>setDriveModal(null)}
            style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",zIndex:1200,display:"flex",alignItems:"center",justifyContent:"center",padding:16,backdropFilter:"blur(3px)"}}>
            <div onClick={e=>e.stopPropagation()}
              style={{background:darkMode?"#161920":"#fff",borderRadius:16,width:"100%",maxWidth:520,boxShadow:"0 24px 64px rgba(0,0,0,0.4)",overflow:"hidden",maxHeight:"85vh",display:"flex",flexDirection:"column"}}>

              {/* هيدر */}
              <div style={{background:"linear-gradient(135deg,#1a73e8,#4285f4)",padding:"18px 24px",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
                <div style={{color:"#fff"}}>
                  <div style={{fontWeight:800,fontSize:17,display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:22}}>📁</span>
                    ملفات Google Drive
                  </div>
                  <div style={{fontSize:12,opacity:.85,marginTop:3}}>{driveModal.name}</div>
                </div>
                <button onClick={()=>setDriveModal(null)}
                  style={{background:"rgba(255,255,255,0.15)",border:"1.5px solid rgba(255,255,255,0.4)",color:"#fff",borderRadius:8,width:32,height:32,fontSize:16,cursor:"pointer",fontWeight:700}}>✕</button>
              </div>

              <div style={{padding:"20px 24px",overflowY:"auto",flex:1}}>
                {/* الروابط الموجودة */}
                {getDriveLinks(driveModal.id).length===0?(
                  <div style={{textAlign:"center",padding:"24px 0",color:darkMode?"#a0a8bb":"#6b7280"}}>
                    <div style={{fontSize:44,marginBottom:8}}>📂</div>
                    <p style={{fontSize:13}}>لا توجد ملفات مرتبطة بعد</p>
                    <p style={{fontSize:11,marginTop:4}}>أضف رابط من Google Drive أدناه</p>
                  </div>
                ):(
                  <div style={{marginBottom:16}}>
                    <div style={{fontSize:13,fontWeight:700,color:darkMode?"#f0f2f7":"#374151",marginBottom:10}}>
                      الملفات المرتبطة ({getDriveLinks(driveModal.id).length})
                    </div>
                    <div style={{display:"grid",gap:8}}>
                      {getDriveLinks(driveModal.id).map((link,i)=>(
                        <div key={i} style={{background:darkMode?"#1e222b":"#f8faff",border:`1px solid ${darkMode?"#2a2f3d":"#dbeafe"}`,borderRadius:10,padding:"11px 14px",display:"flex",alignItems:"center",gap:10}}>
                          <span style={{fontSize:20,flexShrink:0}}>
                            {link.label.includes("جواز")||link.label.includes("passport")?"🛂":
                             link.label.includes("عقد")?"📝":
                             link.label.includes("سجل")?"🏢":
                             link.label.includes("صورة")||link.label.includes("photo")?"🖼️":
                             link.label.includes("ترخيص")?"📜":"📄"}
                          </span>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontWeight:600,fontSize:13,color:darkMode?"#f0f2f7":"#1e3a5f"}}>{link.label}</div>
                            <div style={{fontSize:10,color:darkMode?"#a0a8bb":"#6b7280",marginTop:2}}>أُضيف: {link.addedAt}</div>
                          </div>
                          <div style={{display:"flex",gap:6,flexShrink:0}}>
                            <a href={link.url} target="_blank" rel="noreferrer"
                              style={{background:"#1a73e8",color:"#fff",padding:"5px 12px",borderRadius:7,fontSize:12,fontWeight:700,textDecoration:"none",display:"flex",alignItems:"center",gap:4}}>
                              🔗 فتح
                            </a>
                            <button onClick={()=>removeDriveLink(driveModal.id,i)}
                              style={{background:darkMode?"#2a1515":"#fef2f2",color:"#dc2626",border:"1px solid #fecaca",borderRadius:7,padding:"5px 10px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>
                              🗑️
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* إضافة رابط جديد */}
                <div style={{background:darkMode?"#1e222b":"#f9fafb",border:`1px solid ${darkMode?"#2a2f3d":"#e5e7eb"}`,borderRadius:12,padding:"16px"}}>
                  <div style={{fontSize:13,fontWeight:700,color:darkMode?"#f0f2f7":"#374151",marginBottom:12,display:"flex",alignItems:"center",gap:6}}>
                    ➕ إضافة رابط Drive جديد
                  </div>
                  <div style={{marginBottom:10}}>
                    <label style={{display:"block",fontSize:11,fontWeight:600,color:darkMode?"#a0a8bb":"#6b7280",marginBottom:4}}>اسم الملف</label>
                    <input value={driveLinkLabel} onChange={e=>setDriveLinkLabel(e.target.value)}
                      placeholder="مثال: جواز السفر، عقد العمل، السجل التجاري..."
                      style={{...inp}}/>
                  </div>
                  <div style={{marginBottom:12}}>
                    <label style={{display:"block",fontSize:11,fontWeight:600,color:darkMode?"#a0a8bb":"#6b7280",marginBottom:4}}>رابط Google Drive</label>
                    <input value={driveLinkInput} onChange={e=>setDriveLinkInput(e.target.value)}
                      placeholder="https://drive.google.com/..."
                      style={{...inp,direction:"ltr",textAlign:"right"}}/>
                  </div>
                  <div style={{marginBottom:12}}>
                    <label style={{display:"block",fontSize:11,fontWeight:600,color:darkMode?"#a0a8bb":"#6b7280",marginBottom:4}}>📅 تاريخ انتهاء المستند <span style={{fontWeight:400,color:darkMode?"#6b7585":"#9ca3af"}}>(اختياري)</span></label>
                    <input type="date" value={driveLinkExpiry} onChange={e=>setDriveLinkExpiry(e.target.value)}
                      style={{...inp,direction:"ltr"}}/>
                  </div>

                  <div style={{marginBottom:12}}>
                    <label style={{display:"block",fontSize:11,fontWeight:600,color:darkMode?"#a0a8bb":"#6b7280",marginBottom:4}}>📅 تاريخ انتهاء المستند (اختياري)</label>
                    <input type="date" value={driveLinkExpiry||""} onChange={e=>setDriveLinkExpiry(e.target.value)} style={{...inp}}/>
                    {driveLinkExpiry&&(()=>{
                      const d=Math.ceil((new Date(driveLinkExpiry)-new Date())/86400000);
                      return <div style={{fontSize:11,marginTop:4,color:d<0?"#dc2626":d<=30?"#d97706":"#16a34a",fontWeight:600}}>
                        {d<0?`❌ منتهي منذ ${Math.abs(d)} يوم`:d<=30?`⚠️ ينتهي بعد ${d} يوم`:`✅ ينتهي بعد ${d} يوم`}
                      </div>;
                    })()}
                  </div>

                  {/* تلميح كيف تنسخ الرابط */}
                  <div style={{background:darkMode?"#161000":"#fefce8",border:`1px solid ${darkMode?"#92400e":"#fcd34d"}`,borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:11,color:darkMode?"#fde68a":"#78350f"}}>
                    💡 <strong>كيف تحصل على الرابط؟</strong> افتح الملف في Drive → كليك يمين → "نسخ الرابط" أو "المشاركة"
                  </div>

                  <button onClick={()=>addDriveLink(driveModal.id)} disabled={!driveLinkInput.trim()}
                    style={{width:"100%",background:driveLinkInput.trim()?"#1a73e8":"#d1d5db",color:"#fff",border:"none",borderRadius:10,padding:"10px",fontWeight:700,fontSize:14,cursor:driveLinkInput.trim()?"pointer":"not-allowed",fontFamily:"inherit"}}>
                    📎 إضافة الرابط
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}


        {/* ══════ تبويب الملفات ══════ */}
        {activeTab==="files"&&(()=>{
          const dm = darkMode;
          const card = {background:dm?"#161920":"#fff",borderRadius:14,padding:"20px 22px",boxShadow:dm?"0 2px 10px rgba(0,0,0,0.5)":"0 2px 10px rgba(0,0,0,0.07)",marginBottom:16};

          const COMPANY_IDS = {
            "انجال المشاعر": "company_anjal",
            "دلتا الماسية":  "company_delta",
            "البيوت الذكية": "company_smart",
          };
          const COMPANY_DOCS = [
            "السجل التجاري","الترخيص التجاري","رخصة البلدية","شهادة الزكاة",
            "عقد التأسيس","الهوية الضريبية","رخصة المنشأة","وثيقة تأمين",
          ];
          const SMART_HOME_DOCS = [
            "السجل التجاري","شهادة العضوية","مزاولة النشاط",
            "رخصة المراقبة الأمنية","شهادة ضريبة القيمة المضافة",
          ];
          const EMP_DOCS = ["جواز السفر","تصريح العمل","عقد العمل","صورة شخصية","شهادة مهنية","تأشيرة","إقامة (نسخة)","أخرى"];

          return (
            <div>
              {/* ── قائمة المستندات المنتهية قريباً ── */}
              {(()=>{
                const expiring = getExpiringDocs();
                if(expiring.length===0) return null;
                return (
                  <div style={{background:dm?"#1a1500":"#fefce8",border:`1px solid ${dm?"#92400e":"#fcd34d"}`,borderRadius:12,padding:"14px 18px",marginBottom:16}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
                      <div style={{fontWeight:800,fontSize:14,color:dm?"#fde68a":"#78350f",display:"flex",alignItems:"center",gap:8}}>
                        ⚠️ مستندات تنتهي قريباً
                        <span style={{background:dm?"#2a1500":"#fef3c7",color:dm?"#fde68a":"#d97706",padding:"1px 10px",borderRadius:10,fontSize:12,fontWeight:700}}>{expiring.length}</span>
                      </div>
                      <button onClick={notifyExpiringDocs}
                        style={{background:"#d97706",color:"#fff",border:"none",borderRadius:8,padding:"5px 14px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:700,display:"flex",alignItems:"center",gap:5}}>
                        🔔 إشعار
                      </button>
                    </div>
                    <div style={{display:"grid",gap:6}}>
                      {expiring.map((d,i)=>{
                        const c=d.days<0?"#dc2626":d.days<30?"#d97706":"#d97706";
                        const bg=dm?(d.days<0?"#2a1515":"#1a1200"):(d.days<0?"#fef2f2":"#fefce8");
                        return (
                          <div key={i} style={{background:bg,border:`1px solid ${d.days<0?"#fca5a5":"#fcd34d"}`,borderRadius:9,padding:"9px 12px",display:"flex",alignItems:"center",gap:10,justifyContent:"space-between",flexWrap:"wrap"}}>
                            <div style={{display:"flex",alignItems:"center",gap:8}}>
                              <span style={{fontSize:18}}>{d.link.label.includes("جواز")?"🛂":d.link.label.includes("عقد")?"📝":d.link.label.includes("سجل")?"🏢":d.link.label.includes("ترخيص")?"📜":"📄"}</span>
                              <div>
                                <div style={{fontWeight:700,fontSize:13,color:dm?"#f0f2f7":"#374151"}}>{d.link.label}</div>
                                <div style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280"}}>{d.owner?.name||d.compName||"—"}</div>
                              </div>
                            </div>
                            <div style={{display:"flex",alignItems:"center",gap:8}}>
                              <span style={{fontWeight:700,fontSize:12,color:c,background:dm?"rgba(0,0,0,0.3)":"rgba(255,255,255,0.7)",padding:"2px 9px",borderRadius:8}}>
                                {d.days<0?`منتهي منذ ${Math.abs(d.days)} يوم`:`${d.days} يوم`}
                              </span>
                              <a href={d.link.url} target="_blank" rel="noreferrer"
                                style={{background:"#1a73e8",color:"#fff",padding:"3px 9px",borderRadius:6,fontSize:11,fontWeight:700,textDecoration:"none"}}>🔗</a>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })()}

              {/* ── ملفات الشركتين ── */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:16,marginBottom:0}}>
                {Object.entries(COMPANY_IDS).map(([compName, compId])=>{
                  const cc=COMPANY_COLORS[compName]||{bg:"#f9fafb",text:"#374151",border:"#e5e7eb"};
                  const links=getDriveLinks(compId);
                  return (
                    <div key={compId} style={card}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                        <div style={{display:"flex",alignItems:"center",gap:10}}>
                          <div style={{width:38,height:38,borderRadius:10,background:cc.bg,border:`2px solid ${cc.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>🏢</div>
                          <div>
                            <div style={{fontWeight:800,fontSize:15,color:dm?"#f0f2f7":cc.text}}>{compName}</div>
                            <div style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280"}}>{links.length} ملف مرتبط</div>
                          </div>
                        </div>
                        <button onClick={()=>{setDriveModal({type:"company",id:compId,name:compName});setDriveLinkInput("");setDriveLinkLabel("");setDriveLinkExpiry("");}}
                          style={{background:"#1a73e8",color:"#fff",border:"none",borderRadius:9,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:5}}>
                          ➕ إضافة ملف
                        </button>
                      </div>

                      {/* اقتراحات الملفات */}
                      {/* بيانات السجل التجاري */}
                      {getCrData(compName)&&(
                        <div style={{background:dm?"#1e222b":"#f9fafb",borderRadius:10,padding:"12px 14px",marginBottom:12,border:`1px solid ${dm?"#2a2f3d":"#e5e7eb"}`}}>
                          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                            <div style={{fontSize:12,fontWeight:700,color:dm?"#f0f2f7":"#374151"}}>📋 بيانات السجل التجاري</div>
                            <button onClick={()=>setEditCrModal({compName,data:{...getCrData(compName)}})}
                              style={{background:dm?"#1e222b":"#f3f4f6",color:dm?"#f0f2f7":"#374151",border:`1px solid ${dm?"#2a2f3d":"#e5e7eb"}`,borderRadius:7,padding:"3px 10px",fontSize:11,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>
                              ✏️ تعديل
                            </button>
                          </div>
                          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"4px 12px",fontSize:11}}>
                            {[
                              {l:"الاسم الرسمي",v:getCrData(compName).name},
                              {l:"السجل التجاري",v:getCrData(compName).crNumber||"—"},
                              {l:"الرقم الموحد",v:getCrData(compName).unifiedNumber},
                              {l:"نوع الكيان",v:getCrData(compName).type},
                              {l:"الحالة",v:getCrData(compName).status},
                              {l:"المدير العام",v:getCrData(compName).manager||"—"},
                              {l:"مدير المنشأة",v:getCrData(compName).facilityManager||"—"},
                            ].map(x=>(
                              <div key={x.l}>
                                <span style={{color:dm?"#a0a8bb":"#6b7280"}}>{x.l}: </span>
                                <span style={{fontWeight:600,color:dm?"#f0f2f7":"#374151"}}>{x.v}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}

                      <div style={{marginBottom:12}}>
                        <div style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280",marginBottom:6}}>ملفات مقترحة:</div>
                        <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
                          {(compName==="البيوت الذكية"?SMART_HOME_DOCS:COMPANY_DOCS).map(doc=>{
                            const exists=links.some(l=>l.label===doc);
                            return (
                              <span key={doc} style={{background:exists?(dm?"#081a12":"#f0fdf4"):(dm?"#1e222b":"#f9fafb"),color:exists?"#16a34a":(dm?"#a0a8bb":"#6b7280"),border:`1px solid ${exists?"#86efac":(dm?"#2a2f3d":"#e5e7eb")}`,padding:"3px 9px",borderRadius:12,fontSize:11,fontWeight:exists?700:400}}>
                                {exists?"✅ ":""}{doc}
                              </span>
                            );
                          })}
                        </div>
                      </div>

                      {/* الملفات المضافة */}
                      {links.length===0?(
                        <div style={{textAlign:"center",padding:"16px 0",color:dm?"#a0a8bb":"#9ca3af",fontSize:12}}>
                          📂 لم يتم ربط أي ملفات بعد
                        </div>
                      ):(
                        <div style={{display:"grid",gap:6}}>
                          {links.map((link,i)=>(
                            <div key={i} style={{background:dm?"#1e222b":"#f8faff",border:`1px solid ${dm?"#2a2f3d":"#dbeafe"}`,borderRadius:9,padding:"9px 12px",display:"flex",alignItems:"center",gap:8}}>
                              <span style={{fontSize:16}}>📄</span>
                              <div style={{flex:1}}>
                                <div style={{fontWeight:600,fontSize:12,color:dm?"#f0f2f7":"#1e3a5f"}}>{link.label}</div>
                                <div style={{fontSize:10,color:dm?"#a0a8bb":"#6b7280"}}>أُضيف: {link.addedAt}</div>
                              </div>
                              <a href={link.url} target="_blank" rel="noreferrer"
                                style={{background:"#1a73e8",color:"#fff",padding:"4px 10px",borderRadius:7,fontSize:11,fontWeight:700,textDecoration:"none"}}>
                                🔗 فتح
                              </a>
                              <button onClick={()=>removeDriveLink(compId,i)}
                                style={{background:"none",border:"none",color:"#dc2626",cursor:"pointer",fontSize:14,padding:"2px 4px"}}>🗑️</button>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* ── ملفات الموظفين ── */}
              <div style={card}>
                <div style={{fontWeight:800,fontSize:15,color:dm?"#f0f2f7":"#1e3a5f",marginBottom:4,display:"flex",alignItems:"center",gap:8}}>
                  👤 ملفات الموظفين والمرافقين
                  <span style={{background:dm?"#1e222b":"#f3f4f6",color:dm?"#a0a8bb":"#6b7280",padding:"2px 10px",borderRadius:10,fontSize:12,fontWeight:600}}>
                    {records.filter(r=>getDriveLinks(r.id).length>0).length} لديهم ملفات
                  </span>
                </div>
                <p style={{fontSize:12,color:dm?"#a0a8bb":"#6b7280",marginBottom:16}}>
                  لإضافة ملف لموظف، اضغط زر 📁 في بطاقته بتبويب "القائمة"
                </p>

                {/* فلتر الملفات */}
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:10}}>
                  {records.filter(r=>getDriveLinks(r.id).length>0).map(r=>{
                    const links=getDriveLinks(r.id);
                    const cc=COMPANY_COLORS[r.company]||{bg:"#f9fafb",text:"#374151",border:"#e5e7eb"};
                    return (
                      <div key={r.id} style={{background:dm?"#1e222b":"#f9fafb",borderRadius:12,border:`1px solid ${dm?"#2a2f3d":"#e5e7eb"}`,padding:"14px",display:"flex",flexDirection:"column",gap:10}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                          <div>
                            <div style={{fontWeight:700,fontSize:13,color:dm?"#f0f2f7":"#1e3a5f"}}>{r.name}</div>
                            <div style={{display:"flex",gap:6,marginTop:4,flexWrap:"wrap"}}>
                              <span style={{background:cc.bg,color:cc.text,border:`1px solid ${cc.border}`,padding:"1px 7px",borderRadius:6,fontSize:10}}>{r.company}</span>
                              <span style={{background:dm?"#1c1f26":"#f3f4f6",color:dm?"#a0a8bb":"#6b7280",padding:"1px 7px",borderRadius:6,fontSize:10}}>🪪 {r.iqamaNumber}</span>
                            </div>
                          </div>
                          <button onClick={()=>{setDriveModal({type:"employee",id:r.id,name:r.name});setDriveLinkInput("");setDriveLinkLabel("");setDriveLinkExpiry("");}}
                            style={{background:"#1a73e8",color:"#fff",border:"none",borderRadius:8,padding:"5px 10px",fontSize:11,cursor:"pointer",fontFamily:"inherit",fontWeight:700,flexShrink:0}}>
                            ➕
                          </button>
                        </div>
                        <div style={{display:"grid",gap:5}}>
                          {links.map((link,i)=>(
                            <div key={i} style={{background:dm?"#161920":"#fff",border:`1px solid ${dm?"#2a2f3d":"#e5e7eb"}`,borderRadius:8,padding:"7px 10px",display:"flex",alignItems:"center",gap:7}}>
                              <span style={{fontSize:14}}>
                                {link.label.includes("جواز")?"🛂":link.label.includes("عقد")?"📝":link.label.includes("صورة")?"🖼️":"📄"}
                              </span>
                              <div style={{flex:1,fontSize:11,fontWeight:600,color:dm?"#e8eaf0":"#374151"}}>{link.label}</div>
                              <a href={link.url} target="_blank" rel="noreferrer"
                                style={{background:"#1a73e8",color:"#fff",padding:"3px 9px",borderRadius:6,fontSize:11,fontWeight:700,textDecoration:"none",flexShrink:0}}>
                                🔗
                              </a>
                            </div>
                          ))}
                        </div>
                      </div>
                    );
                  })}
                  {records.filter(r=>getDriveLinks(r.id).length>0).length===0&&(
                    <div style={{gridColumn:"1/-1",textAlign:"center",padding:"40px 0",color:dm?"#a0a8bb":"#9ca3af"}}>
                      <div style={{fontSize:44,marginBottom:8}}>📁</div>
                      <p>لم يتم ربط ملفات Drive لأي موظف بعد</p>
                      <p style={{fontSize:12,marginTop:4}}>اضغط زر 📁 في بطاقة الموظف لإضافة ملفاته</p>
                    </div>
                  )}
                </div>
              </div>
            </div>
          );
        })()}

        {/* ══════ نافذة نتيجة الاستيراد ══════ */}
        {importModal && importResult && (
          <div onClick={()=>setImportModal(false)}
            style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",zIndex:1200,display:"flex",alignItems:"center",justifyContent:"center",padding:16,backdropFilter:"blur(3px)"}}>
            <div onClick={e=>e.stopPropagation()}
              style={{background:darkMode?"#161920":"#fff",borderRadius:16,width:"100%",maxWidth:440,boxShadow:"0 24px 64px rgba(0,0,0,0.4)",overflow:"hidden"}}>

              <div style={{background:"linear-gradient(135deg,#1e40af,#3b82f6)",padding:"18px 24px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div style={{color:"#fff"}}>
                  <div style={{fontWeight:800,fontSize:17}}>📥 نتيجة الاستيراد</div>
                  <div style={{fontSize:12,opacity:.85,marginTop:3}}>من ملف Excel</div>
                </div>
                <button onClick={()=>setImportModal(false)}
                  style={{background:"rgba(255,255,255,0.15)",border:"1.5px solid rgba(255,255,255,0.4)",color:"#fff",borderRadius:8,width:32,height:32,fontSize:16,cursor:"pointer",fontWeight:700}}>✕</button>
              </div>

              <div style={{padding:"24px"}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:20}}>
                  {[
                    {label:"إجمالي الصفوف",val:importResult.total,color:"#374151",bg:darkMode?"#1e222b":"#f9fafb"},
                    {label:"تم تحديثه",val:importResult.updated,color:"#16a34a",bg:darkMode?"#081a12":"#f0fdf4"},
                    {label:"غير موجود",val:importResult.notFound.length,color:importResult.notFound.length>0?"#dc2626":"#16a34a",bg:importResult.notFound.length>0?(darkMode?"#2a1515":"#fef2f2"):(darkMode?"#081a12":"#f0fdf4")},
                  ].map(c=>(
                    <div key={c.label} style={{background:c.bg,borderRadius:10,padding:"14px 10px",textAlign:"center"}}>
                      <div style={{fontSize:26,fontWeight:800,color:c.color}}>{c.val}</div>
                      <div style={{fontSize:11,color:darkMode?"#a0a8bb":"#6b7280",marginTop:3}}>{c.label}</div>
                    </div>
                  ))}
                </div>

                {importResult.updated > 0 && (
                  <div style={{background:darkMode?"#081a12":"#f0fdf4",border:`1px solid ${darkMode?"#166534":"#86efac"}`,borderRadius:8,padding:"10px 14px",marginBottom:12,fontSize:13,color:darkMode?"#6ee7b7":"#15803d",fontWeight:600}}>
                    ✅ تم تحديث {importResult.updated} إقامة تلقائياً وتغيير حالتها إلى "مكتمل"
                  </div>
                )}

                {importResult.notFound.length > 0 && (
                  <div style={{background:darkMode?"#2a1515":"#fef2f2",border:`1px solid ${darkMode?"#7f1d1d":"#fca5a5"}`,borderRadius:8,padding:"10px 14px",marginBottom:12,fontSize:12,color:"#dc2626"}}>
                    <strong>أرقام غير موجودة في النظام:</strong>
                    <div style={{marginTop:6,display:"flex",flexWrap:"wrap",gap:6}}>
                      {importResult.notFound.slice(0,8).map(n=>(
                        <span key={n} style={{background:"#fee2e2",padding:"2px 8px",borderRadius:6}}>{n}</span>
                      ))}
                      {importResult.notFound.length>8 && <span style={{color:"#6b7280"}}>+{importResult.notFound.length-8} أخرى</span>}
                    </div>
                  </div>
                )}

                <div style={{background:darkMode?"#161000":"#fefce8",border:`1px solid ${darkMode?"#92400e":"#fcd34d"}`,borderRadius:8,padding:"10px 14px",fontSize:12,color:darkMode?"#fde68a":"#78350f",marginBottom:20}}>
                  💡 <strong>تلميح:</strong> تأكد أن ملف Excel من مقيم يحتوي على أعمدة "رقم الاقامة" و"تاريخ انتهاء الاقامة"
                </div>

                <button onClick={()=>setImportModal(false)}
                  style={{width:"100%",background:"#2563eb",color:"#fff",border:"none",borderRadius:10,padding:"11px",fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>
                  إغلاق
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ══════ نافذة الإشعارات ══════ */}
        {notifModal && (
          <div onClick={()=>setNotifModal(false)}
            style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",zIndex:1200,display:"flex",alignItems:"center",justifyContent:"center",padding:16,backdropFilter:"blur(3px)"}}>
            <div onClick={e=>e.stopPropagation()}
              style={{background:darkMode?"#161920":"#fff",borderRadius:16,width:"100%",maxWidth:420,boxShadow:"0 24px 64px rgba(0,0,0,0.4)",overflow:"hidden"}}>

              <div style={{background:"linear-gradient(135deg,#92400e,#d97706)",padding:"18px 24px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div style={{color:"#fff"}}>
                  <div style={{fontWeight:800,fontSize:17}}>🔔 إشعارات انتهاء الإقامة</div>
                  <div style={{fontSize:12,opacity:.85,marginTop:3}}>تذكير تلقائي قبل الانتهاء</div>
                </div>
                <button onClick={()=>setNotifModal(false)}
                  style={{background:"rgba(255,255,255,0.15)",border:"1.5px solid rgba(255,255,255,0.4)",color:"#fff",borderRadius:8,width:32,height:32,fontSize:16,cursor:"pointer",fontWeight:700}}>✕</button>
              </div>

              <div style={{padding:"24px"}}>
                {/* إحصائية سريعة */}
                <div style={{background:darkMode?"#1e222b":"#fffbeb",border:`1px solid ${darkMode?"#92400e":"#fcd34d"}`,borderRadius:10,padding:"14px 16px",marginBottom:20}}>
                  <div style={{fontSize:13,color:darkMode?"#fde68a":"#78350f",fontWeight:700,marginBottom:8}}>📊 الإقامات القريبة من الانتهاء</div>
                  {[7,14,30,60].map(d=>{
                    const count = records.filter(r=>{const days=getDaysLeft(r.expiryDate);return days>=0&&days<=d;}).length;
                    return (
                      <div key={d} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"4px 0",borderBottom:`1px solid ${darkMode?"#2a2f3d":"#fde68a"}`}}>
                        <span style={{fontSize:12,color:darkMode?"#a0a8bb":"#78350f"}}>خلال {d} يوم</span>
                        <span style={{fontWeight:700,color:count>0?"#d97706":darkMode?"#6ee7b7":"#16a34a",fontSize:14}}>{count} إقامة</span>
                      </div>
                    );
                  })}
                </div>

                {/* اختيار عدد الأيام */}
                <div style={{marginBottom:20}}>
                  <label style={{display:"block",fontSize:13,fontWeight:700,color:darkMode?"#a0a8bb":"#374151",marginBottom:8}}>
                    ⏰ أشعرني قبل انتهاء الإقامة بـ:
                  </label>
                  <div style={{display:"flex",gap:8}}>
                    {[7,14,30,60].map(d=>(
                      <button key={d} onClick={()=>setNotifDays(d)}
                        style={{flex:1,padding:"8px 4px",borderRadius:8,border:`2px solid ${notifDays===d?"#d97706":"#d1d5db"}`,background:notifDays===d?"#d97706":darkMode?"#161920":"#fff",color:notifDays===d?"#fff":darkMode?"#f0f2f7":"#374151",fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>
                        {d} يوم
                      </button>
                    ))}
                  </div>
                </div>

                <div style={{background:darkMode?"#1e222b":"#f9fafb",borderRadius:8,padding:"10px 14px",fontSize:12,color:darkMode?"#a0a8bb":"#6b7280",marginBottom:20,lineHeight:1.8}}>
                  <strong style={{color:darkMode?"#f0f2f7":"#374151"}}>كيف يعمل؟</strong><br/>
                  • افتح التطبيق على جهازك أو جوالك<br/>
                  • اضغط "تفعيل الإشعارات" واقبل الإذن<br/>
                  • ستظهر إشعارات فورية للإقامات المنتهية قريباً<br/>
                  • للإشعارات المجدولة يومياً: أضف التطبيق للشاشة الرئيسية (PWA)
                </div>

                <div style={{display:"flex",gap:10}}>
                  <button onClick={requestNotifications}
                    style={{flex:1,background:"#d97706",color:"#fff",border:"none",borderRadius:10,padding:"12px",fontWeight:800,fontSize:14,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
                    🔔 تفعيل الإشعارات الآن
                  </button>
                  <button onClick={()=>setNotifModal(false)}
                    style={{background:darkMode?"#1e222b":"#f3f4f6",color:darkMode?"#f0f2f7":"#374151",border:darkMode?"1px solid #2a2f3d":"none",borderRadius:10,padding:"12px 16px",fontWeight:600,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>
                    لاحقاً
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ══════ نافذة التجديد ══════ */}
        {renewModal && (
          <div onClick={()=>setRenewModal(null)}
            style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",zIndex:1100,display:"flex",alignItems:"center",justifyContent:"center",padding:16,backdropFilter:"blur(3px)"}}>
            <div onClick={e=>e.stopPropagation()}
              style={{background:darkMode?"#161920":"#fff",borderRadius:16,width:"100%",maxWidth:440,boxShadow:"0 24px 64px rgba(0,0,0,0.4)",overflow:"hidden"}}>

              {/* هيدر */}
              <div style={{background:"linear-gradient(135deg,#14532d,#16a34a)",padding:"18px 24px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div>
                  <div style={{color:"#fff",fontWeight:800,fontSize:17}}>🔄 تجديد الإقامة</div>
                  <div style={{color:"rgba(255,255,255,0.85)",fontSize:12,marginTop:3}}>{renewModal.name}</div>
                </div>
                <button onClick={()=>setRenewModal(null)}
                  style={{background:"rgba(255,255,255,0.15)",border:"1.5px solid rgba(255,255,255,0.4)",color:"#fff",borderRadius:8,width:32,height:32,fontSize:16,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700}}>✕</button>
              </div>

              {/* المحتوى */}
              <div style={{padding:"22px 24px"}}>

                {/* معلومات الإقامة الحالية */}
                <div style={{background:darkMode?"#1e222b":"#f0fdf4",borderRadius:10,padding:"12px 16px",marginBottom:18,border:`1px solid ${darkMode?"#166534":"#86efac"}`}}>
                  <div style={{fontSize:12,color:darkMode?"#a0a8bb":"#6b7280",marginBottom:6}}>الإقامة الحالية</div>
                  <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
                    <div>
                      <div style={{fontSize:11,color:darkMode?"#a0a8bb":"#6b7280"}}>تاريخ الانتهاء الحالي</div>
                      <div style={{fontWeight:700,fontSize:14,color:darkMode?"#f0f2f7":"#1e3a5f"}}>
                        {renewModal.expiryDate ? new Date(renewModal.expiryDate).toLocaleDateString("ar-SA",{year:"numeric",month:"long",day:"numeric"}) : "—"}
                      </div>
                    </div>
                    <div>
                      <div style={{fontSize:11,color:darkMode?"#a0a8bb":"#6b7280"}}>الأيام المتبقية</div>
                      <div style={{fontWeight:700,fontSize:14,color:getDaysLeft(renewModal.expiryDate)<0?"#dc2626":getDaysLeft(renewModal.expiryDate)<=30?"#d97706":"#16a34a"}}>
                        {renewModal.expiryDate ? (getDaysLeft(renewModal.expiryDate)<0 ? `منتهية منذ ${Math.abs(getDaysLeft(renewModal.expiryDate))} يوم` : `${getDaysLeft(renewModal.expiryDate)} يوم`) : "—"}
                      </div>
                    </div>
                  </div>
                </div>

                {/* تاريخ الانتهاء الجديد */}
                <div style={{marginBottom:16}}>
                  <label style={{display:"block",fontSize:13,fontWeight:700,color:darkMode?"#a0a8bb":"#374151",marginBottom:6}}>
                    📅 تاريخ انتهاء الإقامة الجديد <span style={{color:"#dc2626"}}>*</span>
                  </label>
                  <input type="date" value={renewDate} onChange={e=>setRenewDate(e.target.value)}
                    min={new Date().toISOString().slice(0,10)}
                    style={{...inp,fontSize:"15px",padding:"10px 14px",border:`2px solid ${renewDate?"#16a34a":"#d1d5db"}`}}/>
                  {renewDate && (
                    <div style={{fontSize:12,color:"#16a34a",marginTop:4,fontWeight:600}}>
                      ✅ صالحة لـ {getDaysLeft(renewDate)} يوم من اليوم
                    </div>
                  )}
                </div>

                {/* ملاحظة */}
                <div style={{marginBottom:20}}>
                  <label style={{display:"block",fontSize:13,fontWeight:700,color:darkMode?"#a0a8bb":"#374151",marginBottom:6}}>
                    📝 ملاحظة (اختياري)
                  </label>
                  <input type="text" value={renewNote} onChange={e=>setRenewNote(e.target.value)}
                    placeholder="مثال: تم التجديد عبر المنصة الوطنية..."
                    style={{...inp}}/>
                </div>

                {/* تنبيه التغييرات */}
                <div style={{background:darkMode?"#161000":"#fefce8",border:`1px solid ${darkMode?"#92400e":"#fcd34d"}`,borderRadius:8,padding:"10px 14px",marginBottom:20,fontSize:12,color:darkMode?"#fde68a":"#78350f"}}>
                  <strong>سيتم تلقائياً:</strong> تحديث تاريخ الانتهاء · تغيير حالة التجديد إلى "مكتمل" · تسجيل تاريخ اليوم كآخر تجديد
                </div>

                {/* الأزرار */}
                <div style={{display:"flex",gap:10}}>
                  <button onClick={handleRenew}
                    disabled={!renewDate}
                    style={{flex:1,background:renewDate?"#16a34a":"#d1d5db",color:"#fff",border:"none",borderRadius:10,padding:"12px",fontWeight:800,fontSize:15,cursor:renewDate?"pointer":"not-allowed",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:8,transition:"background 0.2s"}}>
                    🔄 تأكيد التجديد
                  </button>
                  <button onClick={()=>setRenewModal(null)}
                    style={{background:darkMode?"#1e222b":"#f3f4f6",color:darkMode?"#f0f2f7":"#374151",border:darkMode?"1px solid #2a2f3d":"none",borderRadius:10,padding:"12px 20px",fontWeight:600,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>
                    إلغاء
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}


        {/* ══════ صفحة التقارير ══════ */}
        {activeTab==="reports"&&(()=>{
          const dm = darkMode;
          const card = {background:dm?"#161920":"#fff",borderRadius:14,padding:"20px 22px",boxShadow:dm?"0 2px 10px rgba(0,0,0,0.5)":"0 2px 10px rgba(0,0,0,0.07)",marginBottom:18};
          const titleStyle = {fontSize:15,fontWeight:800,color:dm?"#f0f2f7":"#1e3a5f",marginBottom:16,display:"flex",alignItems:"center",gap:8};
          const fmt = n => n.toLocaleString("ar-SA") + " ريال";

          // قائمة التقارير المتاحة
          const REPORT_LIST = [
            {id:"status",      label:"⏳ حالة الإقامة",          desc:"توزيع السارية والمنتهية"},
            {id:"company",     label:"🏢 توزيع الشركة",          desc:"انجال المشاعر ودلتا الماسية"},
            {id:"nationality", label:"🏆 توزيع الجنسية",         desc:"جميع الجنسيات بنسبها"},
            {id:"jobs",        label:"💼 تقرير المهن التفصيلي",   desc:"جميع المهن مع أسماء الموظفين"},
            {id:"location",    label:"🌍 داخل وخارج المملكة",     desc:"تقرير الوجود والغياب"},
            {id:"calendar",    label:"📅 تقويم الانتهاء",        desc:"الإقامات المنتهية شهرياً"},
            {id:"costs",       label:"💰 تكاليف التجديد",        desc:"التكاليف الشهرية المتوقعة"},
            {id:"renewed",     label:"✅ الإقامات المجددة",       desc:"سجل التجديدات وتكلفتها"},
          ];

          const toggleReport = (id) => {
            const next = new Set(selectedReports);
            if(next.has(id)) { if(next.size>1) next.delete(id); }
            else next.add(id);
            setSelectedReports(next);
          };

          // ── تصدير Excel للتقارير ──
          const exportReportsExcel = () => {
            const wb = XLSX.utils.book_new();

            if(selectedReports.has("status")) {
              const ws = XLSX.utils.json_to_sheet([
                {الحالة:"سارية",        العدد:records.filter(r=>getStatus(r)==="سارية").length},
                {الحالة:"تنتهي قريباً",العدد:records.filter(r=>getStatus(r)==="تنتهي قريباً").length},
                {الحالة:"منتهية",       العدد:records.filter(r=>getStatus(r)==="منتهية").length},
                {الحالة:"قيد التجديد", العدد:records.filter(r=>getStatus(r)==="قيد التجديد").length},
                {الحالة:"مرافق",        العدد:records.filter(r=>r.type==="مرافق").length},
              ]);
              ws["!cols"]=[{wch:18},{wch:10}];
              XLSX.utils.book_append_sheet(wb, ws, "حالة الإقامة");
            }

            if(selectedReports.has("nationality")) {
              const natMap={};
              records.forEach(r=>{if(r.nationality)natMap[r.nationality]=(natMap[r.nationality]||0)+1;});
              const ws = XLSX.utils.json_to_sheet(
                Object.entries(natMap).sort((a,b)=>b[1]-a[1]).map(([n,c])=>({
                  الجنسية:n, العدد:c, النسبة:`${(c/records.length*100).toFixed(1)}%`
                }))
              );
              ws["!cols"]=[{wch:16},{wch:10},{wch:10}];
              XLSX.utils.book_append_sheet(wb, ws, "توزيع الجنسية");
            }

            if(selectedReports.has("company")) {
              const ws = XLSX.utils.json_to_sheet([
                {الشركة:"انجال المشاعر",العدد:records.filter(r=>r.company==="انجال المشاعر").length},
                {الشركة:"دلتا الماسية", العدد:records.filter(r=>r.company==="دلتا الماسية").length},
                {الشركة:"موظفون",        العدد:records.filter(r=>r.type!=="مرافق").length},
                {الشركة:"مرافقون",       العدد:records.filter(r=>r.type==="مرافق").length},
              ]);
              ws["!cols"]=[{wch:18},{wch:10}];
              XLSX.utils.book_append_sheet(wb, ws, "توزيع الشركة");
            }

            if(selectedReports.has("jobs")) {
              // ورقة 1: ملخص المهن
              const jobMap={};
              records.filter(r=>r.type!=="مرافق").forEach(r=>{
                const j=(r.jobTitle||r.notes||"غير محدد").trim();
                if(j)jobMap[j]=(jobMap[j]||0)+1;
              });
              const wsSummary = XLSX.utils.json_to_sheet(
                Object.entries(jobMap).sort((a,b)=>b[1]-a[1]).map(([j,c])=>({
                  المهنة:j, العدد:c, النسبة:`${(c/records.filter(r=>r.type!=="مرافق").length*100).toFixed(1)}%`
                }))
              );
              wsSummary["!cols"]=[{wch:32},{wch:10},{wch:10}];
              XLSX.utils.book_append_sheet(wb, wsSummary, "ملخص المهن");
              // ورقة 2: تفصيل بأسماء الموظفين
              const wsDetail = XLSX.utils.json_to_sheet(
                records.filter(r=>r.type!=="مرافق")
                  .sort((a,b)=>(a.jobTitle||a.notes||"").localeCompare(b.jobTitle||b.notes||"","ar"))
                  .map(r=>({
                    المهنة: r.jobTitle||r.notes||"غير محدد",
                    الاسم: r.name,
                    "رقم الإقامة": r.iqamaNumber,
                    الجنسية: r.nationality||"-",
                    الشركة: r.company,
                    "حالة الإقامة": getStatus(r),
                  }))
              );
              wsDetail["!cols"]=[{wch:30},{wch:28},{wch:14},{wch:12},{wch:18},{wch:14}];
              XLSX.utils.book_append_sheet(wb, wsDetail, "تفصيل الموظفين بالمهنة");
            }

            if(selectedReports.has("location")) {
              const inside  = records.filter(r=>r.outsideKingdom!=="نعم");
              const outside = records.filter(r=>r.outsideKingdom==="نعم");
              // ملخص
              const wsSummary = XLSX.utils.json_to_sheet([
                {الوجود:"داخل المملكة", العدد:inside.length,  النسبة:`${(inside.length/records.length*100).toFixed(1)}%`},
                {الوجود:"خارج المملكة", العدد:outside.length, النسبة:`${(outside.length/records.length*100).toFixed(1)}%`},
              ]);
              wsSummary["!cols"]=[{wch:18},{wch:10},{wch:10}];
              XLSX.utils.book_append_sheet(wb, wsSummary, "ملخص الوجود");
              // خارج المملكة تفصيل
              const wsOut = XLSX.utils.json_to_sheet(
                outside.map(r=>({
                  الاسم: r.name, "رقم الإقامة": r.iqamaNumber,
                  الجنسية: r.nationality||"-", المهنة: r.jobTitle||r.notes||"-",
                  الشركة: r.company, "حالة الإقامة": getStatus(r),
                  "تاريخ الانتهاء": r.expiryDate||"-",
                }))
              );
              wsOut["!cols"]=[{wch:28},{wch:14},{wch:12},{wch:28},{wch:18},{wch:14},{wch:14}];
              XLSX.utils.book_append_sheet(wb, wsOut, "خارج المملكة");
            }

            if(selectedReports.has("calendar")||selectedReports.has("costs")) {
              const now=new Date();
              const rows=[];
              for(let i=0;i<12;i++){
                const m=new Date(now.getFullYear(),now.getMonth()+i,1);
                const nextM=new Date(now.getFullYear(),now.getMonth()+i+1,1);
                const label=m.toLocaleDateString("ar-SA",{month:"long",year:"numeric"});
                const expCount=records.filter(r=>{
                  if(!r.expiryDate)return false;
                  const d=new Date(r.expiryDate);
                  return d>=m&&d<nextM;
                }).length;
                const cost=records.filter(r=>{
                  if(!r.expiryDate||r.type==="مرافق")return false;
                  const d=new Date(r.expiryDate);
                  return d>=m&&d<nextM;
                }).length*(2425+163)+records.filter(r=>{
                  if(!r.expiryDate||r.type!=="مرافق")return false;
                  const head=records.find(e=>e.iqamaNumber===r.familyHeadId);
                  if(!head?.expiryDate)return false;
                  const d=new Date(head.expiryDate);
                  return d>=m&&d<nextM;
                }).length*1200;
                rows.push({الشهر:label,"عدد المنتهية":expCount,"تكلفة التجديد (ريال)":cost});
              }
              const ws=XLSX.utils.json_to_sheet(rows);
              ws["!cols"]=[{wch:20},{wch:16},{wch:20}];
              XLSX.utils.book_append_sheet(wb, ws, "التقويم والتكاليف");
            }

            if(selectedReports.has("renewed")) {
              const renewed=records.filter(r=>r.lastRenewalDate);
              const ws=XLSX.utils.json_to_sheet(renewed.map(r=>({
                الاسم:r.name,
                "رقم الإقامة":r.iqamaNumber,
                النوع:r.type==="مرافق"?"مرافق":"موظف",
                الشركة:r.company,
                "تاريخ آخر تجديد":r.lastRenewalDate,
                "تاريخ الانتهاء الجديد":r.expiryDate,
                "التكلفة (ريال)":r.type!=="مرافق"?2588:1200,
                الملاحظة:r.lastRenewalNote||"",
              })));
              ws["!cols"]=[{wch:28},{wch:14},{wch:10},{wch:18},{wch:16},{wch:18},{wch:14},{wch:24}];
              XLSX.utils.book_append_sheet(wb, ws, "الإقامات المجددة");
            }

            XLSX.writeFile(wb, `تقارير_الإقامات_${new Date().toISOString().slice(0,10)}.xlsx`);
          };

          // ── تصدير PDF للتقارير ──
          const exportReportsPDF = () => {
            const today=new Date().toLocaleDateString("ar-SA",{year:"numeric",month:"long",day:"numeric"});
            let sections="";

            if(selectedReports.has("status")) {
              const data=[
                {l:"سارية",c:records.filter(r=>getStatus(r)==="سارية").length,col:"#16a34a"},
                {l:"تنتهي قريباً",c:records.filter(r=>getStatus(r)==="تنتهي قريباً").length,col:"#d97706"},
                {l:"منتهية",c:records.filter(r=>getStatus(r)==="منتهية").length,col:"#dc2626"},
                {l:"مرافق",c:records.filter(r=>r.type==="مرافق").length,col:"#7c3aed"},
              ];
              sections+=`<h2 style="color:#6B1A1A;border-bottom:2px solid #6B1A1A;padding-bottom:6px">⏳ حالة الإقامة</h2>
              <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:20px">
              ${data.map(d=>`<div style="background:#f9fafb;border:1px solid #e5e7eb;border-top:3px solid ${d.col};border-radius:8px;padding:12px;text-align:center">
                <div style="font-size:26px;font-weight:800;color:${d.col}">${d.c}</div>
                <div style="font-size:11px;color:#6b7280">${d.l}</div>
              </div>`).join("")}</div>`;
            }

            if(selectedReports.has("nationality")) {
              const natMap={};
              records.forEach(r=>{if(r.nationality)natMap[r.nationality]=(natMap[r.nationality]||0)+1;});
              const sorted=Object.entries(natMap).sort((a,b)=>b[1]-a[1]);
              const natColors=["#6B1A1A","#F5A800","#2563eb","#16a34a","#d97706","#7c3aed"];
              sections+=`<h2 style="color:#6B1A1A;border-bottom:2px solid #6B1A1A;padding-bottom:6px">🏆 توزيع الجنسية</h2>
              <table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-size:12px">
              <thead><tr style="background:#6B1A1A;color:#fff"><th style="padding:8px">الجنسية</th><th style="padding:8px">العدد</th><th style="padding:8px">النسبة</th><th style="padding:8px;width:40%">التوزيع</th></tr></thead>
              <tbody>${sorted.map(([n,c],i)=>`<tr style="background:${i%2===0?"#f9fafb":"#fff"}">
                <td style="padding:7px;text-align:center;font-weight:600">${n}</td>
                <td style="padding:7px;text-align:center;font-weight:700;color:${natColors[i%natColors.length]}">${c}</td>
                <td style="padding:7px;text-align:center">${(c/records.length*100).toFixed(1)}%</td>
                <td style="padding:7px"><div style="height:14px;background:${natColors[i%natColors.length]};border-radius:4px;width:${(c/records.length*100).toFixed(0)}%"></div></td>
              </tr>`).join("")}</tbody></table>`;
            }

            if(selectedReports.has("jobs")) {
              const jobMap={};
              const empOnly=records.filter(r=>r.type!=="مرافق");
              empOnly.forEach(r=>{const j=(r.jobTitle||r.notes||"غير محدد").trim();if(j)jobMap[j]=(jobMap[j]||0)+1;});
              const sorted=Object.entries(jobMap).sort((a,b)=>b[1]-a[1]);
              // تجميع الموظفين بالمهنة
              const byJob={};
              empOnly.forEach(r=>{const j=(r.jobTitle||r.notes||"غير محدد").trim();if(!byJob[j])byJob[j]=[];byJob[j].push(r);});
              sections+=`<h2 style="color:#6B1A1A;border-bottom:2px solid #6B1A1A;padding-bottom:6px">💼 تقرير المهن التفصيلي (${empOnly.length} موظف - ${sorted.length} مهنة)</h2>
              <table style="width:100%;border-collapse:collapse;margin-bottom:8px;font-size:11px">
              <thead><tr style="background:#6B1A1A;color:#fff"><th style="padding:7px">#</th><th style="padding:7px">المهنة</th><th style="padding:7px">العدد</th><th style="padding:7px">النسبة</th><th style="padding:7px">الأسماء</th></tr></thead>
              <tbody>${sorted.map(([j,c],i)=>`<tr style="background:${i%2===0?"#f9fafb":"#fff"}">
                <td style="padding:6px;text-align:center;font-weight:700">${i+1}</td>
                <td style="padding:6px;font-weight:700;color:#6B1A1A">${j}</td>
                <td style="padding:6px;text-align:center;font-weight:800;color:#6B1A1A">${c}</td>
                <td style="padding:6px;text-align:center">${(c/empOnly.length*100).toFixed(1)}%</td>
                <td style="padding:6px;font-size:10px;color:#374151">${(byJob[j]||[]).map(r=>r.name).join(" · ")}</td>
              </tr>`).join("")}</tbody></table>
              <div style="margin-bottom:20px;font-size:10px;color:#6b7280;text-align:left">* المهن مرتبة تنازلياً حسب العدد</div>`;
            }

            if(selectedReports.has("location")) {
              const inside  = records.filter(r=>r.outsideKingdom!=="نعم");
              const outside = records.filter(r=>r.outsideKingdom==="نعم");
              const insideEmp  = inside.filter(r=>r.type!=="مرافق");
              const outsideEmp = outside.filter(r=>r.type!=="مرافق");
              sections+=`<h2 style="color:#6B1A1A;border-bottom:2px solid #6B1A1A;padding-bottom:6px">🌍 تقرير الوجود داخل وخارج المملكة</h2>
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px">
                <div style="background:#f0fdf4;border:1px solid #86efac;border-radius:8px;padding:14px;text-align:center">
                  <div style="font-size:28px;font-weight:800;color:#16a34a">${inside.length}</div>
                  <div style="font-size:13px;font-weight:700;color:#15803d">🏠 داخل المملكة</div>
                  <div style="font-size:11px;color:#6b7280;margin-top:4px">(${(inside.length/records.length*100).toFixed(1)}%) — ${insideEmp.length} موظف + ${inside.length-insideEmp.length} مرافق</div>
                </div>
                <div style="background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;padding:14px;text-align:center">
                  <div style="font-size:28px;font-weight:800;color:#dc2626">${outside.length}</div>
                  <div style="font-size:13px;font-weight:700;color:#dc2626">✈️ خارج المملكة</div>
                  <div style="font-size:11px;color:#6b7280;margin-top:4px">(${(outside.length/records.length*100).toFixed(1)}%) — ${outsideEmp.length} موظف + ${outside.length-outsideEmp.length} مرافق</div>
                </div>
              </div>
              ${outside.length>0?`<h3 style="color:#dc2626;margin-bottom:8px;font-size:13px">✈️ المتواجدون خارج المملكة (${outside.length})</h3>
              <table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-size:11px">
              <thead><tr style="background:#dc2626;color:#fff"><th style="padding:7px">الاسم</th><th style="padding:7px">الجنسية</th><th style="padding:7px">المهنة</th><th style="padding:7px">الشركة</th><th style="padding:7px">حالة الإقامة</th></tr></thead>
              <tbody>${outside.map((r,i)=>`<tr style="background:${i%2===0?"#fef2f2":"#fff"}">
                <td style="padding:6px;font-weight:600">${r.name}</td>
                <td style="padding:6px;text-align:center">${r.nationality||"-"}</td>
                <td style="padding:6px">${r.jobTitle||r.notes||"-"}</td>
                <td style="padding:6px;text-align:center">${r.company}</td>
                <td style="padding:6px;text-align:center;font-weight:700;color:${getStatus(r)==="منتهية"?"#dc2626":getStatus(r)==="تنتهي قريباً"?"#d97706":"#16a34a"}">${getStatus(r)}</td>
              </tr>`).join("")}</tbody></table>`:""}`;
            }

            if(selectedReports.has("calendar")||selectedReports.has("costs")) {
              const now=new Date();
              const rows=[];
              for(let i=0;i<12;i++){
                const m=new Date(now.getFullYear(),now.getMonth()+i,1);
                const nextM=new Date(now.getFullYear(),now.getMonth()+i+1,1);
                const label=m.toLocaleDateString("ar-SA",{month:"long",year:"numeric"});
                const expCount=records.filter(r=>{if(!r.expiryDate)return false;const d=new Date(r.expiryDate);return d>=m&&d<nextM;}).length;
                const cost=records.filter(r=>{if(!r.expiryDate||r.type==="مرافق")return false;const d=new Date(r.expiryDate);return d>=m&&d<nextM;}).length*(2425+163);
                rows.push({label,expCount,cost});
              }
              sections+=`<h2 style="color:#6B1A1A;border-bottom:2px solid #6B1A1A;padding-bottom:6px">📅 تقويم الانتهاء والتكاليف الشهرية</h2>
              <table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-size:12px">
              <thead><tr style="background:#6B1A1A;color:#fff"><th style="padding:8px">الشهر</th><th style="padding:8px">عدد المنتهية</th><th style="padding:8px">تكلفة التجديد</th></tr></thead>
              <tbody>${rows.map((r,i)=>`<tr style="background:${i%2===0?"#f9fafb":"#fff"}">
                <td style="padding:7px;font-weight:600">${r.label}</td>
                <td style="padding:7px;text-align:center;font-weight:700;color:${r.expCount>3?"#dc2626":r.expCount>0?"#d97706":"#16a34a"}">${r.expCount}</td>
                <td style="padding:7px;text-align:center;font-weight:700;color:#16a34a">${r.cost.toLocaleString("ar-SA")} ريال</td>
              </tr>`).join("")}</tbody></table>`;
            }

            if(selectedReports.has("renewed")) {
              const renewed=records.filter(r=>r.lastRenewalDate);
              const totalCost=renewed.filter(r=>r.type!=="مرافق").length*2588+renewed.filter(r=>r.type==="مرافق").length*1200;
              sections+=`<h2 style="color:#6B1A1A;border-bottom:2px solid #6B1A1A;padding-bottom:6px">✅ الإقامات المجددة (${renewed.length} سجل)</h2>
              ${renewed.length===0?'<p style="color:#6b7280;text-align:center;padding:20px">لا توجد تجديدات مسجلة</p>':`
              <table style="width:100%;border-collapse:collapse;margin-bottom:10px;font-size:11px">
              <thead><tr style="background:#6B1A1A;color:#fff"><th style="padding:7px">الاسم</th><th style="padding:7px">الشركة</th><th style="padding:7px">تاريخ التجديد</th><th style="padding:7px">ينتهي</th><th style="padding:7px">التكلفة</th></tr></thead>
              <tbody>${renewed.sort((a,b)=>new Date(b.lastRenewalDate)-new Date(a.lastRenewalDate)).map((r,i)=>`<tr style="background:${i%2===0?"#f9fafb":"#fff"}">
                <td style="padding:6px;font-weight:600">${r.name}</td>
                <td style="padding:6px;text-align:center">${r.company}</td>
                <td style="padding:6px;text-align:center">${new Date(r.lastRenewalDate).toLocaleDateString("ar-SA")}</td>
                <td style="padding:6px;text-align:center">${r.expiryDate?new Date(r.expiryDate).toLocaleDateString("ar-SA"):"—"}</td>
                <td style="padding:6px;text-align:center;font-weight:700;color:#16a34a">${(r.type!=="مرافق"?2588:1200).toLocaleString("ar-SA")} ر</td>
              </tr>`).join("")}</tbody></table>
              <div style="background:#f0fdf4;border:1px solid #86efac;border-radius:8px;padding:10px 16px;display:flex;justify-content:space-between">
                <span style="font-weight:700;color:#15803d">إجمالي تكلفة التجديدات</span>
                <span style="font-weight:900;font-size:16px;color:#15803d">${totalCost.toLocaleString("ar-SA")} ريال</span>
              </div>`}`;
            }

            const html=`<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="UTF-8"><title>تقارير الإقامات</title>
            <style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Segoe UI',Tahoma,sans-serif;direction:rtl;padding:24px;color:#1f2937;font-size:13px}
            .header{background:linear-gradient(135deg,#6B1A1A,#F5A800);color:#fff;padding:20px 24px;border-radius:12px;margin-bottom:24px;display:flex;justify-content:space-between;align-items:center}
            @media print{body{padding:12px}}</style></head><body>
            <div class="header">
              <div><h1 style="font-size:18px;font-weight:800">📊 تقارير متابعة الإقامات</h1><p style="opacity:.85;margin-top:4px;font-size:12px">تاريخ التقرير: ${today}</p></div>
              <div style="text-align:left;font-size:12px">إجمالي السجلات: <strong>${records.length}</strong></div>
            </div>
            ${sections}
            <script>window.onload=()=>window.print();</script></body></html>`;
            const win=window.open(URL.createObjectURL(new Blob([html],{type:"text/html;charset=utf-8"})),"_blank");
            if(!win)alert("يرجى السماح بفتح النوافذ المنبثقة");
          };

          // ── بيانات التقارير ──
          const employees = records.filter(r=>r.type!=="مرافق");

          // 1. حالة الإقامة
          const statusData = [
            {label:"سارية",     count:records.filter(r=>getStatus(r)==="سارية").length,        color:"#16a34a",bg:"#dcfce7"},
            {label:"تنتهي قريباً",count:records.filter(r=>getStatus(r)==="تنتهي قريباً").length, color:"#d97706",bg:"#fef3c7"},
            {label:"منتهية",    count:records.filter(r=>getStatus(r)==="منتهية").length,        color:"#dc2626",bg:"#fee2e2"},
            {label:"قيد التجديد",count:records.filter(r=>getStatus(r)==="قيد التجديد").length, color:"#2563eb",bg:"#dbeafe"},
            {label:"مرافق",     count:records.filter(r=>r.type==="مرافق").length,              color:"#7c3aed",bg:"#f3e8ff"},
          ].filter(d=>d.count>0);
          const totalStatus = statusData.reduce((s,d)=>s+d.count,0);

          // 2. توزيع الجنسية
          const natMap = {};
          records.forEach(r=>{if(r.nationality){natMap[r.nationality]=(natMap[r.nationality]||0)+1;}});
          const natData = Object.entries(natMap).sort((a,b)=>b[1]-a[1]);
          const natColors = ["#6B1A1A","#F5A800","#2563eb","#16a34a","#d97706","#7c3aed","#dc2626","#0891b2","#65a30d","#e11d48"];

          // 3. توزيع الشركة
          const compData = [
            {label:"انجال المشاعر",count:records.filter(r=>r.company==="انجال المشاعر").length,color:"#6B1A1A"},
            {label:"دلتا الماسية", count:records.filter(r=>r.company==="دلتا الماسية").length, color:"#F5A800"},
          ];

          // 4. المهن الأكثر
          const jobMap = {};
          employees.forEach(r=>{const j=(r.jobTitle||r.notes||"غير محدد").trim();if(j)jobMap[j]=(jobMap[j]||0)+1;});
          const jobData = Object.entries(jobMap).sort((a,b)=>b[1]-a[1]).slice(0,8);

          // 5. تقويم الانتهاء شهرياً (12 شهر قادم)
          const monthlyData = [];
          const now = new Date();
          for(let i=0;i<12;i++){
            const m = new Date(now.getFullYear(), now.getMonth()+i, 1);
            const nextM = new Date(now.getFullYear(), now.getMonth()+i+1, 1);
            const label = m.toLocaleDateString("ar-SA",{month:"short",year:"2-digit"});
            const count = records.filter(r=>{
              if(!r.expiryDate) return false;
              const d=new Date(r.expiryDate);
              return d>=m && d<nextM;
            }).length;
            monthlyData.push({label,count,month:m});
          }
          const maxMonthly = Math.max(...monthlyData.map(d=>d.count),1);

          // 6. تكاليف التجديد الشهرية (الإجمالي لو تم تجديد كل من ينتهي)
          const monthlyCostData = monthlyData.map(({label,month})=>{
            const nextM = new Date(month.getFullYear(), month.getMonth()+1, 1);
            const expiring = records.filter(r=>{
              if(!r.expiryDate||r.type==="مرافق") return false;
              const d=new Date(r.expiryDate);
              return d>=month && d<nextM;
            });
            const cost = expiring.length * (2425+163); // ربع واحد بسعر عادي
            const depCost = records.filter(r=>{
              if(!r.expiryDate||r.type!=="مرافق") return false;
              const head = records.find(e=>e.iqamaNumber===r.familyHeadId);
              if(!head||!head.expiryDate) return false;
              const d=new Date(head.expiryDate);
              return d>=month && d<nextM;
            }).length * 1200;
            return {label, cost:cost+depCost, month};
          });
          const maxCost = Math.max(...monthlyCostData.map(d=>d.cost),1);

          // 7. الإقامات المجددة وتكلفتها
          const renewedRecords = records.filter(r=>r.lastRenewalDate);
          const totalRenewedCost = renewedRecords.filter(r=>r.type!=="مرافق").length*(2425+163)
            + renewedRecords.filter(r=>r.type==="مرافق").length*1200;

          // SVG Donut
          const DonutChart = ({data,size=180})=>{
            const total = data.reduce((s,d)=>s+d.count,0);
            if(total===0) return null;
            let cumAngle = -Math.PI/2;
            const cx=size/2, cy=size/2, r=size*0.36, inner=size*0.22;
            const slices = data.map(d=>{
              const angle = (d.count/total)*2*Math.PI;
              const x1=cx+r*Math.cos(cumAngle), y1=cy+r*Math.sin(cumAngle);
              cumAngle+=angle;
              const x2=cx+r*Math.cos(cumAngle), y2=cy+r*Math.sin(cumAngle);
              const ix1=cx+inner*Math.cos(cumAngle-angle), iy1=cy+inner*Math.sin(cumAngle-angle);
              const ix2=cx+inner*Math.cos(cumAngle), iy2=cy+inner*Math.sin(cumAngle);
              const lg = angle>Math.PI?1:0;
              return {d:`M ${x1} ${y1} A ${r} ${r} 0 ${lg} 1 ${x2} ${y2} L ${ix2} ${iy2} A ${inner} ${inner} 0 ${lg} 0 ${ix1} ${iy1} Z`, color:d.color, label:d.label, count:d.count};
            });
            return (
              <svg viewBox={`0 0 ${size} ${size}`} width={size} height={size} style={{overflow:"visible"}}>
                {slices.map((s,i)=>(
                  <path key={i} d={s.d} fill={s.color} stroke={dm?"#161920":"#fff"} strokeWidth={2}/>
                ))}
                <text x={cx} y={cy-8} textAnchor="middle" fill={dm?"#f0f2f7":"#1e3a5f"} fontSize={size*0.11} fontWeight="800">{total}</text>
                <text x={cx} y={cy+10} textAnchor="middle" fill={dm?"#a0a8bb":"#6b7280"} fontSize={size*0.07}>إجمالي</text>
              </svg>
            );
          };

          return (
            <div>
              {/* ── شريط التخصيص والتصدير ── */}
              <div style={{background:dm?"#161920":"#fff",borderRadius:14,padding:"18px 22px",marginBottom:18,boxShadow:dm?"0 2px 10px rgba(0,0,0,0.5)":"0 2px 10px rgba(0,0,0,0.07)"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12,marginBottom:14}}>
                  <div style={{fontWeight:800,fontSize:15,color:dm?"#f0f2f7":"#1e3a5f",display:"flex",alignItems:"center",gap:8}}>
                    📊 تخصيص التقارير
                    <span style={{background:dm?"#1e222b":"#f3f4f6",color:dm?"#a0a8bb":"#6b7280",padding:"2px 10px",borderRadius:10,fontSize:12,fontWeight:600}}>
                      {selectedReports.size} من {REPORT_LIST.length} محدد
                    </span>
                  </div>
                  {/* أزرار التصدير */}
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>{const next=new Set(REPORT_LIST.map(r=>r.id));setSelectedReports(next);}}
                      style={{padding:"6px 14px",borderRadius:8,border:`1px solid ${dm?"#2a2f3d":"#d1d5db"}`,background:dm?"#1e222b":"#f9fafb",color:dm?"#f0f2f7":"#374151",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>
                      ✅ تحديد الكل
                    </button>
                    <button onClick={exportReportsExcel}
                      style={{padding:"6px 16px",borderRadius:8,border:"1px solid #86efac",background:dm?"#081a12":"#f0fdf4",color:"#16a34a",fontSize:13,cursor:"pointer",fontFamily:"inherit",fontWeight:700,display:"flex",alignItems:"center",gap:5}}>
                      📊 Excel
                    </button>
                    <button onClick={exportReportsPDF}
                      style={{padding:"6px 16px",borderRadius:8,border:"1px solid #fca5a5",background:dm?"#2a1515":"#fef2f2",color:"#dc2626",fontSize:13,cursor:"pointer",fontFamily:"inherit",fontWeight:700,display:"flex",alignItems:"center",gap:5}}>
                      📄 PDF
                    </button>
                  </div>
                </div>
                {/* بطاقات اختيار التقارير */}
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:8}}>
                  {REPORT_LIST.map(r=>{
                    const isSelected = selectedReports.has(r.id);
                    return (
                      <div key={r.id} onClick={()=>toggleReport(r.id)}
                        style={{border:`2px solid ${isSelected?"#6B1A1A":dm?"#2a2f3d":"#e5e7eb"}`,borderRadius:10,padding:"10px 14px",cursor:"pointer",background:isSelected?(dm?"#2d0f0f":"#fdf0f0"):(dm?"#1e222b":"#fafafa"),transition:"all 0.15s",display:"flex",alignItems:"center",gap:10}}>
                        <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${isSelected?"#6B1A1A":dm?"#4b5563":"#d1d5db"}`,background:isSelected?"#6B1A1A":"transparent",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
                          {isSelected&&<span style={{color:"#fff",fontSize:11,fontWeight:900}}>✓</span>}
                        </div>
                        <div>
                          <div style={{fontSize:13,fontWeight:700,color:isSelected?"#6B1A1A":(dm?"#f0f2f7":"#374151")}}>{r.label}</div>
                          <div style={{fontSize:10,color:dm?"#a0a8bb":"#6b7280",marginTop:1}}>{r.desc}</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* ── صف 1: حالة الإقامة + الشركة ── */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(300px,1fr))",gap:16,marginBottom:0}}>

                {/* حالة الإقامة */}
                {selectedReports.has("status")&&<div style={card}>
                  <div style={titleStyle}>⏳ توزيع حالة الإقامة</div>
                  <div style={{display:"flex",alignItems:"center",gap:20,flexWrap:"wrap"}}>
                    <DonutChart data={statusData} size={160}/>
                    <div style={{flex:1,minWidth:120}}>
                      {statusData.map(d=>(
                        <div key={d.label} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                          <div style={{width:12,height:12,borderRadius:3,background:d.color,flexShrink:0}}/>
                          <div style={{flex:1,fontSize:13,color:dm?"#f0f2f7":"#374151"}}>{d.label}</div>
                          <div style={{fontWeight:700,fontSize:13,color:d.color}}>{d.count}</div>
                          <div style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280"}}>({Math.round(d.count/totalStatus*100)}%)</div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>}

                {/* الشركة */}
                {selectedReports.has("company")&&<div style={card}>
                  <div style={titleStyle}>🏢 توزيع حسب الشركة</div>
                  {compData.map(d=>(
                    <div key={d.label} style={{marginBottom:14}}>
                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
                        <span style={{fontSize:13,fontWeight:700,color:dm?"#f0f2f7":"#374151"}}>{d.label}</span>
                        <span style={{fontSize:13,fontWeight:800,color:d.color}}>{d.count} سجل</span>
                      </div>
                      <div style={{background:dm?"#1e222b":"#f3f4f6",borderRadius:8,height:28,overflow:"hidden"}}>
                        <div style={{width:`${(d.count/records.length*100).toFixed(1)}%`,height:"100%",background:d.color,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"flex-end",paddingRight:8,transition:"width 0.6s",minWidth:d.count>0?40:0}}>
                          <span style={{color:"#fff",fontSize:11,fontWeight:700}}>{(d.count/records.length*100).toFixed(0)}%</span>
                        </div>
                      </div>
                    </div>
                  ))}
                  <div style={{borderTop:`1px solid ${dm?"#2a2f3d":"#e5e7eb"}`,paddingTop:12,marginTop:4,display:"flex",gap:16}}>
                    {[{l:"موظفون",v:employees.length,c:"#2563eb"},{l:"مرافقون",v:records.filter(r=>r.type==="مرافق").length,c:"#7c3aed"}].map(x=>(
                      <div key={x.l} style={{flex:1,background:dm?"#1e222b":"#f9fafb",borderRadius:8,padding:"10px",textAlign:"center",border:`1px solid ${dm?"#2a2f3d":"#e5e7eb"}`}}>
                        <div style={{fontSize:22,fontWeight:800,color:x.c}}>{x.v}</div>
                        <div style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280"}}>{x.l}</div>
                      </div>
                    ))}
                  </div>
                </div>}
              </div>

              {/* ── صف 2: الجنسية ── */}
              {selectedReports.has("nationality")&&<div style={card}>
                <div style={titleStyle}>🏆 توزيع حسب الجنسية</div>
                <div style={{display:"grid",gap:10}}>
                  {natData.map(([nat,cnt],i)=>(
                    <div key={nat} style={{display:"flex",alignItems:"center",gap:10}}>
                      <div style={{width:28,height:28,borderRadius:6,background:natColors[i%natColors.length],display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:11,fontWeight:800,flexShrink:0}}>{i+1}</div>
                      <div style={{flex:1,fontSize:13,fontWeight:600,color:dm?"#f0f2f7":"#374151"}}>{nat}</div>
                      <div style={{width:"45%",background:dm?"#1e222b":"#f3f4f6",borderRadius:6,height:22,overflow:"hidden"}}>
                        <div style={{width:`${(cnt/records.length*100).toFixed(1)}%`,height:"100%",background:natColors[i%natColors.length],borderRadius:6,minWidth:cnt>0?24:0}}/>
                      </div>
                      <div style={{fontWeight:800,color:natColors[i%natColors.length],fontSize:14,minWidth:30,textAlign:"right"}}>{cnt}</div>
                      <div style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280",minWidth:36}}>({Math.round(cnt/records.length*100)}%)</div>
                    </div>
                  ))}
                </div>
              </div>}

              {/* ── صف 3: المهن التفصيلي ── */}
              {selectedReports.has("jobs")&&(()=>{
                const allJobMap={};
                const empOnly=records.filter(r=>r.type!=="مرافق");
                empOnly.forEach(r=>{const j=(r.jobTitle||r.notes||"غير محدد").trim();if(j)allJobMap[j]=(allJobMap[j]||0)+1;});
                const allJobsSorted=Object.entries(allJobMap).sort((a,b)=>b[1]-a[1]);
                const byJob={};
                empOnly.forEach(r=>{const j=(r.jobTitle||r.notes||"غير محدد").trim();if(!byJob[j])byJob[j]=[];byJob[j].push(r);});
                return (
                  <div style={card}>
                    <div style={{...titleStyle,justifyContent:"space-between",flexWrap:"wrap"}}>
                      <span>💼 تقرير المهن التفصيلي</span>
                      <div style={{display:"flex",gap:8,fontSize:12}}>
                        <span style={{background:dm?"#1e222b":"#eff6ff",color:dm?"#93c5fd":"#2563eb",padding:"3px 12px",borderRadius:10,fontWeight:600}}>{allJobsSorted.length} مهنة</span>
                        <span style={{background:dm?"#1e222b":"#f0fdf4",color:dm?"#6ee7b7":"#16a34a",padding:"3px 12px",borderRadius:10,fontWeight:600}}>{empOnly.length} موظف</span>
                      </div>
                    </div>
                    <div style={{display:"grid",gap:10}}>
                      {allJobsSorted.map(([job,cnt],i)=>{
                        const emps=byJob[job]||[];
                        const pct=(cnt/empOnly.length*100).toFixed(0);
                        return (
                          <div key={job} style={{background:dm?"#1e222b":"#f9fafb",borderRadius:10,border:`1px solid ${dm?"#2a2f3d":"#e5e7eb"}`,overflow:"hidden"}}>
                            <div style={{padding:"10px 14px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                              <div style={{width:28,height:28,borderRadius:6,background:natColors[i%natColors.length],display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:11,fontWeight:800,flexShrink:0}}>{i+1}</div>
                              <div style={{flex:1}}>
                                <div style={{fontWeight:700,fontSize:13,color:dm?"#f0f2f7":"#1e3a5f"}}>{job}</div>
                                <div style={{display:"flex",gap:8,marginTop:3}}>
                                  <div style={{background:dm?"#161920":"#e5e7eb",borderRadius:20,height:6,flex:1,overflow:"hidden"}}>
                                    <div style={{width:`${pct}%`,height:"100%",background:natColors[i%natColors.length],borderRadius:20,minWidth:cnt>0?6:0}}/>
                                  </div>
                                  <span style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280",whiteSpace:"nowrap"}}>{cnt} ({pct}%)</span>
                                </div>
                              </div>
                            </div>
                            <div style={{padding:"8px 14px 10px",borderTop:`1px dashed ${dm?"#2a2f3d":"#e5e7eb"}`,display:"flex",flexWrap:"wrap",gap:6}}>
                              {emps.map(r=>{
                                const sc=STATUS_COLORS[getStatus(r)];
                                return <span key={r.id} style={{background:sc.bg,color:sc.text,border:`1px solid ${sc.border}`,padding:"2px 9px",borderRadius:12,fontSize:11,fontWeight:600}}>{r.name}</span>;
                              })}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })()}

              {/* ── صف 3.5: داخل وخارج المملكة ── */}
              {selectedReports.has("location")&&(()=>{
                const inside  = records.filter(r=>r.outsideKingdom!=="نعم");
                const outside = records.filter(r=>r.outsideKingdom==="نعم");
                const insideEmp  = inside.filter(r=>r.type!=="مرافق");
                const outsideEmp = outside.filter(r=>r.type!=="مرافق");
                return (
                  <div style={card}>
                    <div style={titleStyle}>🌍 داخل وخارج المملكة</div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:18}}>
                      {[
                        {label:"🏠 داخل المملكة",count:inside.length,emp:insideEmp.length,dep:inside.length-insideEmp.length,color:"#16a34a",bg:dm?"#081a12":"#f0fdf4",border:dm?"#166534":"#86efac"},
                        {label:"✈️ خارج المملكة",count:outside.length,emp:outsideEmp.length,dep:outside.length-outsideEmp.length,color:"#dc2626",bg:dm?"#2a1515":"#fef2f2",border:dm?"#7f1d1d":"#fca5a5"},
                      ].map(c=>(
                        <div key={c.label} style={{background:c.bg,border:`1px solid ${c.border}`,borderRadius:12,padding:"16px",textAlign:"center"}}>
                          <div style={{fontSize:32,fontWeight:900,color:c.color}}>{c.count}</div>
                          <div style={{fontSize:14,fontWeight:700,color:c.color,marginTop:4}}>{c.label}</div>
                          <div style={{fontSize:12,color:dm?"#a0a8bb":"#6b7280",marginTop:6}}>({(records.length>0?(c.count/records.length*100):0).toFixed(1)}%)</div>
                          <div style={{display:"flex",justifyContent:"center",gap:10,marginTop:8,fontSize:11}}>
                            <span style={{background:dm?"#1e222b":"rgba(255,255,255,0.6)",padding:"2px 8px",borderRadius:8,color:dm?"#a0a8bb":"#374151"}}>👤 {c.emp} موظف</span>
                            <span style={{background:dm?"#1e222b":"rgba(255,255,255,0.6)",padding:"2px 8px",borderRadius:8,color:dm?"#a0a8bb":"#374151"}}>👨‍👩‍👧 {c.dep} مرافق</span>
                          </div>
                        </div>
                      ))}
                    </div>
                    {outside.length>0&&(
                      <>
                        <div style={{fontWeight:700,fontSize:13,color:dm?"#f0f2f7":"#dc2626",marginBottom:10,display:"flex",alignItems:"center",gap:8}}>
                          ✈️ المتواجدون خارج المملكة
                          <span style={{background:dm?"#2a1515":"#fef2f2",color:"#dc2626",padding:"2px 10px",borderRadius:10,fontSize:12}}>{outside.length}</span>
                        </div>
                        <div style={{display:"grid",gap:8}}>
                          {outside.map(r=>{
                            const st=getStatus(r),sc=STATUS_COLORS[st];
                            const cc=COMPANY_COLORS[r.company]||{bg:"#f9f9f9",text:"#374151",border:"#e5e7eb"};
                            return (
                              <div key={r.id} style={{background:dm?"#1e222b":"#fafafa",border:`1px solid ${dm?"#2a2f3d":"#e5e7eb"}`,borderRight:`4px solid ${sc.border}`,borderRadius:10,padding:"10px 14px",display:"flex",flexWrap:"wrap",gap:8,alignItems:"center",justifyContent:"space-between"}}>
                                <div>
                                  <div style={{fontWeight:600,fontSize:13,color:dm?"#f0f2f7":"#1e3a5f"}}>{r.name}</div>
                                  <div style={{display:"flex",gap:8,fontSize:11,color:dm?"#a0a8bb":"#6b7280",marginTop:3,flexWrap:"wrap"}}>
                                    <span>🪪 {r.iqamaNumber}</span>
                                    <span>🌍 {r.nationality}</span>
                                    {(r.jobTitle||r.notes)&&r.type!=="مرافق"&&<span>💼 {r.jobTitle||r.notes}</span>}
                                  </div>
                                </div>
                                <div style={{display:"flex",gap:6,alignItems:"center"}}>
                                  <span style={{background:cc.bg,color:cc.text,border:`1px solid ${cc.border}`,padding:"2px 8px",borderRadius:8,fontSize:11,fontWeight:600}}>{r.company}</span>
                                  <span style={{background:sc.bg,color:sc.text,border:`1px solid ${sc.border}`,padding:"2px 8px",borderRadius:8,fontSize:11,fontWeight:600}}>{st}</span>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </>
                    )}
                  </div>
                );
              })()}

              {/* ── صف 4: تقويم الانتهاء الشهري ── */}
              {selectedReports.has("calendar")&&<div style={card}>
                <div style={titleStyle}>📅 تقويم الإقامات المنتهية شهرياً (12 شهر قادم)</div>
                <div style={{display:"flex",gap:6,alignItems:"flex-end",height:140,overflowX:"auto",paddingBottom:4}}>
                  {monthlyData.map((d,i)=>{
                    const h = maxMonthly>0?Math.max((d.count/maxMonthly)*110,d.count>0?8:0):0;
                    const isPast = d.month < new Date(now.getFullYear(),now.getMonth(),1);
                    const isThisMonth = d.month.getMonth()===now.getMonth()&&d.month.getFullYear()===now.getFullYear();
                    const barColor = isPast?"#6b7280":isThisMonth?"#F5A800":d.count>3?"#dc2626":d.count>0?"#d97706":"#86efac";
                    return (
                      <div key={i} style={{flex:1,minWidth:44,display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
                        {d.count>0&&<div style={{fontSize:11,fontWeight:700,color:barColor}}>{d.count}</div>}
                        <div style={{width:"100%",height:`${h}px`,background:barColor,borderRadius:"6px 6px 0 0",minHeight:d.count>0?8:2}}/>
                        <div style={{fontSize:10,color:dm?"#a0a8bb":"#6b7280",textAlign:"center",whiteSpace:"nowrap"}}>{d.label}</div>
                      </div>
                    );
                  })}
                </div>
                <div style={{display:"flex",gap:16,marginTop:12,flexWrap:"wrap"}}>
                  {[{c:"#F5A800",l:"الشهر الحالي"},{c:"#dc2626",l:"أكثر من 3"},{c:"#d97706",l:"1-3 إقامات"},{c:"#86efac",l:"لا انتهاء"}].map(x=>(
                    <div key={x.l} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:dm?"#a0a8bb":"#6b7280"}}>
                      <div style={{width:10,height:10,borderRadius:2,background:x.c}}/>
                      {x.l}
                    </div>
                  ))}
                </div>
              </div>}

              {/* ── صف 5: التكاليف الشهرية ── */}
              {selectedReports.has("costs")&&<div style={card}>
                <div style={titleStyle}>💰 تكاليف التجديد المتوقعة شهرياً (السنة القادمة)</div>
                <div style={{display:"flex",gap:6,alignItems:"flex-end",height:150,overflowX:"auto",paddingBottom:4}}>
                  {monthlyCostData.map((d,i)=>{
                    const h = maxCost>0?Math.max((d.cost/maxCost)*120,d.cost>0?8:0):0;
                    return (
                      <div key={i} style={{flex:1,minWidth:44,display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
                        {d.cost>0&&<div style={{fontSize:9,fontWeight:700,color:"#16a34a",textAlign:"center"}}>{(d.cost/1000).toFixed(1)}k</div>}
                        <div style={{width:"100%",height:`${h}px`,background:d.cost>10000?"#dc2626":d.cost>5000?"#d97706":"#16a34a",borderRadius:"6px 6px 0 0",minHeight:d.cost>0?8:2}}/>
                        <div style={{fontSize:10,color:dm?"#a0a8bb":"#6b7280",textAlign:"center",whiteSpace:"nowrap"}}>{d.label}</div>
                      </div>
                    );
                  })}
                </div>
                <div style={{marginTop:12,background:dm?"#081a12":"#f0fdf4",border:`1px solid ${dm?"#166534":"#86efac"}`,borderRadius:8,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <span style={{fontSize:13,color:dm?"#6ee7b7":"#15803d",fontWeight:600}}>إجمالي التكاليف المتوقعة للسنة القادمة</span>
                  <span style={{fontSize:18,fontWeight:900,color:dm?"#6ee7b7":"#15803d"}}>{fmt(monthlyCostData.reduce((s,d)=>s+d.cost,0))}</span>
                </div>
              </div>}

              {/* ── صف 6: الإقامات المجددة ── */}
              {selectedReports.has("renewed")&&<div style={card}>
                <div style={titleStyle}>✅ الإقامات التي تم تجديدها</div>
                {renewedRecords.length===0?(
                  <div style={{textAlign:"center",padding:"30px 0",color:dm?"#a0a8bb":"#6b7280"}}>
                    <div style={{fontSize:36}}>📋</div>
                    <p style={{marginTop:8}}>لم يتم تسجيل أي تجديد بعد</p>
                    <p style={{fontSize:12,marginTop:4}}>استخدم زر 🔄 في كل بطاقة لتسجيل التجديد</p>
                  </div>
                ):(
                  <>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12,marginBottom:16}}>
                      {[
                        {l:"إجمالي المجددين",v:renewedRecords.length,c:"#16a34a",bg:dm?"#081a12":"#f0fdf4",bdr:dm?"#166534":"#86efac"},
                        {l:"موظفون مجددون",v:renewedRecords.filter(r=>r.type!=="مرافق").length,c:"#2563eb",bg:dm?"#0f1a2e":"#eff6ff",bdr:dm?"#1e40af":"#bfdbfe"},
                        {l:"مرافقون مجددون",v:renewedRecords.filter(r=>r.type==="مرافق").length,c:"#7c3aed",bg:dm?"#1a1128":"#f3e8ff",bdr:dm?"#5b21b6":"#c4b5fd"},
                        {l:"إجمالي التكلفة",v:fmt(totalRenewedCost),c:"#15803d",bg:dm?"#081a12":"#f0fdf4",bdr:dm?"#166534":"#86efac",isText:true},
                      ].map(x=>(
                        <div key={x.l} style={{background:x.bg,border:`1px solid ${x.bdr}`,borderRadius:10,padding:"12px",textAlign:"center"}}>
                          <div style={{fontSize:x.isText?14:24,fontWeight:800,color:x.c}}>{x.v}</div>
                          <div style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280",marginTop:3}}>{x.l}</div>
                        </div>
                      ))}
                    </div>
                    <div style={{display:"grid",gap:8,maxHeight:280,overflowY:"auto"}}>
                      {renewedRecords.sort((a,b)=>new Date(b.lastRenewalDate)-new Date(a.lastRenewalDate)).map(r=>{
                        const cc=COMPANY_COLORS[r.company]||{bg:"#f9f9f9",text:"#374151",border:"#e5e7eb"};
                        return (
                          <div key={r.id} style={{background:dm?"#1e222b":"#f9fafb",borderRadius:10,padding:"11px 14px",border:`1px solid ${dm?"#2a2f3d":"#e5e7eb"}`,display:"flex",flexWrap:"wrap",gap:10,alignItems:"center",justifyContent:"space-between"}}>
                            <div>
                              <div style={{fontWeight:600,fontSize:13,color:dm?"#f0f2f7":"#1e3a5f"}}>{r.name}</div>
                              <div style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280",marginTop:2,display:"flex",gap:8,flexWrap:"wrap"}}>
                                <span style={{background:cc.bg,color:cc.text,padding:"0 6px",borderRadius:5}}>{r.company}</span>
                                {r.lastRenewalNote&&<span>📝 {r.lastRenewalNote}</span>}
                              </div>
                            </div>
                            <div style={{textAlign:"left"}}>
                              <div style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280"}}>تاريخ التجديد</div>
                              <div style={{fontWeight:700,fontSize:13,color:dm?"#6ee7b7":"#15803d"}}>{new Date(r.lastRenewalDate).toLocaleDateString("ar-SA",{year:"numeric",month:"long",day:"numeric"})}</div>
                              <div style={{fontSize:11,color:dm?"#a0a8bb":"#6b7280"}}>ينتهي: {r.expiryDate?new Date(r.expiryDate).toLocaleDateString("ar-SA"):"—"}</div>
                            </div>
                            <div style={{fontWeight:700,color:dm?"#6ee7b7":"#15803d",fontSize:13}}>{fmt(r.type!=="مرافق"?2588:1200)}</div>
                          </div>
                        );
                      })}
                    </div>
                  </>
                )}
              </div>}

            </div>
          );
        })()}

      </div>
    </div>
  );
}
