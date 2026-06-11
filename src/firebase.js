import { initializeApp } from "firebase/app";
import { getDatabase } from "firebase/database";
import { getAuth } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyBiMLOSZjztaet_41J1Zr-AwhtrB5f_sFY",
  authDomain: "anjal-iqama.firebaseapp.com",
  databaseURL: "https://anjal-iqama-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "anjal-iqama",
  storageBucket: "anjal-iqama.firebasestorage.app",
  messagingSenderId: "397150889353",
  appId: "1:397150889353:web:2c0f32c8ab138ffce02d58",
  measurementId: "G-CJTGKRCEYP"
};

const firebaseApp = initializeApp(firebaseConfig);
export const db  = getDatabase(firebaseApp);
export const auth = getAuth(firebaseApp);
export const DB_PATH = "iqama_records";
