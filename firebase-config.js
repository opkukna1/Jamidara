// ============================================================
// Firebase Configuration
// ============================================================

import {
  initializeApp
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";

import {
  getFirestore
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

import {
  getAuth
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";


// ============================================================
// Firebase Project Configuration
// ============================================================

export const firebaseConfig = {

  apiKey:
    "AIzaSyB6vFcpqboymXfhn6QesgAZEYOE5EBbyh4",

  authDomain:
    "instaquiz-9cc2f.firebaseapp.com",

  projectId:
    "instaquiz-9cc2f",

  storageBucket:
    "instaquiz-9cc2f.firebasestorage.app",

  messagingSenderId:
    "265258409167",

  appId:
    "1:265258409167:web:d4cd66b7bb157b8f6b6bba",

  measurementId:
    "G-TDTKQM6VQ4"
};


// ============================================================
// Initialize Firebase
// ============================================================

const app =
  initializeApp(firebaseConfig);


// ============================================================
// Firestore
// ============================================================

export const db =
  getFirestore(app);


// ============================================================
// Firebase Authentication
// ============================================================

export const auth =
  getAuth(app);
