// FIX: Changed firebase imports to use scoped packages (@firebase/*) to resolve module export errors.
import { initializeApp } from "@firebase/app";
import { getFirestore, doc, setDoc, getDoc } from "@firebase/firestore";
import type { EmployeeRecord } from './types';

export const firebaseConfig = {
  apiKey: "AIzaSyB7VinbNy7AwyJdg5bTBrujvTZDvJqoVlY",
  authDomain: "yassin-584cc.firebaseapp.com",
  projectId: "yassin-584cc",
  storageBucket: "yassin-584cc.firebasestorage.app",
  messagingSenderId: "320597339073",
  appId: "1:320597339073:web:95824697390485d8fbdd40",
  databaseURL: "https://yassin-584cc-default-rtdb.firebaseio.com"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);

console.log("✅ Firebase Initialized (centralized)");

export async function saveEmployeeData(employeeId: string, data: Partial<EmployeeRecord>): Promise<void> {
  try {
    await setDoc(doc(db, "employees", employeeId), data, { merge: true });
    console.log(`Employee data for ${employeeId} saved successfully to Firestore.`);
  } catch (error) {
    console.error("Error saving employee data to Firestore:", error);
  }
}

export async function getEmployeeData(employeeId: string): Promise<EmployeeRecord | null> {
  try {
    const docRef = doc(db, "employees", employeeId);
    const docSnap = await getDoc(docRef);

    if (docSnap.exists()) {
      console.log("Employee data retrieved from Firestore:", docSnap.data());
      return docSnap.data() as EmployeeRecord;
    } else {
      console.log("No such employee document in Firestore!");
      return null;
    }
  } catch (error) {
    console.error("Error getting employee data from Firestore:", error);
    return null;
  }
}
