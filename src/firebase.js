import { initializeApp } from "firebase/app";
import { getFirestore, enableIndexedDbPersistence } from "firebase/firestore";

// 1. Andá a https://console.firebase.google.com
// 2. Creá un proyecto nuevo (gratis)
// 3. Dentro del proyecto: "Compilación" -> "Firestore Database" -> "Crear base de datos"
//    (elegí "modo de prueba" para arrancar rápido)
// 4. En el ícono de engranaje -> "Configuración del proyecto" -> bajá hasta "Tus apps"
//    -> creá una app web (</>) -> copiá el objeto firebaseConfig y pegalo acá abajo,
//    reemplazando todo este bloque.
const firebaseConfig = {
  apiKey: "AIzaSyD8onxbUPAwLEYzPeaa3MSoo_P3Z8AYhqU",
  authDomain: "proyan-1b10c.firebaseapp.com",
  projectId: "proyan-1b10c",
  storageBucket: "proyan-1b10c.firebasestorage.app",
  messagingSenderId: "744742979254",
  appId: "1:744742979254:web:8e41401b2dfbaeb338c817",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);

// Permite que la app siga funcionando y guardando datos sin señal,
// y sincronice apenas vuelva la conexión.
try {
  enableIndexedDbPersistence(db);
} catch (e) {
  console.warn("Persistencia offline no disponible en este navegador:", e);
}
