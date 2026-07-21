# Proyan — Guía para publicarla gratis

Este proyecto es tu app Proyan lista para publicar como Web App (PWA).
Una vez publicada, cualquier mecánico podrá abrir un link desde Chrome
en su Android y "instalarla" como si fuera una app nativa.

## Lo que vas a necesitar (todo gratis)

- Una cuenta de [GitHub](https://github.com)
- Una cuenta de [Firebase](https://console.firebase.google.com) (guarda los datos: órdenes, mecánicos, código de acceso)
- Una cuenta de [Vercel](https://vercel.com) (aloja la app y te da el link público)

---

## Paso 1 — Crear la base de datos (Firebase)

1. Entrá a https://console.firebase.google.com y creá un proyecto nuevo (podés
   llamarlo "proyan").
2. En el menú de la izquierda: **Compilación → Firestore Database → Crear base
   de datos**. Elegí **"Empezar en modo de prueba"** (para arrancar rápido).
3. Andá a **Configuración del proyecto** (ícono de engranaje, arriba a la
   izquierda) → bajá hasta **"Tus apps"** → tocá el ícono **`</>`** (Web) →
   registrá la app (cualquier nombre) → Firebase te va a mostrar un bloque de
   código con un objeto `firebaseConfig`.
4. Copiá esos valores y pegalos en el archivo `src/firebase.js` de este
   proyecto, reemplazando el bloque `firebaseConfig` que dice `"TU_API_KEY"`,
   etc.

⚠️ El "modo de prueba" de Firestore deja la base de datos abierta por 30
días. Antes de que se cumpla el plazo (o directamente ahora si querés
hacerlo bien), andá a **Firestore Database → Reglas** y poné reglas que
exijan tu código de acceso, o al menos algo básico como:

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /proyan/{document} {
      allow read, write: if true; // reemplazar por reglas más estrictas cuando puedas
    }
  }
}
```

## Paso 2 — Subir el código a GitHub

1. Creá un repositorio nuevo en GitHub (por ejemplo "proyan-app").
2. Subí el contenido de esta carpeta al repositorio (podés arrastrar los
   archivos desde la web de GitHub, o usar Git si lo conocés).

## Paso 3 — Publicar en Vercel

1. Entrá a https://vercel.com y creá una cuenta (podés usar tu cuenta de
   GitHub para entrar directo).
2. Tocá **"Add New" → "Project"** y elegí el repositorio que subiste.
3. Vercel detecta automáticamente que es un proyecto Vite/React. No hace
   falta cambiar nada — tocá **"Deploy"**.
4. En 1-2 minutos te va a dar un link público, algo como
   `https://proyan-app.vercel.app`. **Ese es el link que le vas a pasar a
   tus mecánicos.**

## Paso 4 — Instalarla en un Android

1. Abrí el link en **Chrome** desde el celular.
2. Tocá los tres puntitos (⋮) arriba a la derecha.
3. Tocá **"Agregar a pantalla de inicio"** o **"Instalar app"**.
4. Les queda un ícono como cualquier otra app, y se abre en pantalla
   completa (sin la barra del navegador).

## Notas importantes

- **El código de acceso** (02120512, o el que hayas cambiado desde el panel
  Admin → Equipo) sigue funcionando igual: se lo pasás vos a cada mecánico.
- **Los íconos de la app** (`/public/icon-192.png` y `/public/icon-512.png`)
  no están incluidos — la app funciona igual sin ellos, pero si querés un
  ícono lindo en la pantalla de inicio, agregá esos dos archivos PNG a la
  carpeta `public/`.
- Si en algún momento querés pasar esto a una app nativa de verdad
  (archivo .apk instalable), este mismo código y esta misma base de datos
  de Firebase sirven como base para un desarrollador que trabaje con
  React Native o Flutter.
