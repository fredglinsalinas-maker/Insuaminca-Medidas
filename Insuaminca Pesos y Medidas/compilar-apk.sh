#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

echo "============================================================"
echo "  INSUAMINCA - Compilador de APK Android en tu PC"
echo "============================================================"

if ! command -v node >/dev/null 2>&1; then
  echo "[ERROR] No se encontro Node.js. Descargalo desde https://nodejs.org/"
  exit 1
fi

if ! command -v java >/dev/null 2>&1; then
  if [ -d "/Applications/Android Studio.app/Contents/jbr/Contents/Home/bin" ]; then
    export JAVA_HOME="/Applications/Android Studio.app/Contents/jbr/Contents/Home"
    export PATH="$JAVA_HOME/bin:$PATH"
  else
    echo "[ERROR] No se encontro Java JDK 17 ni Android Studio."
    exit 1
  fi
fi

if [ -z "$ANDROID_HOME" ]; then
  if [ -d "$HOME/Library/Android/Sdk" ]; then
    export ANDROID_HOME="$HOME/Library/Android/Sdk"
  elif [ -d "$HOME/Android/Sdk" ]; then
    export ANDROID_HOME="$HOME/Android/Sdk"
  fi
fi

echo "[1/4] Instalando dependencias Capacitor..."
npm install

if [ ! -d "android" ]; then
  echo "[2/4] Generando proyecto nativo Android..."
  npx cap add android
fi

echo "[3/4] Sincronizando compilado web (www/), permisos e iconos..."
npx cap sync android
node scripts/patch-android.mjs

echo "[4/4] Compilando archivo APK instalable..."
cd android
chmod +x gradlew
./gradlew assembleDebug --no-daemon
cd ..

cp "android/app/build/outputs/apk/debug/app-debug.apk" "INSUAMINCA-Medidas-Pesos.apk"

echo ""
echo "============================================================"
echo "  APK GENERADO EXITOSAMENTE:"
echo "  INSUAMINCA-Medidas-Pesos.apk (en esta misma carpeta)"
echo "============================================================"
