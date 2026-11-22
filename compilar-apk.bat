@echo off
setlocal
cd /d "%~dp0"
echo ============================================================
echo   INSUAMINCA - Compilador de APK Android en tu PC
echo ============================================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [ERROR] No se encontro Node.js. Descargalo desde https://nodejs.org/
  pause
  exit /b 1
)

where java >nul 2>nul
if errorlevel 1 (
  if exist "%ProgramFiles%\Android\Android Studio\jbr\bin\java.exe" (
    set "JAVA_HOME=%ProgramFiles%\Android\Android Studio\jbr"
    set "PATH=%JAVA_HOME%\bin;%PATH%"
    echo [INFO] Java detectado en Android Studio JBR.
  ) else if exist "%ProgramFiles%\Android\Android Studio\jre\bin\java.exe" (
    set "JAVA_HOME=%ProgramFiles%\Android\Android Studio\jre"
    set "PATH=%JAVA_HOME%\bin;%PATH%"
    echo [INFO] Java detectado en Android Studio JRE.
  ) else (
    echo [ERROR] No se encontro Java JDK 17 ni Android Studio.
    echo Instala Android Studio o Java JDK 17 ^(Eclipse Temurin^) y vuelve a intentar.
    pause
    exit /b 1
  )
)

if "%ANDROID_HOME%"=="" (
  if exist "%LOCALAPPDATA%\Android\Sdk" (
    set "ANDROID_HOME=%LOCALAPPDATA%\Android\Sdk"
    echo [INFO] Android SDK detectado en %LOCALAPPDATA%\Android\Sdk
  )
)

echo [1/4] Instalando dependencias Capacitor...
call npm install
if errorlevel 1 goto :error

if not exist android (
  echo [2/4] Generando proyecto nativo Android...
  call npx cap add android
  if errorlevel 1 goto :error
)

echo [3/4] Sincronizando compilado web ^(www/^), permisos e iconos...
call npx cap sync android
if errorlevel 1 goto :error
call node scripts/patch-android.mjs
if errorlevel 1 goto :error

if not exist "android\local.properties" (
  if "%ANDROID_HOME%"=="" (
    echo [ERROR] No se detecto el Android SDK en %LOCALAPPDATA%\Android\Sdk.
    echo Abre Android Studio una vez y completa el asistente inicial ^(SDK Setup^).
    pause
    exit /b 1
  )
)

echo [4/4] Compilando archivo APK instalable...
cd android
call gradlew.bat assembleDebug --no-daemon
if errorlevel 1 goto :error
cd ..

copy /Y "android\app\build\outputs\apk\debug\app-debug.apk" "INSUAMINCA-Medidas-Pesos.apk" >nul

echo.
echo ============================================================
echo   APK GENERADO EXITOSAMENTE:
echo   INSUAMINCA-Medidas-Pesos.apk ^(en esta misma carpeta^)
echo ============================================================
pause
exit /b 0

:error
echo.
echo [ERROR] Ocurrio un fallo durante la compilacion del APK.
echo Verifica tener Android Studio ^(SDK Platform 34^) y Java JDK 17 instalados.
pause
exit /b 1
