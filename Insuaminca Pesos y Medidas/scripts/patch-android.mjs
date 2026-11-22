import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const root = process.cwd();
const androidDir = path.join(root, 'android');
const androidMainDir = path.join(androidDir, 'app', 'src', 'main');
const manifestPath = path.join(androidMainDir, 'AndroidManifest.xml');

// 1. Auto-detect Android SDK and write android/local.properties if missing
if (fs.existsSync(androidDir)) {
  const localPropsPath = path.join(androidDir, 'local.properties');
  const candidates = [
    process.env.ANDROID_HOME,
    process.env.ANDROID_SDK_ROOT,
    process.env.LOCALAPPDATA ? path.join(process.env.LOCALAPPDATA, 'Android', 'Sdk') : '',
    path.join(os.homedir(), 'AppData', 'Local', 'Android', 'Sdk'),
    path.join(os.homedir(), 'Library', 'Android', 'Sdk'),
    path.join(os.homedir(), 'Android', 'Sdk'),
  ].filter(Boolean);

  const detectedSdk = candidates.find((p) => p && fs.existsSync(p));
  if (detectedSdk && !fs.existsSync(localPropsPath)) {
    const escaped = detectedSdk.replace(/\\/g, '\\\\').replace(/:/g, '\\:');
    fs.writeFileSync(localPropsPath, `sdk.dir=${escaped}\n`, 'utf-8');
    console.log('[patch-android] SDK Android detectado en:', detectedSdk);
  }
}

// 2. Ensure Camera & Network permissions in AndroidManifest.xml
if (fs.existsSync(manifestPath)) {
  let manifest = fs.readFileSync(manifestPath, 'utf-8');
  const additions = [
    '<uses-permission android:name="android.permission.CAMERA" />',
    '<uses-permission android:name="android.permission.ACCESS_NETWORK_STATE" />',
    '<uses-feature android:name="android.hardware.camera" android:required="false" />',
  ].filter((line) => !manifest.includes(line));

  if (additions.length > 0 && manifest.includes('<application')) {
    manifest = manifest.replace('<application', additions.join('\n    ') + '\n\n    <application');
    fs.writeFileSync(manifestPath, manifest, 'utf-8');
    console.log('[patch-android] Permisos de cámara y red añadidos en AndroidManifest.xml');
  }
}

// 3. Apply INSUAMINCA corporate icons and adaptive background color
const icon512 = path.join(root, 'www', 'pwa-512x512.png');
const icon192 = path.join(root, 'www', 'pwa-192x192.png');
const resDir = path.join(androidMainDir, 'res');

if (fs.existsSync(resDir) && fs.existsSync(icon512)) {
  const mipmaps = [
    { dir: 'mipmap-mdpi', src: fs.existsSync(icon192) ? icon192 : icon512 },
    { dir: 'mipmap-hdpi', src: fs.existsSync(icon192) ? icon192 : icon512 },
    { dir: 'mipmap-xhdpi', src: icon512 },
    { dir: 'mipmap-xxhdpi', src: icon512 },
    { dir: 'mipmap-xxxhdpi', src: icon512 },
  ];

  for (const m of mipmaps) {
    const targetDir = path.join(resDir, m.dir);
    if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });
    for (const fileName of ['ic_launcher.png', 'ic_launcher_round.png', 'ic_launcher_foreground.png']) {
      fs.copyFileSync(m.src, path.join(targetDir, fileName));
    }
  }

  const valuesDir = path.join(resDir, 'values');
  if (!fs.existsSync(valuesDir)) fs.mkdirSync(valuesDir, { recursive: true });
  fs.writeFileSync(
    path.join(valuesDir, 'ic_launcher_background.xml'),
    '<?xml version="1.0" encoding="utf-8"?>\n<resources>\n    <color name="ic_launcher_background">#074E2B</color>\n</resources>\n',
    'utf-8'
  );

  const anydpiDir = path.join(resDir, 'mipmap-anydpi-v26');
  for (const xmlName of ['ic_launcher.xml', 'ic_launcher_round.xml']) {
    const xmlPath = path.join(anydpiDir, xmlName);
    if (fs.existsSync(xmlPath)) {
      const xml = fs.readFileSync(xmlPath, 'utf-8').replace(/@mipmap\/ic_launcher_background/g, '@color/ic_launcher_background');
      fs.writeFileSync(xmlPath, xml, 'utf-8');
    }
  }

  console.log('[patch-android] Iconos corporativos INSUAMINCA aplicados a Android');
}
