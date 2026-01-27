#!/bin/bash
set -e

APP_NAME="SistemaAsignacionMonitores"
VERSION="1.0"

echo "🚀 Creando AppImage para $APP_NAME v$VERSION"

if [ ! -f "dist/asignacion_monitores" ]; then
    echo "❌ ERROR: No se encuentra dist/asignacion_monitores"
    exit 1
fi

echo "📁 Creando estructura AppDir..."
rm -rf AppDir
mkdir -p AppDir/usr/bin
mkdir -p AppDir/usr/share/applications
mkdir -p AppDir/usr/share/icons/hicolor/256x256/apps

echo "📋 Copiando ejecutable..."
cp dist/asignacion_monitores AppDir/usr/bin/
chmod +x AppDir/usr/bin/asignacion_monitores

echo "🎨 Copiando icono..."
if [ -f "icon.png" ]; then
    cp icon.png AppDir/usr/share/icons/hicolor/256x256/apps/$APP_NAME.png
    cp icon.png AppDir/$APP_NAME.png
fi

echo "📄 Creando archivo .desktop..."
cat > AppDir/usr/share/applications/$APP_NAME.desktop << EOF
[Desktop Entry]
Name=Sistema de Asignación de Monitores
Comment=Gestión y asignación automática de monitores académicos
Exec=asignacion_monitores
Icon=$APP_NAME
Type=Application
Categories=Office;Education;
Terminal=false
StartupWMClass=asignacion_monitores
EOF

echo "🔗 Creando AppRun..."
cat > AppDir/AppRun << 'EOF'
#!/bin/bash
HERE="$(dirname "$(readlink -f "${0}")")"
export PATH="${HERE}/usr/bin:${PATH}"
export LD_LIBRARY_PATH="${HERE}/usr/lib:${LD_LIBRARY_PATH}"
exec "${HERE}/usr/bin/asignacion_monitores" "$@"
EOF
chmod +x AppDir/AppRun

cp AppDir/usr/share/applications/$APP_NAME.desktop AppDir/

if [ ! -f "appimagetool-x86_64.AppImage" ]; then
    echo "📥 Descargando appimagetool..."
    wget https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage
    chmod +x appimagetool-x86_64.AppImage
fi

echo "🎁 Generando AppImage..."
ARCH=x86_64 ./appimagetool-x86_64.AppImage AppDir $APP_NAME-$VERSION-x86_64.AppImage

chmod +x $APP_NAME-$VERSION-x86_64.AppImage

echo ""
echo "✅ ¡AppImage creado exitosamente!"
echo "📦 Archivo: $APP_NAME-$VERSION-x86_64.AppImage"
echo "📊 Tamaño: $(du -h $APP_NAME-$VERSION-x86_64.AppImage | cut -f1)"
