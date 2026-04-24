# backup_automatico.py
import shutil
import os
from datetime import datetime, timedelta
from pathlib import Path

# 🔍 Rutas (funciona en Mac y Windows)
BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "instance" / "hotel_adams.db"
BACKUP_DIR = BASE_DIR / "backups"

def crear_backup():
    # Crear carpeta de backups si no existe
    BACKUP_DIR.mkdir(exist_ok=True)
    
    if not DB_PATH.exists():
        print("❌ No se encontró la base de datos")
        return

    # Nombre con fecha y hora exacta
    ahora = datetime.now()
    nombre = f"hotel_adams_{ahora.strftime('%Y%m%d_%H%M')}.db"
    destino = BACKUP_DIR / nombre

    try:
        shutil.copy2(DB_PATH, destino)
        print(f"✅ Backup creado: {nombre}")
    except Exception as e:
        print(f"❌ Error al copiar: {e}")

    # 🧹 Limpiar backups antiguos (más de 7 días)
    limite = ahora - timedelta(days=7)
    eliminados = 0
    for archivo in BACKUP_DIR.glob("*.db"):
        if datetime.fromtimestamp(archivo.stat().st_mtime) < limite:
            archivo.unlink()
            eliminados += 1
            
    if eliminados > 0:
        print(f"🗑️ Se eliminaron {eliminados} backups antiguos para ahorrar espacio")

if __name__ == "__main__":
    crear_backup()