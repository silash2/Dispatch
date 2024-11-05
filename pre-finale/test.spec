# interface_0.spec
from PyInstaller.__main__ import run
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Chemin du script principal
script_path = 'pre-finale/interface_0.py'

# Chemin des dossiers à inclure
pathex = ['.', './cession', './data', './Dispatch']

# Collecte des modules et fichiers de données nécessaires
hiddenimports = collect_submodules('Dispatch') + collect_submodules('cession') + collect_submodules('data')
datas = collect_data_files('cession') + collect_data_files('data') + collect_data_files('Dispatch')

# Configuration de l'analyse
a = Analysis(
    [script_path],
    pathex=pathex,
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False
)

# Création du fichier PYZ
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# Configuration de l'exécutable
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='interface_0',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True
)

# Collecte des fichiers et des données
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='interface_0'
)
