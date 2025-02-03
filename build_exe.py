import PyInstaller.__main__

PyInstaller.__main__.run([
    'switch.py',
    '--onefile',
    '--windowed',
    '--uac-admin',
    '--icon=app.ico',  # Optional: Add this line if you have an icon
    '--name=SmartTrade_Patch_Switcher',
    '--add-data=admin_manifest.xml;.'
]) 