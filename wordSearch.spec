# -*- mode: python -*-

block_cipher = None


a = Analysis(['wordSearch.py'],
             pathex=['/Users/alexaufauvre/Workspace/WordSearch'],
             binaries=[],
             datas=[],
             hiddenimports=['docx', 'pptx'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
for d in a.datas:
    if 'pyconfig' in d[0]:
        a.datas.remove(d)
        break

a.datas += [('img/logo-kanbios-resized.png','/Users/alexaufauvre/Workspace/WordSearch/img/logo-kanbios-resized.png', 'img')]
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='wordSearch',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
app = BUNDLE(exe,
             name='wordSearch.app',
             icon=None,
             bundle_identifier=None)
