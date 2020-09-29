# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['C:\\Users\\MSIWillard\\Documents\\Python\\my_project\\payslip_sender\\payslip_sender_1.0.0\\payslip_sender.py'],
             pathex=['C:\\Users\\MSIWillard\\Documents\\Python\\to_compile'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='payslip_sender',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False , icon='C:\\Users\\MSIWillard\\Documents\\Python\\my_project\\payslip_sender\\payslip_sender_1.0.0\\logo.ico')
