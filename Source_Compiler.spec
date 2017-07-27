# -*- mode: python -*-

block_cipher = None


a = Analysis(['Source_Compiler.py', 'Database_Manipulation.py', 'Google_API_Manipulation.py', 'CurrencyConverterNew.py', 'write_log.py', 'Email_Notifications.py', 'client_secret.json'],
             pathex=['C:\\Users\\Stephen\\Documents\\GitHub\\email-parsing'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='Source_Compiler',
          debug=False,
          strip=False,
          upx=True,
          console=True )
