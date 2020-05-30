# -*- mode: python -*-

block_cipher = None


a = Analysis(['app.py'],
             pathex=['D:\\certificate-generator'],  # Set absolute path here
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)

# Set the absolute path to resources here
a.datas += [('blank-light.jpg', 'D:\\certificate-generator\\blank-light.jpg', 'DATA'),
            ('blank-dark.jpg', 'D:\\certificate-generator\\blank-dark.jpg', 'DATA'),
            ('psicon.ico', 'D:\\certificate-generator\\psicon.ico', 'DATA'),
            ('pslogo.png', 'D:\\certificate-generator\\pslogo.png', 'DATA'),
            ('Roboto-Light.ttf', 'D:\\certificate-generator\\Roboto-Light.ttf', 'DATA'),
            ('Roboto-Medium.ttf', 'D:\\certificate-generator\\Roboto-Medium.ttf', 'DATA')]

pyz = PYZ(a.pure, a.zipped_data,
          cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='app',
          debug=False,
          strip=False,
          upx=True,
          console=False , version='version.txt', icon='psicon.ico')
