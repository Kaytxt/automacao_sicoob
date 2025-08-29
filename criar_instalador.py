#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para criar um instalador profissional da Automação Sicoob
Requer: pip install pyinstaller nsis (opcional)
"""

import os
import subprocess
import sys
from pathlib import Path

def verificar_dependencias():
    """Verifica e instala dependências necessárias"""
    print("🔍 Verificando dependências...")
    
    dependencias = [
        'pyinstaller',
        'openpyxl', 
        'pandas',
        'numpy',
        'chardet'
    ]
    
    for dep in dependencias:
        try:
            __import__(dep)
            print(f"✅ {dep} - OK")
        except ImportError:
            print(f"📦 Instalando {dep}...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', dep])
            print(f"✅ {dep} - Instalado")

def limpar_builds_anteriores():
    """Remove builds anteriores"""
    print("🧹 Limpando builds anteriores...")
    
    dirs_para_remover = ['build', 'dist', '__pycache__']
    for dir_name in dirs_para_remover:
        if os.path.exists(dir_name):
            import shutil
            shutil.rmtree(dir_name)
            print(f"🗑️ Removido: {dir_name}")
    
    # Remove arquivos .spec
    for spec_file in Path('.').glob('*.spec'):
        spec_file.unlink()
        print(f"🗑️ Removido: {spec_file}")

def criar_arquivo_spec():
    """Cria arquivo de especificação personalizado"""
    spec_content = """
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['automacao_extrato.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Automacao_Gransoft.xlsx', '.'),
    ],
    hiddenimports=[
        'openpyxl',
        'chardet', 
        'pandas',
        'numpy',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        're',
        'datetime',
        'os'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'PIL', 
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Automacao_Sicoob',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt'
)
"""
    
    with open('automacao_sicoob.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content.strip())
    
    print("📄 Arquivo .spec criado: automacao_sicoob.spec")

def criar_version_info():
    """Cria arquivo de informações da versão"""
    version_content = """
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x4,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        '040904B0',
        [StringStruct('CompanyName', 'Automação Sicoob'),
        StringStruct('FileDescription', 'Processamento Automático de Extratos Sicoob'),
        StringStruct('FileVersion', '1.0.0.0'),
        StringStruct('InternalName', 'automacao_sicoob'),
        StringStruct('LegalCopyright', '© 2025 Automação Sicoob'),
        StringStruct('OriginalFilename', 'Automacao_Sicoob.exe'),
        StringStruct('ProductName', 'Automação Sicoob'),
        StringStruct('ProductVersion', '1.0.0.0')])
      ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""
    
    with open('version_info.txt', 'w', encoding='utf-8') as f:
        f.write(version_content.strip())
    
    print("📄 Arquivo version_info.txt criado")

def gerar_executavel():
    """Gera o executável usando PyInstaller"""
    print("🚀 Gerando executável...")
    
    # Comando PyInstaller otimizado
    cmd = [
        'pyinstaller',
        '--clean',
        '--noconfirm', 
        'automacao_sicoob.spec'
    ]
    
    try:
        subprocess.run(cmd, check=True)
        print("✅ Executável gerado com sucesso!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao gerar executável: {e}")
        return False

def criar_readme():
    """Cria arquivo README para distribuição"""
    readme_content = """
# 🚀 Automação Sicoob - Processamento de Extratos

## 📋 COMO USAR:

1. **Execute o programa:**
   - Clique duas vezes em `Automacao_Sicoob.exe`

2. **Selecione seu extrato:**
   - Escolha o arquivo de extrato do Sicoob (.xlsx ou .csv)

3. **Aguarde o processamento:**
   - O programa consolidará automaticamente as transações
   - Apenas valores de débito serão incluídos
   - Valores de crédito serão ignorados

4. **Verifique o resultado:**
   - A planilha `Automacao_Gransoft.xlsx` será atualizada
   - Formatação original preservada
   - Dados ordenados por data

## ⚠️ REQUISITOS:

- Windows 10/11 (64-bit)
- Arquivo `Automacao_Gransoft.xlsx` na mesma pasta
- Extrato do Sicoob em formato .xlsx ou .csv

## 🆘 SUPORTE:

- Certifique-se que o extrato é do Sicoob
- Mantenha os arquivos na mesma pasta
- Execute como administrador se necessário

## 📱 CONTATO:

Para suporte ou dúvidas, entre em contato com o desenvolvedor.

Versão: 1.0
Data: 2025
"""
    
    # Salva na pasta dist para distribuição
    os.makedirs('dist', exist_ok=True)
    with open('dist/LEIA-ME.txt', 'w', encoding='utf-8') as f:
        f.write(readme_content.strip())
    
    print("📄 README criado: dist/LEIA-ME.txt")

def copiar_planilha_template():
    """Copia a planilha template para a pasta de distribuição"""
    import shutil
    
    if os.path.exists('Automacao_Gransoft.xlsx'):
        shutil.copy2('Automacao_Gransoft.xlsx', 'dist/')
        print("📋 Planilha template copiada para dist/")
    else:
        print("⚠️ Planilha template não encontrada!")

def main():
    """Função principal"""
    print("=" * 50)
    print("🏗️  GERADOR DE EXECUTÁVEL - AUTOMAÇÃO SICOOB")
    print("=" * 50)
    print()
    
    # Verifica se o código fonte existe
    if not os.path.exists('automacao_extrato.py'):
        print("❌ Arquivo 'automacao_extrato.py' não encontrado!")
        print("   Certifique-se de estar na pasta correta.")
        input("Pressione ENTER para sair...")
        return
    
    try:
        # Passo 1: Verificar dependências
        verificar_dependencias()
        print()
        
        # Passo 2: Limpar builds anteriores
        limpar_builds_anteriores()
        print()
        
        # Passo 3: Criar arquivos de configuração
        criar_arquivo_spec()
        criar_version_info()
        print()
        
        # Passo 4: Gerar executável
        if gerar_executavel():
            print()
            
            # Passo 5: Criar arquivos de distribuição
            criar_readme()
            copiar_planilha_template()
            print()
            
            # Verificar resultado
            exe_path = 'dist/Automacao_Sicoob.exe'
            if os.path.exists(exe_path):
                size_mb = os.path.getsize(exe_path) / (1024 * 1024)
                print("🎉 SUCESSO!")
                print(f"📁 Executável: {exe_path}")
                print(f"📏 Tamanho: {size_mb:.1f} MB")
                print()
                print("📦 ARQUIVOS PARA DISTRIBUIR:")
                print("   - dist/Automacao_Sicoob.exe")
                print("   - dist/Automacao_Gransoft.xlsx")  
                print("   - dist/LEIA-ME.txt")
                print()
                print("✅ Pronto para distribuição!")
                
                # Abrir pasta
                resposta = input("Deseja abrir a pasta dist? (s/n): ").lower().strip()
                if resposta in ['s', 'sim', 'y', 'yes']:
                    if os.name == 'nt':  # Windows
                        os.startfile('dist')
            else:
                print("❌ Executável não foi encontrado após geração!")
                
        else:
            print("❌ Falha na geração do executável!")
            
    except Exception as e:
        print(f"❌ Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        print()
        input("Pressione ENTER para sair...")

if __name__ == "__main__":
    main()