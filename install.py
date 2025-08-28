#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de instalação automática das dependências para Automação Sicoob
Execute este arquivo antes de usar o programa principal.
"""

import subprocess
import sys
import importlib

def instalar_biblioteca(nome):
    """Instala uma biblioteca usando pip"""
    try:
        print(f"📦 Instalando {nome}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", nome])
        print(f"✅ {nome} instalado com sucesso!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao instalar {nome}: {e}")
        return False

def verificar_biblioteca(nome):
    """Verifica se uma biblioteca está instalada"""
    try:
        importlib.import_module(nome)
        print(f"✅ {nome} já está instalado")
        return True
    except ImportError:
        print(f"⚠️ {nome} não está instalado")
        return False

def main():
    """Função principal do instalador"""
    print("🚀 Instalador de Dependências - Automação Sicoob")
    print("=" * 50)
    
    bibliotecas_necessarias = {
        'openpyxl': 'openpyxl',
        'chardet': 'chardet',
        'pandas': 'pandas'
    }
    
    bibliotecas_opcionais = {
        'xlrd': 'xlrd'  # Para arquivos Excel mais antigos
    }
    
    print("🔍 Verificando bibliotecas necessárias...")
    
    # Verifica bibliotecas necessárias
    bibliotecas_para_instalar = []
    for nome_importacao, nome_pip in bibliotecas_necessarias.items():
        if not verificar_biblioteca(nome_importacao):
            bibliotecas_para_instalar.append(nome_pip)
    
    if bibliotecas_para_instalar:
        print(f"\n📋 Será necessário instalar: {', '.join(bibliotecas_para_instalar)}")
        
        resposta = input("\n❓ Deseja instalar as bibliotecas automaticamente? (s/n): ").lower().strip()
        
        if resposta in ['s', 'sim', 'y', 'yes']:
            print("\n🔄 Iniciando instalação...")
            
            sucessos = 0
            for biblioteca in bibliotecas_para_instalar:
                if instalar_biblioteca(biblioteca):
                    sucessos += 1
            
            print(f"\n📊 Resultado: {sucessos}/{len(bibliotecas_para_instalar)} bibliotecas instaladas com sucesso")
            
            if sucessos == len(bibliotecas_para_instalar):
                print("\n🎉 Todas as dependências foram instaladas!")
                print("✅ Você pode executar o programa principal agora.")
            else:
                print("\n⚠️ Algumas bibliotecas não foram instaladas.")
                print("💡 Tente executar manualmente: pip install " + " ".join(bibliotecas_para_instalar))
        else:
            print("\n📝 Para instalar manualmente, execute:")
            print(f"pip install {' '.join(bibliotecas_para_instalar)}")
    
    else:
        print("\n🎉 Todas as bibliotecas necessárias já estão instaladas!")
        print("✅ Você pode executar o programa principal.")
    
    # Verifica bibliotecas opcionais
    print("\n🔍 Verificando bibliotecas opcionais...")
    for nome_importacao, nome_pip in bibliotecas_opcionais.items():
        if not verificar_biblioteca(nome_importacao):
            print(f"💡 Biblioteca opcional '{nome_pip}' não instalada (não é obrigatória)")
    
    print("\n" + "=" * 50)
    print("✅ Verificação concluída!")
    
    input("\nPressione ENTER para sair...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n❌ Instalação cancelada pelo usuário.")
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")
        input("Pressione ENTER para sair...")