# Criacao-user-AD

Criação e Gerenciamento de Usuários no Active Directory - PowerShell GUI

Este projeto é uma ferramenta em PowerShell com interface gráfica (GUI) para facilitar a criação e gerenciamento de usuários no Active Directory (AD). Utiliza Windows Forms para construção da interface e o módulo ActiveDirectory para execução dos comandos.

Funcionalidades

1. Criação de Usuário

Formulário com campos:

Nome Completo

Login

Senha

E-mail

Departamento

Cargo

Sub-OU (unidade organizacional complementar)

OU principal

Grupo(s) (opcional)

Sufixo UPN

Validação de dados obrigatórios

Criação do usuário com senha, UPN, OU e grupo(s)

2. Gerenciamento de Usuário

Verificação de status (habilitado/desabilitado)

Habilitar ou desabilitar um usuário do AD

Requisitos

Windows PowerShell 5.1+

Permissão de administrador (ou execução com ExecutionPolicy Bypass)

Módulo ActiveDirectory instalado (RSAT: Active Directory Tools)

Como Usar

Clone este repositório:

Execute o script principal no PowerShell:

nome-do-script.ps1

Escolha entre:

Abrir a interface de criação de usuário

Abrir a interface de gerenciamento de usuários

Importante: Altere os valores de exemplo do script para refletir o seu ambiente de Active Directory:

Sufixos UPN (ex: @empresa.com.br)

OUs (ex: OU=Usuarios,DC=empresa,DC=com,DC=br)







