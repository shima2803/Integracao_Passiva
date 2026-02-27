# Automação Alpheratz – Integração CPJWCS (Evento itp45)

Este projeto consiste em um robô desenvolvido em Python para automatizar o tratamento diário do evento itp45, integrando informações do banco CPJWCS com o sistema Alpheratz.

O objetivo é reduzir atividades manuais, aumentar a padronização e diminuir riscos de erro operacional.

## O que o sistema faz

Consulta automaticamente o banco CPJWCS e identifica os registros do evento itp45 do dia.

Acessa o Alpheratz via automação.

Pesquisa o contrato pelo número de integração.

Abre o workflow correspondente.

Valida a etapa correta antes de avançar.

Preenche automaticamente os campos da tela com base no texto registrado no evento.

Salva as informações no sistema.

## Benefícios

Redução de retrabalho manual.

Maior agilidade operacional.

Padronização no preenchimento.

Diminuição de erros humanos.

Automatização de rotina repetitiva.

## Requisitos

Python 3.x

Selenium

PyMySQL

Navegador Google Chrome compatível

Acesso ao banco CPJWCS

Acesso ao sistema Alpheratz
---
## Observação de Segurança

As URLs, usuários e senhas presentes no código original foram alterados ou substituídos por valores fictícios nesta versão publicada, por motivos de segurança e confidencialidade.
