# SisNCA Completo

Este projeto reúne três módulos principais, cada um com funcionalidades específicas para o gerenciamento de atendimentos, integração com cidadãos e automação de e-mails via Google Apps Script.

## Estrutura do Projeto

- **Atendimento/**: Responsável pelo gerenciamento de atendimentos, incluindo interface e lógica de controle.
  - `src/appsscript.json`: Configurações do projeto Apps Script.
  - `src/Code.js`: Código principal do script de atendimento.
  - `src/painel.html`: Interface do painel de atendimento.

- **Cidadao/**: Voltado para a interação com cidadãos, consultas e funcionalidades relacionadas.
  - `src/appsscript.json`: Configurações do projeto Apps Script.
  - `src/Código.js`: Código principal do módulo cidadão.
  - `src/cidadao.html`: Interface do cidadão.
  - `src/consulta.html`: Interface de consulta.
  - `VERSOES.md`: Histórico de versões do módulo cidadão.

- **Email/**: Automatiza o envio de e-mails e integra com planilhas Google via Apps Script.
  - `src/appsscript.json`: Configurações do projeto Apps Script.
  - `src/Código.js`: Código principal do script de envio de e-mails.

Cada módulo está em uma pasta separada e pode ser desenvolvido de forma independente.

## Como clonar e configurar o projeto

1. Clone o repositório principal:
   ```powershell
   git clone https://github.com/ncapgepa/sisnca.git
   ```

2. Não há mais submódulos. Todo o código está neste repositório.

## Manual de Instruções

### 1. Estrutura de Pastas
- Os arquivos principais de cada módulo estão na pasta `src/` de cada subdiretório.

### 2. Email (Google Apps Script)
- Para editar ou publicar o módulo Email, utilize o [clasp](https://github.com/google/clasp) para sincronizar com o Google Apps Script.
- Exemplo de clonagem:
  ```powershell
  clasp clone <scriptId> --rootDir Email/src
  ```
  Substitua `<scriptId>` pelo ID do script correspondente.

#### Permissões necessárias para o Email
O projeto utiliza as seguintes permissões:
- `https://www.googleapis.com/auth/spreadsheets`
- `https://www.googleapis.com/auth/script.send_mail`
- `https://www.googleapis.com/auth/script.container.ui`

- O projeto está configurado para rodar na timezone America/Sao_Paulo.
- O acesso ao webapp está liberado para qualquer usuário anônimo.

### 3. Informações sobre a Planilha Google e a Pasta no Drive

#### Estrutura da Planilha Google
A planilha utilizada pelo sistema deve conter uma aba chamada **Pedidos Prescrição** com as seguintes colunas:

| Coluna           | Descrição                                                                                                    | Exemplo de Conteúdo                                      |
|------------------|-------------------------------------------------------------------------------------------------------------|----------------------------------------------------------|
| Protocolo        | Gerado automaticamente pelo sistema. Único e não editável.                                                  | PGE-PRESC-2024-0001                                      |
| Timestamp        | Data e hora do envio do formulário. Preenchido automaticamente.                                             | 23/06/2024 14:30:15                                      |
| NomeSolicitante  | Nome do Titular ou Representante Legal.                                                                     | José da Silva                                            |
| Email            | E-mail de contato.                                                                                          | jose.silva@email.com                                     |
| Telefone         | Telefone de contato com DDD.                                                                                | (91) 99999-8888                                          |
| TipoPessoa       | Tipo de pessoa (Pessoa Física, Empresário, Sócio, Procurador).                                              | Pessoa Física                                            |
| CDAs             | Números das CDAs, separados por vírgula.                                                                    | 12345, 67890, 11223                                      |
| LinkDocumentos   | Link para a pasta no Google Drive com os documentos do solicitante.                                         | https://drive.google.com/drive/folders/123xyz...         |
| Status           | Status atual do pedido. Controlado pelo atendente.                                                          | Novo, Em Análise, Pendente, Deferido, Indeferido         |
| AtendenteResp    | Nome do atendente que está com o caso.                                                                      | Maria Souza                                              |
| Historico        | Registros de cada mudança de status e observações internas.                                                 | 24/06: Análise inicial. 25/06: Documentação pendente.    |
| DataEncerramento | Data em que o status foi mudado para Deferido/Indeferido.                                                   | 30/06/2024                                               |

Cada linha representa um pedido único realizado pelo Portal do Cidadão. Os campos são utilizados tanto para acompanhamento pelo solicitante quanto para gestão interna pelo atendente.

#### Pasta no Google Drive
- Para cada solicitação, é criada uma pasta no Google Drive para armazenar os documentos enviados pelo solicitante.
- O link para essa pasta deve ser registrado na coluna **LinkDocumentos** da planilha.
- Recomenda-se organizar as pastas por protocolo ou nome do solicitante para facilitar a localização e auditoria.

### 4. Recomendações Gerais
- Consulte este README para instruções detalhadas de uso e configuração.
- Para dúvidas ou problemas, entre em contato com o responsável pelo projeto.

---

Este manual serve como guia rápido para instalação, configuração e uso dos módulos do SisNCA Completo.
