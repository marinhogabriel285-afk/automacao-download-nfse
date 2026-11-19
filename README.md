# 🤖 Automação de Download e Relatório Financeiro de NFS-e

Script em Python desenvolvido para automatizar o download de Notas Fiscais de Serviço Eletrônicas (NFS-e) recebidas diretamente da API oficial do **Portal Nacional (`adn.nfse.gov.br`)**, realizando o cálculo automático de retenções federais e consolidando os dados em uma planilha Excel pronta para fechamento contábil.

---

## 🚀 Funcionalidades

* **Autenticação Segura (mTLS):** Conexão direta com os servidores do governo utilizando Certificado Digital A1 (`.pfx`) localmente, sem depender de plataformas de terceiros.
* **Filtro Inteligente por Competência:** Busca por mês/ano e separa automaticamente apenas as notas **recebidas** (tomadas).
* **Cálculo Tributário Automatizado:** Lê as tags oficiais do XML (`vServ`, `vRetCP`, `vRetPIS`, `vRetCOFINS`, `vRetCSLL`) e calcula o valor líquido real descontando retenções previdenciárias e contribuições sociais.
* **Relatório Consolidado em Excel (`.xlsx`):** Gera uma planilha formatada com totais automáticos via fórmulas nativas do Excel.
* **Gerador Local de DANFSe (PDF):** Opção para gerar o documento auxiliar em PDF a partir do XML assinado.

---

## 🛠️ Tecnologias Utilizadas

* **Python 3**
* **Requests & Requests-PKCS12:** Comunicação HTTP/HTTPS com autenticação mútua TLS 1.2.
* **XML ElementTree:** Parsing e extração de dados fiscais estruturados.
* **OpenPyXL:** Geração e estilização de planilhas Excel.
* **ReportLab:** Geração dinâmica de documentos PDF.

---

## 📦 Como Executar o Projeto

1. **Instale as dependências:**
   ```bash
   pip install requests requests-pkcs12 reportlab openpyxl

2. **Coloque o seu certificado.pfx**
na mesma pasta do script e preencha as variáveis no topo do arquivo baixar_nfse.py:
CERTIFICADO_PFX = "seu_certificado.pfx"
SENHA_CERTIFICADO = "sua_senha"
MEU_CNPJ = "00000000000000"

3. **Executar**
Na barra de pesquisa da pasta que estao o certificao e o arquivo baixar_nfse.py, abra o cmd e execute:
python baixar_nfse.py
<img width="515" height="123" alt="image" src="https://github.com/user-attachments/assets/9c62bac6-885a-4e3b-bcfe-16de1257c142" />
<img width="701" height="133" alt="image" src="https://github.com/user-attachments/assets/ac7b5257-e407-42af-8299-07b210c3e7d0" />
<img width="616" height="95" alt="image" src="https://github.com/user-attachments/assets/3f9f4ae1-dc89-4afb-994b-4899a9456bae" />



