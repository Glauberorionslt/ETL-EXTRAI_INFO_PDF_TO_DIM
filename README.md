Projeto de Extração e Análise de Pedidos de Compra
Este projeto, desenvolvido em Python, tem como objetivo extrair informações de arquivos PDF referentes a pedidos de compra. As extrações são divididas em tabelas de dimensão e fatos, que são posteriormente analisadas para fornecer uma visão detalhada sobre a contabilização de produtos alimentícios, bem como o controle de seus preços e impostos.

Funcionalidades
Extração de Dados de PDF: Utiliza a biblioteca pdfplumber para extrair automaticamente informações detalhadas dos pedidos de compra armazenados em arquivos PDF.
Estruturação dos Dados: Os dados extraídos são organizados em tabelas de dimensão e fatos, facilitando a análise e visualização.
Análise de Dados: A análise dos dados extraídos permite contabilizar produtos alimentícios e controlar seus preços e impostos.
Automação de Monitoramento: Implementa um sistema de monitoramento automático para detectar e processar novos arquivos PDF que são adicionados a uma pasta dedicada.
Bibliotecas Utilizadas
pdfplumber: Utilizada para a extração de dados de arquivos PDF.
os: Utilizada para operações relacionadas ao sistema de arquivos, como monitoramento de novas adições de arquivos.
pandas: Utilizada para manipulação e análise de dados, permitindo a criação de tabelas de dimensão e fatos.
Como Funciona
Extração de Dados:

Os arquivos PDF de pedidos de compra são processados pelo script que utiliza pdfplumber para extrair informações relevantes.
Estruturação dos Dados:

Os dados extraídos são organizados em tabelas de dimensão e fatos utilizando pandas, o que facilita a análise subsequente.
Análise de Dados:

As tabelas de dimensão e fatos são analisadas para obter uma visão detalhada sobre a contabilização de produtos alimentícios, controle de preços e impostos.
Automação e Monitoramento:

O projeto inclui um componente de automação que monitora uma pasta específica para novos arquivos PDF. Quando um novo arquivo é detectado, ele é automaticamente processado e integrado ao sistema de análise.
Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias, correções de bugs ou novas funcionalidades.
