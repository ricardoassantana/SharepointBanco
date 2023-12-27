### SharepointBanco
## Desenvolvimento de Solução para Conexão Automática entre SharePoint e Banco de Dados

![Solução Sharepoint para Banco](https://github.com/ricardoassantana/SharepointBanco/blob/main/Imagem_conexao.png)


Ao enfrentar o desafio de conectar o SharePoint a um banco de dados de forma automática, deparei-me com uma barreira significativa: a falta de permissão para acessar a API do SharePoint diretamente, devido às restrições impostas pela TI. Esta limitação tornou-se um obstáculo crítico, uma vez que várias planilhas em Excel precisavam ser integradas ao banco para dar continuidade a um projeto essencial.

Diante desses desafios, surgiu a ideia de conectar as planilhas Excel localmente armazenadas ao SharePoint. Inicialmente, criei uma série de arquivos locais vazios no formato Excel. Em seguida, utilizei o Power Query para estabelecer uma conexão direta com os arquivos hospedados no SharePoint (Excel Online). Esse procedimento resultou em uma integração entre os arquivos locais e as versões online do Excel. 

O próximo passo seria encontrar uma solução que permitisse automatizar o processo de atualização das planilhas antes de integrá-las ao banco de dados.

A solução veio através do desenvolvimento de um script em Python, que possibilitou na atualização das planilhas Excel com os dados do SharePoint. Essa abordagem permitiu que eu atualizasse as informações de maneira automática, eliminando a necessidade de intervenção manual ou com apenas um clique!
