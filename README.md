[![OS - Windows](https://img.shields.io/badge/OS-Windows-blue?logo=windows&logoColor=white)](https://www.microsoft.com/ "Go to Microsoft homepage")  [![GustavoBarbaro - Excel_Scraping_FII_Strategy_Invest](https://img.shields.io/static/v1?label=GustavoBarbaro&message=Excel_Scraping_FII_Strategy_Invest&color=gree&logo=github)](https://github.com/GustavoBarbaro/Excel_Scraping_FII_Strategy_Invest "Go to GitHub repo") [![License](https://img.shields.io/badge/License-MIT-yellow)](#license)

---

**Esse projeto não deve ser considerado como uma recomendação de investimento !!!**



---

# Sobre o projeto

Esse projeto foi desenvolvido com a intenção de estudar e praticar web-scraping no Excel.

Investindo todos os meses em fundos imobiliários, senti a necessidade de ter uma ferramenta que compilasse todos os fundos listados na bolsa e aplicasse alguns filtros para que eu pudesse escolher em qual iria aportar no mês.

Então decidi criar a ferramenta eu mesmo, utilizando como base o Excel e o Selenium para o scraping. A ferramenta copia os fundos listados para uma planilha e aplica os filtros automaticamente, além de reunir também quais estão no processo de emissão, ou seja emitindo novas cotas.

Apenas reforçando, inicialmente eu desenvolvi essa ferramente para meu uso prórpiro, mas achei interessante compartilhar para outros que gostariam de analisá-la e quem sabe, melhorá-la.

Não recomendo que outras pessoas a utilizem para investir as cegas sem antes estudar as estratégias e ver que elas fazem sentido para você.



---


# Requisitos

* ![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)

* ![Selenium](https://img.shields.io/badge/-selenium-%43B02A?style=for-the-badge&logo=selenium&logoColor=white)

* ![Google Chrome](https://img.shields.io/badge/Google%20Chrome-4285F4?style=for-the-badge&logo=GoogleChrome&logoColor=white)



---


# Instalação

## Download do Selenium

Para que o projeto funcione é necessário o download do Selenium para uso no Visual Basic Analysis. Caso contrário a biblioteca não aparecerá dentro do Excel.

O donwload se encontra nesse [repositório](https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0).

Após isso basta executar o instalador normalmente.

## Download do Web Driver para Chrome

Faça download do web driver do navegador Chrome, para que o selenium possa realizar o scraping corretamente.

[Download do Web driver para Chrome](https://sites.google.com/chromium.org/driver/)

É muito importante baixar o web driver da mesma versão do Chrome instalada. Recomendo atualizar seu navegador para a versão mais recente e baixar a última versão do web driver.

Para atualizar o Chrome ou ver que versão está instalada: clique nos três pontos verticais no canto superior direito no Chrome > Configurações > Sobre o Google Chrome.

### Atualizando o Web driver do Selenium

Após o download do web driver, extraia o arquivo .zip e copie o executável. Após isso, vá até o caminho onde o selenium foi instalado, geralmente é:

```
C:\Users\SEU_USUARIO\AppData\Local\SeleniumBasic
```

e substitua o executável pelo que foi copiado.



---


# Habilitando a biblioteca selenium dentro do Excel

## Habilitando a guia Desenvolvedor

Abra o Exel, vá em Arquivo > Opções > Personalizar Faixa de Opções

Habilite a caixa Desenvolvedor

## Habilitando a biblioteca

Vá na guia Desenvolvedor (ficará logo depois da guia exibição) e clique em Visual Basic.

Na janela que se abriu, clique em Ferramentas > Referências.

Marque a caixa **Selenium Type Library**



---


# Usabilidade

A figura abaixo representa a página inicial.

<img src="https://user-images.githubusercontent.com/48565991/184003134-7e4cc53a-849b-4b20-a50c-447b5ce10c7b.png" width="600" height="300" />

Nela o usuário poderá entrar com a quantidade mínima de patrimônimo líquido que ele deseja filtrar os fundos.

* Ao clicar no botão Filtra Fundos, o sistema realizará o scraping automaticamente, filtrará os fundos, e atualizará as guias: Top 15 e Base de Dados;

* Os três botões abaixo permitem se deslocar entre as guias;



---


# A Estratégia


A estratégia de recomendação de FIIS consiste em ordenar os fundos que possuem os maiores dividendos e que estão mais baratos. Para isso foi-se criado um sistema de pontuação baseado em dois indicadores: o *dividend yield* e o P/VPA. O sistema de pontuação funcionada da seguinte forma:

1. Primeiro ordenamos a coluna do *dividend yield* do maior para o menor valor, pois estamos querendo o fundo que mais pagou dividendos. Com a coluna ordenada, é estabelecido um ranking dos melhores colocados, começando em 1 para o fundo que mais pagou dividendos.

2. Em seguida ordenamos a coluna P/VPA do menor para o maior, pois estamos a procura dos fundos que estão mais descontados. Logo após, é atribuindo um ranking dos melhores colocados, começando em 1 para o fundo que está mais barato.

3. Por último, para cada FII é somado o seu valor no ranking do *dividend yield* e do P/VPA formando um terceiro ranking. Esse novo ranking é ordenado de forma crescente, pois os FIIS que tiveram o menor valor são os com os melhores rendimentos e mais descontados.




---


