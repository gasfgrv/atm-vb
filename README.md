# CAIXA ELETRÔNICO
- Criar um menu de opções:
>  - Menu Principal
>      1. Carregar Notas
>      2. Retirar Notas
>      3. Estatística
>      4. Fim   
- Carregar a quantidade de notas em uma área da memória com 6 ocorrências.
- Solicitar que o cliente faça a retirada de valores obedecendo ao critério do maior pelo menor.
- Dar a opção para o cliente escolher o valor e a quantidade de notas. P. ex.: 1 x 20, 2 x 10.
- Caso não tenha o valor da maior cédula, disponibilizar a próxima.
- Se o valor a ser solicitado for maior que o saldo total do caixa, enviar a mensagem **“EXCEDEU O LIMITE DO CAIXA”**.
- Solicitar até 100 retiradas ou até não haver mais notas.
- No momento da solicitação do valor, coletar também o código do banco que o cliente tem conta, segundo:

> | Código: |        1        |     2    |   3  |   4   |
> |:-------:|:---------------:|:--------:|:----:|:-----:|
> |  Banco: | Banco do Brasil | Santader | Itaú | Caixa |

- No final, exibir a estatística, separada por bancos, com:
  - O maior e o menor valor sacado;
  - A média dos saques;
  - Valor total dos saques;
  - Valor das sobras dos caixas.