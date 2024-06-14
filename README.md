# Louvor-Shalom

Modelo de PO que utiliza MILP para alocação de musicos na escala dos domingos na IP - Shalom, Campinas-SP.

Para utilização deste modelo há a necessidade do python instalado e das seguintes bibliotecas:

- pandas
- pulp

Abaixo temos a descrição do problema:

Um grupo musical de uma igreja precisa tocar dodos os domingos do ano, com raras exceções. Com isso, precisa-se alocar pessoas em posições de acordo com suas capacidades, para cadda dia do ano, sabendo que nem todos restão disponíveis em todos os dias e o número de dias que cada um irá participar por mês é indicado por cada participante.

Vale lembrar, também, que além de músicos, há um vocal que canta junto, ou seja, existe a necessidade de alocar pessoas em instrumentos e vozes para cada dia da semana, sem sobrecarregar ninguém e sem que ninguém toque mais que um instrumento por vez.

Apesar da explicação ser breve, há um problema. Uma pessoa que domina um instrumento e uma voz pode ser alocada em ambas as posições simultaneamente, ou seja, cantará e tocará ao mesmo tempo. Isto posto foi desenvolvido dois modelos:

- V1 - Versão simplificada. Aloca musicos em posições sem a participação de um músico em mais de uma posição ao mesmo tempo. Problema tridimensional. ✔
- V2 - Versão mais complexa. Aloca uma pessoa em um instrumento e/ou uma voz. Problema Quadridimensional. ✔
