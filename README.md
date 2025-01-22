# nomesEmListas
Código desenvolvido para o **Cursinho FEAUSP**! 

Jeito simples de verificar aprovações nos vestibulares de universidades públicas! Carregue uma planilha do Excel com os nomes de seus alunos e os caminhos das listas no seu computador, aí é só rodar e pronto. 

Sua planilha de nomes deve ser configurada da seguinte forma:
  - Coluna A "Nome": nomes completos dos alunos
  - Coluna B "CPF": CPFs dos alunos
  - Coluna C "CPF ajustado": CPFs formatados de acordo com a Fuvest (XXX.XXX), pode ser feito com a seguinte fórmula no Excel
      =ESQUERDA(A1; 3)&"."&MID(A1; 4; 3)
Além disso, cada aba deve corresponder a uma turma ou a qualquer identificador interessante.

É importante pontuar que o programa somente procura nomes em listas e retorna onde esse nome está presente, então é bom checar o resultado final com cuidado já que alguns alunos podem ter mais de uma aprovação. Use a fórmula ÚNICO no Excel para verificar se há as duplicatas e CONT.SE para checar o número de aprovações por aluno.
