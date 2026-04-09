# Tutorial de Configuração por Unidade

Este tutorial orienta cada escola a ajustar o sistema para que os dados fiquem iguais aos campos do agendamento da sua própria unidade no Microsoft Bookings.

Use as imagens enviadas como referência visual durante os passos abaixo.

## Objetivo

Cada unidade deve:

- informar o link correto do seu Bookings
- fazer login com o usuário da unidade
- ajustar as listas de opções dos campos para refletir os dados reais da escola

## Antes de começar

Tenha em mãos:

- o link oficial do Bookings da unidade
- acesso ao usuário que fará o login no sistema
- os nomes exatos usados no formulário da unidade, como:
  - equipe
  - público
  - componente
  - principal recurso
  - tipo de atividade
  - turmas
  - horários

Importante:

- os nomes devem ser cadastrados exatamente como aparecem no agendamento da unidade
- se um nome estiver diferente, a automação pode não encontrar a opção correta

## Imagem 1: Tela principal

Na tela principal, localize o campo **Link do Bookings da unidade**.

Passos:

1. Cole o link oficial do Bookings da sua escola.
2. Clique em **Salvar link**.
3. Clique em **Fazer Login / Trocar usuário**.
4. Faça o login com a conta da unidade.
5. Depois do login, volte ao programa e confirme a janela de login quando solicitado.

Recomendação:

- sempre confirme se o link salvo pertence à unidade correta antes de continuar

## Imagem 2: Modo Desenvolvedor

Depois de salvar o link e fazer o login, clique em **Modo Desenvolvedor**.

Nessa tela, a opção que cada escola vai usar para personalizar os dados é:

- **Editar opções de campos**

Os demais botões são técnicos e não precisam ser usados no dia a dia pelas unidades, salvo orientação específica.

## Imagem 3: Editor de opções de campos

Ao abrir o editor, escolha no campo superior qual lista deseja alterar.

Campos disponíveis para edição:

- `escolha_reserva`
- `componente`
- `publico`
- `turma`
- `principal_recurso`
- `tipo_atividade`
- `equipe`
- `horario`

Como preencher:

- digite os valores separados por vírgula
- mantenha a escrita exatamente igual à usada no Bookings

Exemplo de `escolha_reserva`:

```text
50min, 1h40min
```

Depois de editar:

1. Clique em **Salvar valores**.
2. Repita o processo para os outros campos que precisarem ser adaptados.

Se quiser desfazer uma edição que ainda não foi salva:

- clique em **Recarregar campo**

## Imagem 4: Exemplo do campo equipe

No campo `equipe`, cada unidade deve cadastrar os espaços, ambientes ou responsáveis exatamente como aparecem no agendamento da escola.

Exemplo:

```text
FABLAB CE109, Ateliê de Arte CE109, Auditório I CE109
```

Orientações para esse campo:

- use os nomes completos
- se houver acentos, mantenha os acentos
- se houver siglas ou códigos da unidade, mantenha exatamente como aparecem no Bookings

## Como descobrir os valores corretos da unidade

Abra o Bookings da sua escola e confira os nomes exibidos nos campos do formulário.

Depois, copie esses nomes para o editor do sistema.

Os campos mais importantes para revisar são:

- `equipe`
- `publico`
- `componente`
- `principal_recurso`
- `tipo_atividade`
- `turma`
- `horario`

## Ordem recomendada de configuração

1. Salvar o link do Bookings da unidade.
2. Fazer login com o usuário correto.
3. Abrir **Modo Desenvolvedor**.
4. Abrir **Editar opções de campos**.
5. Ajustar os campos da unidade.
6. Salvar os valores.
7. Fechar e abrir o app novamente.
8. Fazer um **Teste de reservar** antes de usar **Reservar**.

## Atenção importante

- após salvar novos valores no editor, reinicie o app para aplicar as alterações
- use primeiro o botão **Teste de reservar**
- só use **Reservar** quando tudo estiver conferido

## Checklist para cada escola

Antes de usar o sistema, confirme:

- link do Bookings salvo corretamente
- usuário da unidade logado
- campo `equipe` atualizado
- campo `publico` atualizado
- campo `componente` atualizado
- campo `principal_recurso` atualizado
- campo `tipo_atividade` atualizado
- turmas e horários conferidos
- teste realizado com sucesso

## Resumo rápido

Cada escola precisa apenas fazer 3 coisas:

1. Informar o link do seu Bookings.
2. Fazer login com a conta correta.
3. Ajustar os campos no editor para ficar igual ao formulário da unidade.

Se os nomes cadastrados no sistema forem iguais aos nomes do Bookings da escola, a automação funcionará corretamente.
