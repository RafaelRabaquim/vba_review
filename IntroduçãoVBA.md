# Introdução ao VBA no Office

Você está enfrentando uma limpeza repetitiva de cinquenta tabelas no Word? 
Você deseja que um documento específico solicite que o usuário forneça informações quando ele for aberto? 
Você está tendo dificuldade para saber como colocar seus contatos do Microsoft Outlook em uma planilha do Microsoft Excel de maneira eficiente?

Você pode executar essas tarefas e fechar bons negócios usando o **Visual Basic for Applications (VBA)** para Office − *uma linguagem de programação simples*, mas poderosa, que você pode usar para estender aplicativos do Office.

Este artigo explora algumas das principais razões para aproveitar a capacidade de programação em VBA. Ele explora a linguagem VBA e as ferramentas prontas que podem ser usadas para trabalhar com suas soluções. Por fim, ele inclui algumas dicas e formas de evitar algumas frustrações e equívocos comuns que acontecem na programação.
 Observação

Interessado em desenvolver soluções que ampliem a experiência do Office em várias plataformas? Confira o novo modelo de Suplementos do Office. Os suplementos do Office têm um pequeno espaço comparado aos suplementos e soluções do VSTO, e você pode criá-los usando quase qualquer tecnologia de programação Web, como HTML5, JavaScript, CSS3 e XML.

## Quando e por que usar o VBA?

> Há três razões principais para considerar a programação em VBA no Office.

### Automação e repetição

O VBA é eficiente quando se trata de soluções repetitivas para problemas de formatação ou correção. Por exemplo, você já alterou o estilo do parágrafo no topo de cada página no Word? Você já precisou reformatar várias tabelas que foram coladas do Excel em um documento do Word ou em um email do Outlook? Você já precisou fazer a mesma alteração em vários contatos do Outlook?
Se existe uma alteração que precisa ser feita mais de dez ou vinte vezes, talvez valha a pena automatizá-la com o VBA. Se é uma alteração que precisa ser feita centenas de vezes, merece sem dúvida ser considerada. Quase todas as alterações de formatação ou de edição que podem ser feitas manualmente podem ser feitas no VBA.

### Extensões para a interação do usuário

Há ocasiões em que você deseja incentivar ou obrigar os usuários a interagir com um documento ou aplicativo do Office de determinada forma que não faça parte do aplicativo padrão. Por exemplo, você pode pedir aos usuários para executar uma ação específica ao abrir, salvar ou imprimir um documento.

### Interação entre aplicativos do Office

Você precisa copiar todos os seus contatos do Outlook para o Word e, em seguida, formatá-los de alguma maneira específica? Ou precisa mover dados do Excel para um conjunto de slides do PowerPoint? Às vezes, o simples ato de copiar e colar não produz o resultado desejado ou é muito lento. Use a programação em VBA para interagir com os detalhes de dois ou mais aplicativos do Office ao mesmo tempo e modificar o conteúdo de um aplicativo com base no conteúdo contido em outro aplicativo.


>Fazendo as coisas de outra forma

A programação em VBA é uma solução poderosa, mas nem sempre é a melhor abordagem. Às vezes, faz sentido usar outros recursos para atingir seus objetivos.
A principal pergunta a ser feita é se há uma maneira mais fácil. Antes de começar um projeto do VBA, considere as ferramentas internas e as funcionalidades padrão. Por exemplo, se você tem em mãos uma tarefa demorada de layout ou edição, considere o uso de estilos ou teclas aceleradoras para solucionar o problema. É possível executar a tarefa uma vez e depois usar Ctrl+Y (Refazer) para repeti-la? É possível criar um novo documento com o formato ou modelo correto e, em seguida, copiar o conteúdo para esse novo documento?
Os aplicativos do Office são poderosos. A solução de que você precisa talvez já esteja disponível. Dedique algum tempo para saber mais sobre o Office antes de começar a programar.
Antes de começar um projeto VBA, verifique se você tem tempo para trabalhar com o VBA. Programar requer foco e pode ser imprevisível. Especialmente se for iniciante, nunca comece a programar a menos que você tenha tempo para trabalhar com toda a atenção. Tentar gravar um "script rápido" para resolver um problema quando uma data limite se aproxima pode levar a situações muito estressantes. Se você está com pressa, use métodos convencionais, mesmo que sejam monótonos e repetitivos.

## Programação em VBA

> Usando códigos para fazer com que aplicativos realizem tarefas

Você pode achar que escrever códigos é algo misterioso ou difícil, mas os princípios básicos usam raciocínios comuns e são bastante acessíveis. Os aplicativos do Microsoft Office são criados de forma a expor itens chamados objetos. Esses itens podem receber instruções da mesma forma que um telefone faz quando você usa seus botões. Quando você pressiona um botão, o telefone reconhece a instrução e inclui o número correspondente na sequência que você está discando. Em programação, você interage com o aplicativo enviando instruções a vários objetos no aplicativo. Esses objetos são expansivos, mas têm limites. Eles só podem fazer aquilo que foram projetados para fazer e farão apenas o que você os instruir a fazer.

Por exemplo, pense em um usuário que abre um documento no Word, faz algumas alterações, salva o documento e, em seguida, fecha-o. No mundo da programação em VBA, o Word expõe um objeto Document. Usando o código VBA, você pode instruir o objeto Document a fazer coisas como abrir, salvar ou fechar.

A seção a seguir discute como os objetos são organizados e descritos.

### O modelo de objeto

Desenvolvedores organizam os objetos de programação em uma hierarquia chamada de modelo de objeto do aplicativo. O Word, por exemplo, tem um objeto Application de nível superior que contém um objeto Document. O objeto Document contém objetos Paragraph e assim por diante. Os modelos de objeto espelham praticamente aquilo que você vê na interface do usuário. Eles são um mapa conceitual do aplicativo e seus recursos.

A definição de um objeto é chamada de classe, portanto, você pode ver esses dois termos. Tecnicamente, uma classe é a descrição ou o modelo usado para criar, ou instanciar, um objeto.
Depois que um objeto existe, é possível manipulá-lo definindo suas propriedades e chamando seus métodos. Se você pensar no objeto como um substantivo, as propriedades são os adjetivos que descrevem o substantivo e os métodos são os verbos que dão vida ao substantivo. Ao alterar uma propriedade, você altera algumas qualidades da aparência ou do comportamento do objeto. Chamar um dos métodos do objeto faz com que o objeto execute alguma ação.

O código VBA neste artigo é executado em um aplicativo aberto do Office em que muitos dos objetos que o código manipula já se encontram em execução. Por exemplo, os objetos Application, Worksheet no Excel, Document no Word, Presentation no PowerPoint, Explorer e Folder no Outlook. Após conhecer o layout básico do modelo de objeto e algumas propriedades importantes de Application que dão acesso ao seu estado atual, é possível começar a estender e manipular esse aplicativo do Office com o VBA no Office.

### Métodos

No Word, por exemplo, você pode alterar as propriedades e chamar os métodos do documento atual do Word, usando a propriedade ActiveDocument do objeto aplicativo. Essa propriedade ActiveDocument retorna uma referência para o objeto documento que está ativo no aplicativo do Word. "Retorna uma referência para" significa "oferece acesso a".
O código a seguir faz exatamente aquilo que ele diz, ou seja, salva o documento ativo no aplicativo.

[Continue lendo...](https://docs.microsoft.com/pt-br/office/vba/library-reference/concepts/getting-started-with-vba-in-office#methods)
