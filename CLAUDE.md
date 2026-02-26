# Claude Configuration for Vibecoding Project

Olá! Meu nome é Saul Gonçalves e você pode me chamar de "Grande Mestre" (só para nos divertirmos!). Sou um grande fã de Python e adoro usar emojis no meu código e nas minhas comunicações.

Este arquivo `CLAUDE.md` serve para te dar contexto sobre minhas preferências e práticas de codificação para o projeto `vibecoding`, que é meu projeto principal de vibecoding no GitHub. Quero garantir que você, Claude, tenha o entendimento necessário para me auxiliar da melhor forma possível.

Idioma padrão de comunicação deve ser o Português-Brasil 🇧🇷

Sempre estou trabalhando com ambiente virtual e com o sistema operacional windows.
Ao instalar as bibliotecas, ative o ambiente virtual.

# Instruções para Desenvolvimento de Código

## Princípios Fundamentais

Ao desenvolver código, você deve seguir rigorosamente os seguintes princípios:

### SOLID

**S - Single Responsibility Principle (Princípio da Responsabilidade Única)**
- Cada classe/função deve ter apenas uma responsabilidade
- Uma razão para mudar = uma responsabilidade
- Evite classes "faz-tudo" (God Classes)

**O - Open/Closed Principle (Princípio Aberto/Fechado)**
- Aberto para extensão, fechado para modificação
- Use abstrações, interfaces e herança
- Evite modificar código existente que já funciona

**L - Liskov Substitution Principle (Princípio da Substituição de Liskov)**
- Objetos derivados devem ser substituíveis por seus objetos base
- Subclasses devem manter o comportamento esperado da classe pai
- Não quebre contratos estabelecidos

**I - Interface Segregation Principle (Princípio da Segregação de Interface)**
- Interfaces específicas são melhores que interfaces genéricas
- Clientes não devem depender de métodos que não usam
- Prefira múltiplas interfaces pequenas

**D - Dependency Inversion Principle (Princípio da Inversão de Dependência)**
- Dependa de abstrações, não de implementações concretas
- Módulos de alto nível não devem depender de módulos de baixo nível
- Use injeção de dependência

### YAGNI (You Aren't Gonna Need It)

- **Não implemente funcionalidades até que sejam realmente necessárias**
- Evite over-engineering e especulações sobre necessidades futuras
- Mantenha o código simples e focado nos requisitos atuais
- Adicione complexidade apenas quando comprovadamente necessária

### KISS (Keep It Simple, Stupid)

- **Simplicidade é a máxima sofisticação**
- Prefira soluções simples e diretas
- Evite abstrações desnecessárias
- Código claro é melhor que código "inteligente"
- Se existe uma forma mais simples, use-a

## Diretrizes de Implementação

### Estrutura do Código
- Use nomes descritivos para classes, métodos e variáveis
- Mantenha funções pequenas e focadas
- Evite aninhamento excessivo (máximo 3 níveis)
- Prefira composição sobre herança quando apropriado

### Tratamento de Dependências
- Injete dependências via construtor ou parâmetros
- Use interfaces para definir contratos
- Evite dependências circulares
- Minimize o acoplamento entre componentes

### Refatoração Contínua
- Refatore código duplicado
- Simplifique lógica complexa quando possível
- Remove código morto regularmente
- Mantenha testes atualizados durante refatorações

## Exemplo de Aplicação dos Princípios

```typescript
// ❌ Violação dos princípios
class UserManager {
  saveUser(user: User) { /* salva no banco */ }
  sendEmail(user: User) { /* envia email */ }
  validateUser(user: User) { /* valida dados */ }
  generateReport() { /* gera relatório */ }
}

// ✅ Seguindo os princípios
interface UserRepository {
  save(user: User): void;
}

interface EmailService {
  sendWelcomeEmail(user: User): void;
}

interface UserValidator {
  validate(user: User): boolean;
}

class UserService {
  constructor(
    private userRepo: UserRepository,
    private emailService: EmailService,
    private validator: UserValidator
  ) {}

  registerUser(user: User): void {
    if (!this.validator.validate(user)) {
      throw new Error('Invalid user data');
    }
    
    this.userRepo.save(user);
    this.emailService.sendWelcomeEmail(user);
  }
}
```

## Checklist de Revisão

Antes de finalizar qualquer implementação, verifique:

- [ ] Cada classe tem uma única responsabilidade?
- [ ] O código está aberto para extensão, mas fechado para modificação?
- [ ] As abstrações podem ser substituídas sem quebrar o sistema?
- [ ] As interfaces são específicas e coesas?
- [ ] As dependências estão invertidas (depende de abstrações)?
- [ ] Implementei apenas o que é necessário agora?
- [ ] A solução é a mais simples possível que funciona?
- [ ] O código é fácil de entender e manter?

## Lembre-se

> "Premature optimization is the root of all evil" - Donald Knuth

> "Any fool can write code that a computer can understand. Good programmers write code that humans can understand" - Martin Fowler

Foque em escrever código limpo, testável e mantível seguindo estes princípios fundamentais.

## Filosofia de Codificação

Prezo muito por código bem documentado. Isso significa que todo arquivo deve ter uma documentação clara no início, explicando seu propósito. Além disso, cada método ou função dentro do arquivo também deve ser documentado individualmente.

Penso muito na manutenabilidade do código. Quero que meu eu do futuro consiga entender o que escrevi sem precisar de uma inteligência artificial para explicar. Por isso, prefiro ter muitos arquivos pequenos e significativos, organizados em uma subpasta chamada `lib`, em vez de um único arquivo gigante com centenas de funções.

Adoro a sintaxe do Markdown e tento manter meus arquivos compatíveis com esse formato sempre que possível. Também gosto de um toque de cor e emojis para deixar tudo mais divertido! 🎨. Porem NÃO UTILIZAR EM TESTES e saídas do terminal por questões de compatibilidade.

## Ferramentas

Minha caixa de ferramentas padrão inclui:

* **git:** Para controle de versão.
* **glow:** Para visualizar arquivos Markdown no terminal.
* **just:** Como um executor de tarefas conveniente.
* **VSCode:** Meu editor de código preferido.
* **GCP (Google Cloud Platform):** Minha nuvem preferida, já que trabalho no Google Cloud.
* **Claude:** Meu assistente de IA preferido para codificação e desenvolvimento.

## Práticas de Git

Como este código é gerenciado com git, por favor, siga estas diretrizes:

* **Remoção e Movimentação de Arquivos:** Não utilize os comandos `rm` ou `mv` diretamente. Use `git rm` para remover arquivos e `git mv` para renomeá-los ou movê-los.
* **Alterações Perigosas:** Se uma alteração no código for potencialmente arriscada, crie uma branch de feature específica para essa modificação e faça seus commits nessa branch.
* **Branch Principal (main):** Adicione código diretamente à branch `main` somente se a alteração for simples e segura.
* **Changelog:** Certifique-se de que haja um arquivo `CHANGELOG.md` na raiz do projeto. Este arquivo deve conter um histórico de alterações (changelog) e estar vinculado à versão atual do projeto.
* **Gerenciamento de Versões:** Se a versão do projeto for gerenciada de forma explícita (por exemplo, usando `uv` e `project.toml` para projetos Python), siga essa convenção para versionar o código. Utilize a versão semântica.

## Ciclo de Feedback com Claude

Como pretendo invocar sua ajuda, Claude, várias vezes ao longo do desenvolvimento deste projeto, é crucial que você mantenha o contexto em cada interação. Este arquivo `CLAUDE.md` deve te ajudar com isso.

Por exemplo, se em uma interação anterior você me orientou a 'adicionar a função `a` e `b`' e você observar que já existem as funções `a`, `b` e `c`, por favor, não remova a função `c`. Provavelmente existe um motivo, mesmo que não documentado, para termos implementado essa função. Confie no histórico do código, a menos que haja uma instrução explícita para remover algo.

## Estrutura de Subpastas

A organização do projeto em subpastas é importante para mim. Cada pasta dentro do projeto deve seguir estas convenções:

* **README.md:** Cada subpasta deve conter um arquivo `README.md` que explique o propósito e o conteúdo daquela pasta.
* **Estrutura do Projeto:** Exceto na pasta raiz, cada arquivo `README.md` deve conter um capítulo de nível H2 chamado "Estrutura do Projeto". Este capítulo deve apresentar uma visão em árvore da estrutura de pastas, no estilo do comando `tree`. Ao gerar essa árvore, utilize o comando `tree` para obter a estrutura real, mas certifique-se de podar todos os arquivos que forem irrelevantes, ignorados pelo git ou ativos não essenciais. A árvore deve ser concisa o suficiente para fornecer uma boa visão geral da organização do código para outros desenvolvedores.

## Testes

Os testes são uma parte essencial do meu fluxo de trabalho.

* **Rapidez e Significado:** Os testes devem ser rápidos para executar e devem testar funcionalidades significativas do código.
* **Execução:** Como você pode invocar os testes utilizando o comando `just test`, você deve considerar ocasionalmente verificar se suas alterações introduziram alguma quebra de funcionalidade.
* **Cobertura:** Se você identificar que alguma parte do código está quebrada, mas não há testes disponíveis para essa funcionalidade, escreva um teste para cobrir esse caso, a menos que você tenha informações de que essa funcionalidade será removida em breve.

## Cache

Para otimizar o tempo de desenvolvimento, especialmente ao interagir com LLMs ou realizar tarefas demoradas como buscar arquivos grandes, tente implementar um mecanismo de cache o mais cedo possível.

* **Localização:** Prefiro que a pasta de cache (`.cache/`) esteja localizada em um lugar plausível e documentado. Pode ser tanto dentro deste repositório git quanto no diretório home do usuário. A escolha é sua, desde que você documente claramente a localização.
* **Padrão de Expiração:** O cache deve ter um tempo de vida padrão razoável. Se estiver em dúvida, sugiro manter os dados em cache por um dia.
* **Substituição:** O tempo de vida padrão do cache deve ser substituível a cada invocação da tarefa que utiliza o cache, permitindo flexibilidade quando necessário.

## Secrets

Para informações sensíveis, como chaves de API ou tokens, utilizo um arquivo `.env` que não é versionado (ou seja, está no `.gitignore`).

* **.env.dist:** Para fins de documentação, todas as variáveis de ambiente necessárias para o projeto devem ser listadas em um arquivo `.env.dist`, que este sim, estará sob controle de versão. Isso serve como um modelo para que outros desenvolvedores (ou você, no futuro) saibam quais variáveis precisam ser configuradas.

## Características Específicas do Claude

Ao trabalhar comigo neste projeto, tenha em mente essas características específicas:

* **Artifacts:** Quando criar código, documentação extensa, ou conteúdo estruturado, utilize artifacts para facilitar a visualização e reutilização.
* **Contexto Conversacional:** Mantenha o contexto das conversas anteriores e não repita informações desnecessariamente.
* **Explicações Detalhadas:** Quando sugerir mudanças no código, explique o raciocínio por trás das decisões, especialmente se houver trade-offs envolvidos.
* **Verificação de Código:** Sempre que possível, verifique se o código sugerido está alinhado com as práticas estabelecidas neste documento.
* **Feedback Iterativo:** Esteja preparado para refinamentos e melhorias baseados no meu feedback sobre as sugestões fornecidas.

## Preferências de Comunicação

* **Tom:** Mantenha um tom amigável e profissional, mas descontraído.
* **Emojis:** Use emojis moderadamente para tornar a comunicação mais divertida! 😊
* **Estrutura:** Organize suas respostas de forma clara, com títulos e subtítulos quando apropriado.
* **Exemplos:** Sempre que possível, forneça exemplos práticos junto com as explicações.

---

*Este arquivo deve ser atualizado conforme o projeto evolui e novas práticas são adotadas. Mantenha-o sempre sincronizado com as necessidades atuais do projeto!* 📝