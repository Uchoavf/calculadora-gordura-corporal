# Changelog

Histórico de versões da Calculadora de Gordura Corporal (US Navy + IMC).

---

## v0.4.0 — 2026-09-29

**Correção do cálculo US Navy e melhorias de qualidade/usabilidade**

- **Correção:** fórmulas US Navy estavam na versão em polegadas com entrada em cm (resultado superestimado); agora usam as fórmulas métricas oficiais
- Adicionado `sw.js` (antes era registrado mas não existia): app funciona offline
- `manifest.json`: `start_url`/`scope` relativos (funciona no subcaminho do GitHub Pages) e ícone `icon.svg`
- Enter envia o formulário; botão "Limpar"; dados salvos localmente (localStorage)
- Validação dos limites min/max com destaque do campo e foco no erro; aceita vírgula decimal
- Novos resultados: faixa de peso saudável, massa gorda e massa magra estimadas
- Acessibilidade: labels associados, `aria-pressed`, `aria-live`, foco visível, melhor contraste
- Números formatados em pt-BR; aviso de que são estimativas informativas

---

## v0.3.0 — 2026-05-23

**Adiciona cálculo de IMC e exibe resultados comparativos lado a lado**

`b3dd8fd`

- Adicionado campo Peso (kg) para cálculo do IMC
- Cálculo simultâneo dos dois métodos (US Navy + IMC)
- Resultados exibidos lado a lado em cards comparativos
- Classificação do IMC com cores (abaixo do peso / normal / sobrepeso / obesidade I, II, III)
- Tabela de referência do IMC incluída
- Layout reorganizado com seções rotuladas ("Dados para IMC" / "Medidas para US Navy")
- Responsivo: cards empilhados em telas pequenas
- Documentação (INSTRUCOES.md) atualizada

---

## v0.2.0 — 2026-05-23

**Adiciona documentação e corrige deploy**

`c5c8910`

- Documentação do projeto (INSTRUCOES.md) com instruções, fórmulas e tabelas

`781759d`

- Adicionado `.nojekyll` para corrigir o build do GitHub Pages
- Configurado `gh` como credential helper global do Git

---

## v0.1.0 — 2026-05-23

**Projeto inicial: Calculadora de Gordura Corporal — US Navy**

`250e74d`

- Cálculo de gordura corporal pelo método da Marinha dos EUA (US Navy)
- Seleção de gênero (Masculino / Feminino)
- Campos: altura, pescoço, cintura e quadril (mulheres)
- Validação de entrada com mensagens de erro
- Classificação do resultado com cores
- Tabela de referência integrada
- Design responsivo (mobile-first)
- PWA: manifest.json + service worker
- Publicado no GitHub Pages
- Repositório: [Uchoavf/calculadora-gordura-corporal](https://github.com/Uchoavf/calculadora-gordura-corporal)
