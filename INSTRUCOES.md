# Calculadora de Gordura Corporal - US Navy + IMC

Calculadora web que estima o percentual de gordura corporal usando dois métodos e permite comparar os resultados lado a lado:

1. **US Navy Body Fat** — Método da Marinha dos EUA
2. **IMC** — Índice de Massa Corporal

## Como usar

1. **Selecione o gênero:** Masculino ou Feminino
2. **Preencha os dados:**
   - Peso (kg) e Altura (cm) → para o IMC
   - Pescoço, Cintura e Quadril (mulheres) → para o US Navy
3. **Clique em "Calcular os Dois Métodos"**
4. Compare os resultados exibidos lado a lado

## Fórmulas

### US Navy (medidas em cm)

**Homens:**
```
Densidade = 1,0324 - 0,19077 × log10(Cintura - Pescoço) + 0,15456 × log10(Altura)
% Gordura = 495 ÷ Densidade - 450
```

**Mulheres:**
```
Densidade = 1,29579 - 0,35004 × log10(Cintura + Quadril - Pescoço) + 0,22100 × log10(Altura)
% Gordura = 495 ÷ Densidade - 450
```

### IMC

```
IMC = Peso (kg) ÷ Altura² (m)
```

## Tabelas de Referência

### US Navy — Gordura Corporal

| Classificação | Mulheres | Homens |
|---|---|---|
| Gordura Essencial | 10–13% | 2–5% |
| Atletas | 14–20% | 6–13% |
| Boa Forma | 21–24% | 14–17% |
| Média | 25–31% | 18–24% |
| Acima do Peso | 32%+ | 25%+ |

### IMC

| Classificação | IMC (kg/m²) |
|---|---|
| Abaixo do peso | < 18,5 |
| Peso normal | 18,5 – 24,9 |
| Sobrepeso | 25 – 29,9 |
| Obesidade Grau I | 30 – 34,9 |
| Obesidade Grau II | 35 – 39,9 |
| Obesidade Grau III | ≥ 40 |

## Como medir

- **Pescoço:** logo abaixo do pomo de adão, fita horizontal
- **Cintura:** na altura do umbigo, fita horizontal
- **Quadril:** na parte mais larga dos glúteos

## Tecnologias

- HTML5
- CSS3
- JavaScript (vanilla)
- PWA (manifest + service worker, funciona offline)

## Implantação

Hospedado via **GitHub Pages**. Para rodar localmente, basta abrir o arquivo `index.html` em qualquer navegador.

```
https://uchoavf.github.io/calculadora-gordura-corporal/
```
