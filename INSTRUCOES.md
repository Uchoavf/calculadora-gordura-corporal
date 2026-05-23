# Calculadora de Gordura Corporal - US Navy

Calculadora web que estima o percentual de gordura corporal usando o **método da Marinha dos EUA (US Navy Body Fat)**.

## Como usar

1. **Selecione o gênero:** Masculino ou Feminino
2. **Preencha as medidas** em centímetros:
   - Altura
   - Pescoço
   - Cintura
   - Quadril (apenas para mulheres)
3. **Clique em "Calcular % de Gordura"**
4. O resultado aparece com o percentual e a classificação

## Fórmula

### Homens

```
% Gordura = 86,010 × log10(Cintura - Pescoço) - 70,041 × log10(Altura) + 36,76
```

### Mulheres

```
% Gordura = 163,205 × log10(Cintura + Quadril - Pescoço) - 97,684 × log10(Altura) - 78,387
```

## Classificação

| Classificação | Mulheres | Homens |
|---|---|---|
| Gordura Essencial | 10–13% | 2–5% |
| Atletas | 14–20% | 6–13% |
| Boa Forma | 21–24% | 14–17% |
| Média | 25–31% | 18–24% |
| Acima do Peso | 32%+ | 25%+ |

## Como medir

- **Pescoço:** logo abaixo do pomo de adão, fita horizontal
- **Cintura:** na altura do umbigo, fita horizontal
- **Quadril:** na parte mais larga dos glúteos

## Tecnologias

- HTML5
- CSS3
- JavaScript (vanilla)
- PWA (manifest + service worker)

## Implantação

Hospedado via **GitHub Pages**. Para rodar localmente, basta abrir o arquivo `index.html` em qualquer navegador.

```
https://uchoavf.github.io/calculadora-gordura-corporal/
```
