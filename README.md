# Gerador de Slide de Encerramento — EBD em Foco

App web (Flask) que monta o "slide de encerramento" do trimestre a partir das
13 lições completas (.pptx), nos mesmos moldes do app do Tópico 1.

## O que ele faz

Para cada lição enviada, extrai automaticamente:
1. O slide de título da lição (removendo a caixa "Nome do Professor")
2. O bumper "INTRODUÇÃO" + os slides de conteúdo da introdução (o texto tem
   os verbos no futuro — "estudaremos", "veremos" etc. — automaticamente
   ajustados para o passado, já que é um material de encerramento)
3. Apenas os 3 slides-bumper dos tópicos (I, II, III), sem o conteúdo interno

O número da lição é detectado automaticamente pelo texto "LIÇÃO NN" no slide
de título — não importa a ordem em que os arquivos são enviados.

No final, monta: capa fixa da turma + blocos das lições em ordem + pacote
promocional fixo da turma (slides 101–107 do modelo original).

## Turmas

Os pacotes fixos de abertura/rodapé de **adultos** e **jovens** já estão
embutidos no app (`static_packages/`), extraídos dos modelos que você enviou.
Se o pacote promocional mudar em um próximo trimestre, é só gerar novos
`pkg_head_*.pptx` / `pkg_tail_*.pptx` a partir do modelo atualizado (script
usado para isso: peça pro Claude gerar de novo a partir do novo modelo).

## Rodando localmente

```bash
pip install -r requirements.txt
python app.py
# abre em http://localhost:5000
```

## Deploy no Render (mesmo fluxo do Tópico 1)

1. Suba este projeto para um repositório no GitHub
2. No Render.com: New → Web Service → conecte o repositório
3. Environment: Docker (usa o Dockerfile deste projeto)
4. Deploy automático a cada push

## Estrutura

```
app.py              - rotas Flask (upload e geração)
build_deck.py        - monta o pptx final
lesson_detect.py      - detecta o número da lição e os pontos de corte
extract_logic.py       - classifica os slides de cada lição (título/introdução/bumper)
verb_fix.py             - de-para futuro → passado nos textos de introdução
pptx_merge.py             - motor de cópia de slides entre arquivos .pptx
                            (preserva layout/master/tema/mídia)
static_packages/           - pacotes fixos de capa e rodapé (adultos/jovens)
templates/index.html       - página de upload
```

## Limitações conhecidas

- Espera que cada lição siga o padrão: título → ... → "INTRODUÇÃO" → conteúdo
  → "I – ..." → ... → "II – ..." → ... → "III – ..." → .... Lições fora
  desse padrão geram uma mensagem de erro específica.
- O de-para de verbos (`verb_fix.py`) cobre os verbos mais comuns nesse tipo
  de texto, mas não é exaustivo — vale conferir os slides de introdução
  gerados antes de publicar.
