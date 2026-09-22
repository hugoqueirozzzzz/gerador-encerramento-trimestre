import re

# futuro do presente (nós) -> pretérito (nós), para textos de introdução
# que no material de origem falam do que "veremos/estudaremos" e no slide
# de encerramento devem falar do que já "vimos/estudamos".
FUTURE_TO_PAST = {
    'estudaremos': 'estudamos',
    'veremos': 'vimos',
    'leremos': 'lemos',
    'meditaremos': 'meditamos',
    'aprenderemos': 'aprendemos',
    'refletiremos': 'refletimos',
    'analisaremos': 'analisamos',
    'compreenderemos': 'compreendemos',
    'conheceremos': 'conhecemos',
    'entenderemos': 'entendemos',
    'observaremos': 'observamos',
    'descobriremos': 'descobrimos',
    'notaremos': 'notamos',
    'examinaremos': 'examinamos',
    'abordaremos': 'abordamos',
    'trataremos': 'tratamos',
    'destacaremos': 'destacamos',
    'ressaltaremos': 'ressaltamos',
    'contemplaremos': 'contemplamos',
    'revisaremos': 'revisamos',
    'relembraremos': 'relembramos',
    'apresentaremos': 'apresentamos',
    'mostraremos': 'mostramos',
    'perceberemos': 'percebemos',
    'identificaremos': 'identificamos',
    'discutiremos': 'discutimos',
    'buscaremos': 'buscamos',
    'encontraremos': 'encontramos',
    'faremos': 'fizemos',
    'diremos': 'dissemos',
    'teremos': 'tivemos',
    'saberemos': 'soubemos',
    'poderemos': 'pudemos',
    'viremos': 'viemos',
    'falaremos': 'falamos',
    'seguiremos': 'seguimos',
    'iniciaremos': 'iniciamos',
    'começaremos': 'começamos',
    'terminaremos': 'terminamos',
    'concluiremos': 'concluímos',
    'chegaremos': 'chegamos',
    'partiremos': 'partimos',
    'caminharemos': 'caminhamos',
    'exploraremos': 'exploramos',
    'consideraremos': 'consideramos',
}

_PATTERN = re.compile(
    r'\b(' + '|'.join(re.escape(k) for k in FUTURE_TO_PAST) + r')\b',
    re.IGNORECASE,
)


def fix_future_to_past(text):
    def _repl(m):
        word = m.group(0)
        key = word.lower()
        replacement = FUTURE_TO_PAST[key]
        if word[0].isupper():
            replacement = replacement[0].upper() + replacement[1:]
        return replacement
    return _PATTERN.sub(_repl, text)
