import random


class TempoAleatorio:
    def __init__(self):
        self.tempo = random.randint(1, 3)

    def tempo_aleatorio(self):
        return self.tempo