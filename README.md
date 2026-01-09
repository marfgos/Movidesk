# Movidesk

Projeto criado para atualizar o indicador de chamados do MoviDesk.

## Transcrição de Áudio com Whisper

O repositório também inclui um script utilitário para transcrever arquivos de áudio
usando o modelo Whisper da OpenAI. Para utilizá-lo:

1. Instale as dependências do projeto:

   ```bash
   pip install -r requirements.txt
   ```

2. Execute o script informando o caminho do arquivo de áudio. Você pode alterar o
   modelo e o idioma conforme necessário:

   ```bash
   python transcribe_whisper.py /caminho/para/entrevista.wav --model base --language pt
   ```

O resultado será impresso diretamente no terminal.
