"""Ferramenta de linha de comando simples para transcrever arquivos de áudio com o Whisper."""

from __future__ import annotations

import argparse
import pathlib
import sys

import whisper


DEFAULT_MODEL = "base"
DEFAULT_LANGUAGE = "pt"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Analisa argumentos da linha de comando.

    Parameters
    ----------
    argv: list[str] | None
        Lista de argumentos recebida da linha de comando. Se ``None`` os
        argumentos serão lidos de ``sys.argv``.
    """

    parser = argparse.ArgumentParser(
        description=(
            "Transcreve um arquivo de áudio com o modelo Whisper da OpenAI."
        )
    )
    parser.add_argument(
        "audio_path",
        type=pathlib.Path,
        help="Caminho para o arquivo de áudio que será transcrito.",
    )
    parser.add_argument(
        "--model",
        default=DEFAULT_MODEL,
        help=(
            "Nome do modelo Whisper a ser utilizado (tiny, base, small, medium, large). "
            f"Padrão: {DEFAULT_MODEL}."
        ),
    )
    parser.add_argument(
        "--language",
        default=DEFAULT_LANGUAGE,
        help=f"Idioma esperado na transcrição. Padrão: {DEFAULT_LANGUAGE}.",
    )

    return parser.parse_args(argv)


def transcribe_audio(audio_path: pathlib.Path, model_name: str, language: str) -> str:
    """Transcreve ``audio_path`` usando o Whisper e retorna o texto produzido."""

    if not audio_path.exists():
        raise FileNotFoundError(f"Arquivo de áudio não encontrado: {audio_path}")

    model = whisper.load_model(model_name)
    result = model.transcribe(str(audio_path), language=language)
    return result["text"].strip()


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        transcript = transcribe_audio(args.audio_path, args.model, args.language)
    except FileNotFoundError as exc:
        print(exc, file=sys.stderr)
        return 1

    print("==== TRANSCRIÇÃO COMPLETA ====")
    print(transcript)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
