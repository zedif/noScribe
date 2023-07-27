#!/usr/bin/env python

import argparse
from contextlib import redirect_stderr
from io import StringIO
import locale
import os
import pickle
import platform

import i18n
from i18n import t
i18n.set('filename_format', '{locale}.{format}')
i18n.load_path.append('./trans')
try:
    app_locale = locale.getdefaultlocale()[0][0:2]
except:
    app_locale = 'en'
i18n.set('fallback', 'en')
i18n.set('locale', app_locale)

HAS_MPS_SUPPORT = platform.mac_ver()[0] >= '12.3'  # MPS needs macOS 12.3+

if platform.system() == 'Darwin':
    os.environ["PYTORCH_ENABLE_MPS_FALLBACK"] = str(1)


def __cli():

    parser = argparse.ArgumentParser(description=t('app_header'))
    parser.add_argument("-i", "--input", help="Input .wav file", required=True)
    parser.add_argument("-o", "--output", help="Output diarization file", required=True)
    args = parser.parse_args()

    diarization = identify_speakers(wav_audio_file=args.input)
    __save(obj=diarization, path=args.output)


def __save(obj, path):
    with open(path, 'wb') as file:
        pickle.dump(obj, file)
    print(f"Diarization saved to '{os.path.abspath(path)}'.")


def identify_speakers(wav_audio_file):
    """
    Given an audio file in .wav format, carry out speaker diarization and
    store the result as the pickle file `diariztion_file`. The pickled diarization
    is a pyannote.core.Annotation object.

    :param str wav_audio_file: Path spoken audio input file in .wav format
    :param str diarization_file: Path to the diarization output file or None
    :return diariztion as pyannote.core.Annotation obectj
    """

    with redirect_stderr(StringIO()) as f:
        print(t('loading_pyannote'))
        from pyannote.audio import Pipeline

        pipeline = Pipeline.from_pretrained('./models/pyannote_config.yaml')
        if HAS_MPS_SUPPORT:
            pipeline.to("mps")
            print('Using Apple Silicon GPU.')

        diarization = pipeline(wav_audio_file)  # apply the pipeline to the audio file

        # read stderr and log it:
        err = f.readline()
        while err != '':
            print(err, 'error')
            err = f.readline()

        return diarization


if __name__ == "__main__":
    __cli()
