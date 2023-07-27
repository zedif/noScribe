#!/usr/bin/env python

from contextlib import redirect_stderr
import platform
import os
from io import StringIO
from i18n import t
import pickle
import argparse

if platform.system() == 'Darwin':
    os.environ["PYTORCH_ENABLE_MPS_FALLBACK"] = str(1)

def __cli():

    parser = argparse.ArgumentParser(description=t('app_header'))
    parser.add_argument("-i", "--input", help="Input .wav file", required=True)
    parser.add_argument("-o", "--output", help="Output diarization file", required=True)
    args = parser.parse_args()

    diarization_obj = diarization(wav_audio_file=args.input)
    __save(obj=diarization_obj, path=args.output)


def __save(obj, path):
    with open(path, 'wb') as file:
        pickle.dump(obj, file)
    print(f"Diarization saved to '{os.path.abspath(path)}'.")


def diarization(wav_audio_file):
    """
    Given an audio file in .wav format, carry out speaker diarization and
    store the result as the pickle file `diariztion_file`. The pickled diarization
    is a pyannote.core.Annotation object.

    :param str wav_audio_file: Path spoken audio input file in .wav format
    :param str diarization_file: Path to the diarization output file or None
    :return diariztion as pyannote.core.Annotation obectj
    """

    with redirect_stderr(StringIO()) as f:
        print(t('start_identifiying_speakers'), 'highlight')
        print(t('loading_pyannote'))
        from pyannote.audio import Pipeline

        pipeline = Pipeline.from_pretrained('./models/pyannote_config.yaml')
        # if platform.machine() == "arm64": # Intel should also support MPS
        if platform.system() == "Darwin":
            if platform.mac_ver()[0] >= '12.3':  # MPS needs macOS 12.3+
                pipeline.to("mps")
                print('Using Apple Silicon GPU.')

        diarization_obj = pipeline(wav_audio_file) # apply the pipeline to the audio file

        # read stderr and log it:
        err = f.readline()
        while err != '':
            print(err, 'error')
            err = f.readline()

        return diarization_obj

if __name__ == "__main__":
    __cli()