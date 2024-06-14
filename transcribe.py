# noScribe - AI-powered Audio Transcription
# Copyright (C) 2023 Kai Dröge
# ported to MAC by Philipp Schneider (gernophil)

# usage: python transcribe.py task.yaml

# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

from contextlib import contextmanager
import os
import datetime
from pathlib import Path
from sys import argv
from munch import Munch
import i18n
from i18n import t
import AdvancedHTMLParser
import yaml

from noscribe_core import *


# Move to core.py
app_dir = os.path.abspath(os.path.dirname(__file__))
i18n.set('filename_format', '{locale}.{format}')
i18n.load_path.append(os.path.join(app_dir, 'trans'))

default_html = """
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">
<html >
<head >
<meta charset="UTF-8" />
<meta name="qrichtext" content="1" />
<style type="text/css" >
p, li { white-space: pre-wrap; }
</style>
<style type="text/css" > 
 p { font-size: 0.9em; } 
 .MsoNormal { font-family: "Arial"; font-weight: 400; font-style: normal; font-size: 0.9em; }
 @page WordSection1 {mso-line-numbers-restart: continuous; mso-line-numbers-count-by: 1; mso-line-numbers-start: 1; }
 div.WordSection1 {page:WordSection1;} 
</style>
</head>
<body style="font-family: 'Arial'; font-weight: 400; font-style: normal" >
</body>
</html>"""


def _logn(msg='', *_, **__):
    print(msg)


def _pass(*_):
    pass


# Helper for text only output
def html_node_to_text(node: AdvancedHTMLParser.AdvancedTag) -> str:
    """
    Recursively get all text from a html node and its children.
    """
    # For text nodes, return their value directly
    if AdvancedHTMLParser.isTextNode(node): # node.nodeType == node.TEXT_NODE:
        return node
    # For element nodes, recursively process their children
    elif AdvancedHTMLParser.isTagNode(node):
        text_parts = []
        for child in node.childBlocks:
            text = html_node_to_text(child)
            if text:
                text_parts.append(text)
        # For block-level elements, prepend and append newlines
        if node.tagName.lower() in ['p', 'div', 'ul', 'ol', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'br']:
            if node.tagName.lower() == 'br':
                return '\n'
            else:
                return '\n' + ''.join(text_parts).strip() + '\n'
        else:
            return ''.join(text_parts)
    else:
        return ''

def html_to_text(parser: AdvancedHTMLParser.AdvancedHTMLParser) -> str:
    return html_node_to_text(parser.body)


# Helper for WebVTT output
def vtt_escape(txt: str) -> str:
    txt = txt.replace('&', '&amp;')
    txt = txt.replace('<', '&lt;')
    txt = txt.replace('>', '&gt;')
    while txt.find('\n\n') > -1:
        txt = txt.replace('\n\n', '\n')
    return txt

def ms_to_webvtt(milliseconds) -> str:
    """converts milliseconds to the time stamp of WebVTT (HH:MM:SS.mmm)
    """
    # 1 hour = 3600000 milliseconds
    # 1 minute = 60000 milliseconds
    # 1 second = 1000 milliseconds
    hours, milliseconds = divmod(milliseconds, 3600000)
    minutes, milliseconds = divmod(milliseconds, 60000)
    seconds, milliseconds = divmod(milliseconds, 1000)
    return "{:02d}:{:02d}:{:02d}.{:03d}".format(hours, minutes, seconds, milliseconds)

def html_to_webvtt(parser: AdvancedHTMLParser.AdvancedHTMLParser, media_path: str):
    vtt = 'WEBVTT '
    paragraphs = parser.getElementsByTagName('p')
    # The first paragraph contains the title
    vtt += vtt_escape(paragraphs[0].textContent) + '\n\n'
    # Next paragraph contains info about the transcript. Add as a note.
    vtt += vtt_escape('NOTE\n' + html_node_to_text(paragraphs[1])) + '\n\n'
    # Add media source:
    vtt += f'NOTE media: {media_path}\n\n'

    #Add all segments as VTT cues
    segments = parser.getElementsByTagName('a')
    i = 0
    for i in range(len(segments)):
        segment = segments[i]
        name = segment.attributes['name']
        if name is not None:
            name_elems = name.split('_', 4)
            if len(name_elems) > 1 and name_elems[0] == 'ts':
                start = ms_to_webvtt(int(name_elems[1]))
                end = ms_to_webvtt(int(name_elems[2]))
                spkr = name_elems[3]
                txt = vtt_escape(html_node_to_text(segment))
                vtt += f'{i+1}\n{start} --> {end}\n<v {spkr}>{txt.lstrip()}\n\n'
    return vtt


def overlap_len(ss_start, ss_end, ts_start, ts_end):
    # ss...: speaker segment start and end in milliseconds (from pyannote)
    # ts...: transcript segment start and end (from whisper.cpp)
    # returns overlap percentage, i.e., "0.8" = 80% of the transcript segment overlaps with the speaker segment from pyannote
    if ts_end < ss_start: # no overlap, ts is before ss
        return None

    if ts_start > ss_end: # no overlap, ts is after ss
        return 0.0

    ts_len = ts_end - ts_start
    if ts_len <= 0:
        return None

    # ss & ts have overlap
    overlap_start = max(ss_start, ts_start) # Whichever starts later
    overlap_end = min(ss_end, ts_end) # Whichever ends sooner
    ol_len = overlap_end - overlap_start + 1
    return ol_len / ts_len

def find_speaker(diarization, transcript_start, transcript_end, overlapping_speech_selected) -> str:
    # Looks for the shortest segment in diarization that has at least 80% overlap
    # with transcript_start - trancript_end.
    # Returns the speaker name if found.
    # If only an overlap < 80% is found, this speaker name ist returned.
    # If no overlap is found, an empty string is returned.
    spkr = ''
    overlap_found = 0
    overlap_threshold = 0.8
    segment_len = 0
    is_overlapping = False

    for segment in diarization:
        t = overlap_len(segment["start"], segment["end"], transcript_start, transcript_end)
        if t is None: # we are already after transcript_end
            break

        current_segment_len = segment["end"] - segment["start"] # Length of the current segment
        current_segment_spkr = f'S{segment["label"][8:]}' # shorten the label: "SPEAKER_01" > "S01"

        if overlap_found >= overlap_threshold: # we already found a fitting segment, compare length now
            if (t >= overlap_threshold) and (current_segment_len < segment_len): # found a shorter (= better fitting) segment that also overlaps well
                is_overlapping = True
                overlap_found = t
                segment_len = current_segment_len
                spkr = current_segment_spkr
        elif t > overlap_found: # no segment with good overlap yet, take this if the overlap is better then previously found
            overlap_found = t
            segment_len = current_segment_len
            spkr = current_segment_spkr

    if overlapping_speech_selected and is_overlapping:
        return f"//{spkr}"
    else:
        return spkr


class TranscriptSaver(object):

    def __init__(self, file_path, fmt, audio_file_path, logn_callback):
        self.fmt = fmt
        self._validate_format()
        self.file_path = file_path
        self.audio_file_path = audio_file_path
        self.logn_callback = logn_callback
        self.last_save = datetime.datetime.now()

    def _validate_format(self):
        if self.fmt not in ['html', 'txt', 'vtt']:
            raise TypeError(f'Invalid file type "{fmt}".')

    def save(self, d):
        text = self._parse(d)
        with self._open() as file:
            file.write(text)
            file.flush()
        self.last_save = datetime.datetime.now()

    @contextmanager
    def _open(self):
        """
        Create a context manager that encapsulates the selection of the fallback
        file name and can be used like Python's built-in open() function with
        the with keyword:

             with self._open(...) as file:
                 ...
        """
        try:
            with open(self.file_path, 'w', encoding='utf-8') as file:
                yield file
        except (IOError, OSError):
            original_path = Path(self.file_path)
            fallback_path = original_path.with_stem(f'{original_path.stem}_1')

            if fallback_path.exists():
                raise Exception(t('rescue_saving_failed'))
            else:
                self.logn_callback()
                self.logn_callback(
                    f'Rescue saving to {fallback_path}',
                    'error',
                    f'file://{fallback_path}'
                )
                with open(fallback_path, 'w', encoding='utf-8') as file:
                    yield file

    def _parse(self, d):
       if self.fmt == 'html':
           return d.asHTML()
       elif self.fmt == 'txt':
           return html_to_text(d)
       elif self.fmt == 'vtt':
           return html_to_webvtt(d, self.audio_file_path)


def transcribe(cfg, log_callback, logn_callback, set_progress_callback, user_cancel_callback):

    logn_callback(t('loading_whisper'))

    # prepare transcript html
    d = AdvancedHTMLParser.AdvancedHTMLParser()
    d.parseStr(default_html)
    saver = TranscriptSaver(
        cfg.my_transcript_file,
        cfg.transcript.file_ext,
        cfg.audio.file,
        logn_callback
    )

    # add audio file path:
    tag = d.createElement("meta")
    tag.name = "audio_source"
    tag.content = cfg.audio.file
    d.head.appendChild(tag)

    # add app version:
    """ # removed because not really necessary
    tag = d.createElement("meta")
    tag.name = "noScribe_version"
    tag.content = app_version
    d.head.appendChild(tag)
    """

    #add WordSection1 (for line numbers in MS Word) as main_body
    main_body = d.createElement('div')
    main_body.addClass('WordSection1')
    d.body.appendChild(main_body)

    # header
    p = d.createElement('p')
    p.setStyle('font-weight', '600')
    p.appendText(Path(cfg.audio.file).stem) # use the name of the audio file (without extension) as the title
    main_body.appendChild(p)

    # subheader
    p = d.createElement('p')
    s = d.createElement('span')
    s.setStyle('color', '#909090')
    s.setStyle('font-size', '0.8em')
    s.appendText(t('doc_header', version=app_version))
    br = d.createElement('br')
    s.appendChild(br)

    s.appendText(t('doc_header_audio', file=cfg.audio.file))
    br = d.createElement('br')
    s.appendChild(br)

    s.appendText(f'({cfg.option_info})')

    p.appendChild(s)
    main_body.appendChild(p)

    p = d.createElement('p')
    main_body.appendChild(p)

    speaker = ''
    prev_speaker = ''

    try:
        from faster_whisper import WhisperModel
        model = WhisperModel(cfg.model,
                             device=cfg.device,
                             cpu_threads=cfg.number_threads,
                             compute_type=cfg.compute_type,
                             local_files_only=True)
        logn_callback('model loaded', where='file')

        if user_cancel_callback():
            raise Exception(t('err_user_cancelation'))

        whisper_lang = cfg.language if cfg.language != 'auto' else None

        segments, info = model.transcribe(
            cfg.audio.tmp_file, language=whisper_lang,
            beam_size=1, temperature=0, word_timestamps=True,
            initial_prompt=cfg.prompt, vad_filter=True,
            vad_parameters=dict(min_silence_duration_ms=200,
                                threshold=cfg.vad_threshold))

        if cfg.language == "auto":
            logn_callback("Detected language '%s' with probability %f" % (info.language, info.language_probability))

        if user_cancel_callback():
            raise Exception(t('err_user_cancelation'))

        logn_callback(t('start_transcription'))
        logn_callback()

        last_segment_end = 0
        last_timestamp_ms = 0
        first_segment = True

        for segment in segments:
            # check for user cancelation
            if user_cancel_callback():
                if cfg.transcript.auto_save:
                    saver.save(d)
                    logn_callback()
                    log_callback(t('transcription_saved'))
                    logn_callback(cfg.my_transcript_file, link=f'file://{cfg.my_transcript_file}')

                raise Exception(t('err_user_cancelation'))

            # get time of the segment in milliseconds
            start = round(segment.start * 1000.0)
            end = round(segment.end * 1000.0)
            # if we skipped a part at the beginning of the audio we have to add this here again, otherwise the timestaps will not match the original audio:
            orig_audio_start = cfg.audio.start + start
            orig_audio_end = cfg.audio.start + end

            if cfg.timestamps:
                ts = ms_to_str(orig_audio_start)
                ts = f'[{ts}]'

            # check for pauses and mark them in the transcript
            if (cfg.pause > 0) and (start - last_segment_end >= cfg.pause * 1000): # (more than x seconds with no speech)
                pause_len = round((start - last_segment_end)/1000)
                if pause_len >= 60: # longer than 60 seconds
                    pause_str = ' ' + t('pause_minutes', minutes=round(pause_len/60))
                elif pause_len >= 10: # longer than 10 seconds
                    pause_str = ' ' + t('pause_seconds', seconds=pause_len)
                else: # less than 10 seconds
                    pause_str = ' (' + (cfg.pause_marker * pause_len) + ')'

                if first_segment:
                    pause_str = pause_str.lstrip() + ' '

                orig_audio_start_pause = cfg.audio.start + last_segment_end
                orig_audio_end_pause = cfg.audio.start + start
                a = d.createElement('a')
                a.name = f'ts_{orig_audio_start_pause}_{orig_audio_end_pause}_{speaker}'
                a.appendText(pause_str)
                p.appendChild(a)
                log_callback(pause_str)
                if first_segment:
                    logn_callback()
                    logn_callback()
            last_segment_end = end

            # write text to the doc
            # diarization (speaker detection)?
            seg_text = segment.text
            seg_html = seg_text

            if cfg.speaker_detection != 'none':
                with open(cfg.diarization_file, 'r') as file:
                    diarization = yaml.safe_load(file)

                    new_speaker = find_speaker(diarization, start, end, cfg.overlapping_speech_selected)
                    if (speaker != new_speaker) and (new_speaker != ''): # speaker change
                        if new_speaker[:2] == '//': # is overlapping speech, create no new paragraph
                            prev_speaker = speaker
                            speaker = new_speaker
                            seg_text = f' {speaker}:{seg_text}'
                            seg_html = seg_text
                        elif (speaker[:2] == '//') and (new_speaker == prev_speaker): # was overlapping speech and we are returning to the previous speaker 
                            speaker = new_speaker
                            seg_text = f'//{seg_text}'
                            seg_html = seg_text
                        else: # new speaker, not overlapping
                            if speaker[:2] == '//': # was overlapping speech, mark the end
                                last_elem = p.lastElementChild
                                if last_elem:
                                    last_elem.appendText('//')
                                else:
                                    p.appendText('//')
                                log_callback('//')
                            p = d.createElement('p')
                            main_body.appendChild(p)
                            if not first_segment:
                                logn_callback()
                                logn_callback()
                            speaker = new_speaker
                            # add timestamp
                            if cfg.timestamps:
                                seg_html = f'{speaker} <span style="color: {cfg.timestamp_color}" >{ts}</span>:{seg_text}'
                                seg_text = f'{speaker} {ts}:{seg_text}'
                                last_timestamp_ms = start
                            else:
                                if cfg.transcript.file_ext != 'vtt': # in vtt files, speaker names are added as special voice tags so skip this here
                                    seg_text = f'{speaker}:{seg_text}'
                                    seg_html = seg_text
                                else:
                                    seg_html = seg_text.lstrip()
                                    seg_text = f'{speaker}:{seg_text}'

                    else: # same speaker
                        if cfg.timestamps:
                            if (start - last_timestamp_ms) > cfg.timestamp_interval:
                                seg_html = f' <span style="color: {cfg.timestamp_color}" >{ts}</span>{seg_text}'
                                seg_text = f' {ts}{seg_text}'
                                last_timestamp_ms = start
                            else:
                                seg_html = seg_text

            else: # no speaker detection
                if cfg.timestamps and (first_segment or (start - last_timestamp_ms) > cfg.timestamp_interval):
                    seg_html = f' <span style="color: {cfg.timestamp_color}" >{ts}</span>{seg_text}'
                    seg_text = f' {ts}{seg_text}'
                    last_timestamp_ms = start
                else:
                    seg_html = seg_text
                # avoid leading whitespace in first paragraph
                if first_segment:
                    seg_text = seg_text.lstrip()
                    seg_html = seg_html.lstrip()

            # Mark confidence level (not implemented yet in html)
            # cl_level = round((segment.avg_logprob + 1) * 10)
            # TODO: better cl_level for words based on https://github.com/Softcatala/whisper-ctranslate2/blob/main/src/whisper_ctranslate2/transcribe.py
            # if cl_level > 0:
            #     r.style = d.styles[f'noScribe_cl{cl_level}']

            # Create bookmark with audio timestamps start to end and add the current segment.
            # This way, we can jump to the according audio position and play it later in the editor.
            a_html = f'<a name="ts_{orig_audio_start}_{orig_audio_end}_{speaker}" >{seg_html}</a>'
            a = d.createElementFromHTML(a_html)
            p.appendChild(a)

            log_callback(seg_text)

            first_segment = False

            # auto save
            if cfg.transcript.auto_save:
                if (datetime.datetime.now() - saver.last_save).total_seconds() > 20:
                    saver.save(d)

            progr = round((segment.end/info.duration) * 100)
            set_progress_callback(3, progr)

        saver.save(d)
        logn_callback()
        logn_callback()
        logn_callback(t('transcription_finished'), 'highlight')
        if cfg.transcript.file != cfg.my_transcript_file: # used alternative filename because saving under the initial name failed
            log_callback(t('rescue_saving'))
            logn_callback(cfg.my_transcript_file, link=f'file://{cfg.my_transcript_file}')
        else:
            log_callback(t('transcription_saved'))
            logn_callback(cfg.my_transcript_file, link=f'file://{cfg.my_transcript_file}')

    except Exception as e:
        logn_callback()
        logn_callback(t('err_transcription'), 'error')
        logn_callback(e, 'error')
        return


if __name__ == '__main__':

    if len(argv) != 2:
        print(f'Usage: python {argv[0]} task.yaml')
    else:
        config_file_name = argv[1]

        config = ''
        try:
            with open(config_file_name, 'r', encoding='utf-8') as file:
                config = Munch.fromDict(yaml.safe_load(file))
        except Exception as e:
            print(f'Could not open config file: {e}')

        if config:
            transcribe(config, log_callback=_logn, logn_callback=_logn, set_progress_callback=_pass, user_cancel_callback=_pass)

