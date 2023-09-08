AUDIO_DIR := $(WORK)/audio
PROCESSED_DIR := $(AUDIO_DIR)/processed
WHISPER_MODEL := ./models/ggml-medium.bin
AUDIO_LANG := de
NUM_THREADS := 4

INPUT_FILES := $(wildcard $(AUDIO_DIR)/*.mp3)
DOCX_FILES := $(patsubst %.mp3,%.docx,$(INPUT_FILES))

.PHONY: all config

all: $(DOCX_FILES)


config:
	@echo "Python interpreter: $(shell which python)"
	@echo "Python version:     $(shell python --version)"
	@echo "Whisper"
	@echo "  model (WHISPER_MODEL):           $(WHISPER_MODEL)"
	@echo "  model language (AUDIO_LANG):     $(AUDIO_LANG)"
	@echo "  number of threads (NUM_THREADS): $(NUM_THREADS)"
	@echo "Audio files to transcribe:         $(INPUT_FILES)"


%.wav: %.mp3 Makefile
	./ffmpeg-linux-x86_64 -nostdin -loglevel warning -y -i "$<" -ar 16000 -ac 1 -c:a pcm_s16le "$@"


%.speaker: %.wav Makefile
	./diarization.py -i $< -o $@


%.docx: %.wav %.mp3 %.speaker Makefile
	./transcribe.py \
		--wave-input $*.wav \
	   	--diarization-input $*.speaker \
	   	--model $(WHISPER_MODEL) \
		--language $(AUDIO_LANG) \
	    --threads $(NUM_THREADS) \
		--output $@

