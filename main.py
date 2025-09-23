import os

import win32com.client  # Uses your local PowerPoint installation
from deep_translator import GoogleTranslator
from gtts import gTTS
from moviepy import AudioFileClip, ImageClip, concatenate_videoclips
from powerpoint import Presentation

# -----------------------------
# CONFIGURATION
# -----------------------------
CURRENT_DIRECTORY = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(CURRENT_DIRECTORY, "output_videos")
TMP_DIR = os.path.join(CURRENT_DIRECTORY, "tmp_assets")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TMP_DIR, exist_ok=True)
# Languages: default set, but user can add more here.
LANGUAGES = [
    {"code": "en", "label": "English"},
    {"code": "fr", "label": "French"},
    {"code": "es", "label": "Spanish"},
    {"code": "pt", "label": "Portuguese"},
    {"code": "hi", "label": "Hindi"},
    {"code": "si", "label": "Sinhala"},
]
# Google TTS config
default_voice = {
    "en": {"language_code": "en-US", "name": "en-US-Neural2-D"},
    "fr": {"language_code": "fr-FR", "name": "fr-FR-Neural2-A"},
    "es": {"language_code": "es-ES", "name": "es-ES-Neural2-A"},
    "pt": {"language_code": "pt-BR", "name": "pt-BR-Neural2-A"},
    "hi": {"language_code": "hi-IN", "name": "hi-IN-Neural2-A"},
    "si": {"language_code": "si-LK", "name": "si-LK-Standard-A"},
}
POST_AUDIO_PADDING = 1.5  # extra seconds per slide after narration ends


# -----------------------------
# STEP 1: Extract presenter notes from PPTX
# -----------------------------
def extract_notes_from_pptx(pptx_path):
    print(f"Extracting notes from {pptx_path}...")
    prs = Presentation(pptx_path)
    slides_data = []
    for i, slide in enumerate(prs.slides):
        print(f" Slide {i+1}/{len(prs.slides)}")
        notes_text = ""
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes_text = "\n".join(
                p.text for p in slide.notes_slide.notes_text_frame.paragraphs
            ).strip()
        slides_data.append({"index": i, "notes": notes_text or ""})
    return slides_data


# -----------------------------
# STEP 2: Export slides as images using local PowerPoint
# -----------------------------
def export_slides_as_images(pptx_path, out_dir):
    print(f"Exporting slides as images to {out_dir}...")
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True
    presentation = powerpoint.Presentations.Open(
        os.path.abspath(pptx_path), WithWindow=False
    )
    base = os.path.join(os.path.abspath(out_dir), "slide_image")
    presentation.SaveAs(base, 17)
    presentation.Close()
    powerpoint.Quit()
    images = [
        os.path.join(base, f) for f in sorted(os.listdir(base)) if f.endswith(".JPG")
    ]
    return images


# -----------------------------
# STEP 3: Translate text (Google Translate API)
# -----------------------------
def translate_texts(texts, target_language):
    results = []
    for t in texts:
        if not t.strip():
            results.append("")
        else:
            res = GoogleTranslator(source="auto", target=target_language).translate(t)
            results.append(res)
    return results


# -----------------------------
# STEP 4: Generate narration (Google TTS)
# -----------------------------
def synthesize_speech(text, lang_code, out_path):
    print(f"  Synthesizing speech to {out_path}...")
    tts = gTTS(text=text, lang=lang_code)
    tts.save(out_path)
    return out_path


# -----------------------------
# STEP 5: Build video from slides + narration
# -----------------------------
def assemble_video(slide_images, audio_files, out_path):
    print(f" Assembling video to {out_path}...")
    clips = []
    for img, aud in zip(slide_images, audio_files):
        print(f"  Processing slide image {img} with audio {aud}...")
        if aud:
            audio_clip = AudioFileClip(aud)
            duration = audio_clip.duration + POST_AUDIO_PADDING
            slide_clip = ImageClip(img, duration=duration).with_audio(audio_clip)
            print("2", slide_clip)
        else:
            slide_clip = ImageClip(img, duration=4.0)  # default 4s if no audio
            print("1", slide_clip)
        clips.append(slide_clip)
        print(clips)
    print(" Concatenating clips...")
    final = concatenate_videoclips(clips, method="compose")
    print(f" Writing final video file {out_path}...")
    final.write_videofile(
        out_path,
        fps=24,
        codec="libx264",
        audio=True,
        audio_codec="libmp3lame",
        threads=4,
    )


# -----------------------------
# MAIN PIPELINE
# -----------------------------
def run_pipeline(pptx_path, languages):
    slides = extract_notes_from_pptx(pptx_path)
    notes = [s["notes"] for s in slides]
    images = export_slides_as_images(pptx_path, TMP_DIR)
    for lang in languages:
        code = lang["code"]
        print(f"Processing {lang['label']} ({code})...")
        # Translate notes if not English
        if code != "en":
            lang_notes = translate_texts(notes, code)
        else:
            lang_notes = notes
        # Generate TTS audio per slide
        print(" Generating narration...")
        audio_files = []
        for i, txt in enumerate(lang_notes):
            print(f"  Slide {i+1}/{len(lang_notes)}")
            if txt.strip():
                audio_path = os.path.join(TMP_DIR, f"slide_{i+1}_{code}.mp3")
                synthesize_speech(txt, code, audio_path)
                audio_files.append(audio_path)
            else:
                audio_files.append(None)
        # Build MP4
        print(" Assembling video...")
        out_file = os.path.join(
            OUTPUT_DIR, f"{os.path.splitext(os.path.basename(pptx_path))[0]}_{code}.mp4"
        )
        print(out_file)
        assemble_video(images, audio_files, out_file)
        print(f"Saved: {out_file}")


# -----------------------------
# RUN SCRIPT
# -----------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert PPTX with presenter notes into narrated MP4s in multiple languages."
    )
    parser.add_argument("pptx", help="Path to PowerPoint file (.pptx)")
    parser.add_argument(
        "--languages", nargs="*", help="Languages to output (default all)"
    )
    args = parser.parse_args()
    langs = LANGUAGES
    run_pipeline(
        args.pptx,
        langs,
    )
