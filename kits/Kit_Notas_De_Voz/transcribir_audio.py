import os
import whisper
from tkinter import Tk, filedialog

# FFmpeg local (estable, fuera de OneDrive)
ffmpeg_exe = r"C:\ffmpeg_local\ffmpeg.exe"
ffmpeg_dir = r"C:\ffmpeg_local"

os.environ["FFMPEG_BINARY"] = ffmpeg_exe
os.environ["PATH"] = ffmpeg_dir + os.pathsep + os.environ.get("PATH", "")

# Selector de archivo (Windows Explorer)
Tk().withdraw()
audio_file = filedialog.askopenfilename(
    title="Selecciona el archivo de audio",
    filetypes=[("Audio", "*.mp3 *.wav *.m4a *.aac *.ogg *.flac"), ("Todos", "*.*")]
)

if not audio_file:
    raise SystemExit("No se seleccionó ningún archivo.")

print("Archivo seleccionado:\n", audio_file)
print("FFmpeg usado:\n", ffmpeg_exe)

# Transcribir
model = whisper.load_model("small")  # tiny | base | small | medium | large
result = model.transcribe(audio_file, language="es", fp16=False, verbose=False)

print("\n--- TRANSCRIPCION ---\n")
print(result["text"])

# Guardar txt junto al audio
out_txt = os.path.splitext(audio_file)[0] + "_transcripcion.txt"
with open(out_txt, "w", encoding="utf-8") as f:
    f.write(result["text"])

print("\nGuardado en:\n", out_txt)
