import whisper
import sounddevice as sd
import scipy.io.wavfile as wav
import os
import requests
import numpy as np
import threading
import re

# ========== CONFIG ==========

AUDIO_FILENAME = "consultation.wav"
TRANSCRIPT_FILENAME = "consultation.txt"
MODEL = "base"
TOGETHER_MODEL = "mistralai/Mistral-7B-Instruct-v0.1"

PROMPT_INSTRUCTIONS = [
    "Never remove parentheses",
    "Never alter text within parentheses",
    "Clean this up for grammar and punctuation only.",
    "If the title Miss appears before a name, replace it with Ms. instead.",
    "Do not use Miss at all, even if it seems correct in context—use Ms. universally as the default respectful form of address for women.",
    "Leave all other professional or formal titles (e.g., Dr., Mr., Prof.) unchanged.",
    "If a sentence starts with the word 'Results', do not add, remove, or change it by adding 'the' or any other article.",
    "Do not change the tone or meaning.",
    "Do not add or remove content.",
    "Use professional clinical language.",
    "Do not replace the word assist with the word help.",
    "Do not add the word please before the word assist. If the sentence starts with assist, just leave it, even if it may be grammatically incorrect in this case.",
    "If an acronym is present, just leave it as the acronym and do not spell it out (e.g., ADHD).",
    "Fix sentence fragments and run-on sentences."
]

# ========== VOICE COMMANDS CLEANUP ==========

def apply_voice_commands(transcript):
    transcript = re.sub(r"[.,!?]?\s*\bnew paragraph\b", "\n\n", transcript, flags=re.IGNORECASE)

    # Handle in parentheses / end parentheses
    def paren_replacer(match):
        inner = match.group(1).strip()
        return f"({inner})"
    transcript = re.sub(
        r"\bin parenthesis\b\s*(.*?)\s*\bend parenthesis\b",
        paren_replacer,
        transcript,
        flags=re.IGNORECASE | re.DOTALL
    )

    # Smart quotes
    transcript = re.sub(r"\bquote start\b", "“", transcript, flags=re.IGNORECASE)
    transcript = re.sub(r"\bquote end\b", "”", transcript, flags=re.IGNORECASE)

    # Basic punctuation
    punctuation_map = {
        r"\bcomma\b": ",",
        r"\bperiod\b": ".",
        r"\bsemicolon\b": ";",
        r"\bcolon\b": ":"
    }
    for pattern, symbol in punctuation_map.items():
        transcript = re.sub(pattern, symbol, transcript, flags=re.IGNORECASE)

    # Placeholders
    placeholders = {
        r"\binsert client name\b": "[Client Name]",
        r"\binsert date\b": "[Date]"
    }
    for pattern, replacement in placeholders.items():
        transcript = re.sub(pattern, replacement, transcript, flags=re.IGNORECASE)

    # Capitalize single word (e.g., "capitalize extraordinary")
    def capitalize_word(match):
        return match.group(1).capitalize()
    transcript = re.sub(r"\bcapitalize (\w+)\b", capitalize_word, transcript, flags=re.IGNORECASE)

    # Clean up spacing and repeated punctuation
    transcript = re.sub(r'\s+([.,!?;:])', r'\1', transcript)
    transcript = re.sub(r'([.,!?;:])([^\s”])', r'\1 \2', transcript)
    transcript = re.sub(r"\n{3,}", "\n\n", transcript)
    transcript = re.sub(r'([.,!?;:])\s*\1+', r'\1', transcript)

    return transcript.strip()

# ========== AUDIO FUNCTIONS ==========

def record_audio(filename=AUDIO_FILENAME, fs=44100):
    print("🎙️ Recording... Press ENTER to stop.")
    recording = []
    is_recording = True

    def stop_on_enter():
        nonlocal is_recording
        input()
        is_recording = False

    listener = threading.Thread(target=stop_on_enter)
    listener.start()

    def callback(indata, frames, time, status):
        if status:
            print("⚠️", status)
        if is_recording:
            recording.append(indata.copy())
        else:
            raise sd.CallbackStop

    with sd.InputStream(samplerate=fs, channels=1, callback=callback):
        while is_recording:
            sd.sleep(100)

    audio = np.concatenate(recording, axis=0)
    audio = audio / np.max(np.abs(audio))
    wav.write(filename, fs, (audio * 32767).astype(np.int16))
    print(f"✅ Saved audio to {filename}")

# ========== TRANSCRIPTION ==========

def transcribe_audio(model_name=MODEL, audio_file=AUDIO_FILENAME):
    print("🔍 Transcribing audio with Whisper...")

    # Check that the file exists and is not empty
    if not os.path.exists(audio_file) or os.path.getsize(audio_file) == 0:
        raise RuntimeError("❌ Audio file not found or is empty. Please record again.")

    model = whisper.load_model(model_name)
    result = model.transcribe(audio_file)

    transcript = result["text"].strip()
    transcript = apply_voice_commands(transcript)

    with open(TRANSCRIPT_FILENAME, "w", encoding="utf-8") as f:
        f.write(transcript)

    print(f"📝 Transcription saved to {TRANSCRIPT_FILENAME}")
    return transcript

# ========== CLEANING WITH TOGETHER.AI ==========

def clean_with_together(transcript):
    print("🧠 Sending to Together.ai for grammar cleanup...")
    api_key = "6c86d1e21783273b45c286dad76384487fff5a99f0cfe339de63f0ecd8bb765c"
    if not api_key:
        raise ValueError("❌ TOGETHER_API_KEY not set.")

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": TOGETHER_MODEL,
        "messages": [
            {
                "role": "user",
                "content": "\n".join(PROMPT_INSTRUCTIONS) + f"\n\nTranscript:\n{transcript}"
            }
        ],
        "temperature": 0.3,
        "top_p": 0.9,
        "max_tokens": 1024
    }

    response = requests.post(
        "https://api.together.xyz/v1/chat/completions",
        headers=headers,
        json=payload
    )

    if response.status_code != 200:
        raise RuntimeError(f"❌ Together.ai API error:\n{response.text}")

    try:
        cleaned_text = response.json()["choices"][0]["message"]["content"].strip()
    except (KeyError, IndexError):
        raise RuntimeError("❌ Unexpected response format from Together.ai")

    print("✅ Cleaned Transcript:\n", cleaned_text)
    return cleaned_text

# ========== TEST SELECTION ==========

def select_tests():
    tests = [
        "Wechsler Abbreviated Scale of Intelligence-II (WASI-II)",
        "Portion of Wechsler Adult Intelligence Scale - Revised (WAIS-R)",
        "Portion of Wechsler Intelligence Scale for Children (WISC)",
        "Trail Making Test (Part B)",
        "Letter-Number Sequencing",
        "BAARS-IV",
        "Personality Assessment Inventory (PAI)"
    ]

    print("\n🧪 Tests Administered (all selected by default):")
    for i, test in enumerate(tests, 1):
        print(f"{i}. {test}")

    remove_input = input("\nEnter the number(s) of tests to REMOVE, separated by commas (or press ENTER to keep all): ")

    if remove_input.strip():
        try:
            remove_indices = [int(i.strip()) - 1 for i in remove_input.split(",") if i.strip().isdigit()]
            tests = [t for i, t in enumerate(tests) if i not in remove_indices]
        except ValueError:
            print("⚠️ Invalid input. Keeping all tests.")

    print("\n✅ Final Tests Selected:")
    for test in tests:
        print(f" - {test}")

    return tests

# ========== MAIN ==========

def main():
    patient = input("Enter PATIENT NAME: ")
    date = input("Enter DATE OF EVALUATION (MM/DD/YYYY): ")
    ordering_provider = input("Ordered by: ")

    reason = "Assist in evaluation of ADHD and differential diagnosis"

    selected_tests = select_tests()

    input("Press Enter to dictate Clinical Interview...")
    record_audio()
    clinical = transcribe_audio()
    clinical_cleaned = clean_with_together(clinical)

    input("Press Enter to dictate Consultation Findings...")
    record_audio()
    consultation = transcribe_audio()
    consultation_cleaned = clean_with_together(consultation)

    output_file = f"{patient.replace(' ', '_').lower()}_report.docx"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(f"PATIENT: {patient}\nDATE: {date}\n\n")
        f.write(f"ORDERED BY: {ordering_provider}\n\n")
        f.write(f"REASON FOR CONSULT:\n{reason}\n\n")
        f.write("TESTS ADMINISTERED:\n")
        for test in selected_tests:
            f.write(f"{test}\n")
        f.write(f"\nCLINICAL INTERVIEW:\n{clinical_cleaned}\n\n")
        f.write(f"CONSULTATION FINDINGS:\n{consultation_cleaned}\n")

    print(f"📄 Final report saved to: {output_file}")

if __name__ == "__main__":
    main()
