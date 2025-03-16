from docx import Document
from gtts import gTTS
from google.colab import files
import os

def read_docx(file_path):
    """Read text from a Word document (.docx) file."""
    doc = Document(file_path)
    full_text = [paragraph.text for paragraph in doc.paragraphs]
    return '\n'.join(full_text)

def text_to_audio(text, output_file):
    """Convert text to audio using Google Text-to-Speech."""
    tts = gTTS(text=text, lang='en')
    tts.save(output_file)
    print(f"Audio saved as {output_file}")

if __name__ == "__main__":
    # Step 1: Upload the Word document
    uploaded_files = files.upload()  # Upload files in Colab
    for filename in uploaded_files:
        local_file_path = filename  # Filename is the key in the returned dictionary
        print(f"Uploaded file: {local_file_path}")

        # Step 2: Read the Word document
        try:
            text = read_docx(local_file_path)
            if not text.strip():
                print("The document is empty or contains no readable text.")
            else:
                # Step 3: Convert text to audio
                audio_file = os.path.splitext(local_file_path)[0] + ".mp3"  # Dynamic output file name
                text_to_audio(text, audio_file)
        except Exception as e:
            print(f"Error processing {local_file_path}: {e}")
