# Sign-to-Speech: AI Powered Sign Language to Voice and Braille Converter

Sign-to-Speech is an intelligent multimodal system designed to bridge communication gaps for people with hearing and speech impairments. By leveraging AI-based computer vision, the system converts real-time hand gestures into clear voice output and Braille representations using a standard webcam.

<img width="806" height="662" alt="image" src="https://github.com/user-attachments/assets/eedf2bb4-b1cc-4117-b397-20bd268888c9" />


---

## 🚀 Project Purpose
The purpose of this project is to design and implement a vision-based interpretation system that:
* **Recognizes Hand Signs:** Detects static gestures in real-time using a standard webcam.
* **Voice Synthesis:** Converts recognized signs into meaningful text and natural speech.
* **Tactile Accessibility:** Generates equivalent Braille output using Unicode patterns for visually impaired users.

---

## 🏗️ System Architecture
The platform is built on three integrated interface layers:

* **Webcam Interface:** Captures live video frames via OpenCV to track hand movements.
* **Processing Layer:** Uses **Mediapipe Hands** to detect 21 hand landmarks and represent them as numeric feature vectors.
* **System Interface:** Converts recognized labels to speech via TTS engines and simultaneously generates Unicode Braille patterns.

<img width="400" height="300" alt="image" src="https://github.com/user-attachments/assets/39e92f89-795e-4b99-a4e6-16809417eabf" />


---

## 💻 Technical Stack
* **Language:** Python
* **Computer Vision:** OpenCV & MediaPipe Hands
* **Audio Engines:** pyttsx3, SAPI, or gTTS
* **Data Handling:** JSON-based phrase mapping and NumPy for template matching.

---

## ✨ Features and Functionalities
* **Gesture Templates:** Users can record and store custom sign "templates" for alphabets or phrases.
* **Interactive Command Keys:**
    * `c` - Capture a new sign template.
    * `m` - Map a label to a specific phrase.
    * `v` - View existing mappings and templates.
    * `t` - Test the TTS engine output.
* **Braille Output:** Automatically generates Unicode Braille for any spoken output (e.g., ⠠⠓⠑⠇⠇⠓).

<img width="500" height="624" alt="image" src="https://github.com/user-attachments/assets/2871de25-8b50-4736-98b6-e3bc1f8d5e09" />


---

## 🏁 Conclusion
In conclusion, SignSpeak demonstrates the power of Human-Computer Interaction (HCI). By integrating computer vision with voice and tactile outputs, the project provides a low-cost solution for inclusive communication between sign language users and the general public.
