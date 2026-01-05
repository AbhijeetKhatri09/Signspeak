# !/usr/bin/env python3

"""
Sign-to-Speech (simple: single debounce, GUI capture prompt before capture)
 - Press 'c' to enter label (dialog) and then capture template frames
 - Press 'm' to map label -> phrase (dialog)
 - Press 'v' to view templates and phrase-map entries in console
 - Press 't' to test TTS, 'q' to quit

ADDED: prints Braille (Unicode Braille Patterns) for any spoken output (spelled words or phrases).
"""

import time
import collections
import threading
import queue
import tempfile
import os
import sys
import argparse
import atexit
import json

import cv2
import numpy as np
import mediapipe as mp

# Try importing tkinter for GUI dialogs
try:
    import tkinter as tk
    from tkinter import simpledialog
    TK_AVAILABLE = True
except Exception:
    TK_AVAILABLE = False

# --------- CLI args (detect IPython) ---------
parser = argparse.ArgumentParser(description='Sign-to-Speech (simple debounced)')
parser.add_argument('--camera', type=int, default=0, help='Camera index (default 0)')
parser.add_argument('--templates', type=str, default='templates.npz', help='Path to templates.npz (optional)')
parser.add_argument('--phrase-map', type=str, default='label_to_phrase_small.json', help='JSON mapping file')
parser.add_argument('--debounce', type=float, default=3.0, help='Debounce between spoken phrases (sec). Default 3s.')
parser.add_argument('--spelled-pause', type=float, default=1.2, help='Seconds of silence to flush spelled buffer')
parser.add_argument('--thresh', type=float, default=0.25, help='Template match threshold (lower = stricter)')
parser.add_argument('--capture-frames', type=int, default=12, help='Frames to capture when recording a template')

def running_in_ipython():
    try:
        import ipykernel  # type: ignore
        return True
    except Exception:
        return False

if running_in_ipython():
    args = parser.parse_args([])   # use defaults in notebooks
else:
    args = parser.parse_args()

TEMPLATES_PATH = args.templates
PHRASE_MAP_FILE = args.phrase_map
CAPTURE_FRAMES = args.capture_frames

# load phrase map
try:
    phrase_map = json.load(open(PHRASE_MAP_FILE, 'r', encoding='utf-8'))
    print(f'Loaded phrase map from {PHRASE_MAP_FILE} ({len(phrase_map)} entries)', flush=True)
except Exception:
    phrase_map = {}
    print(f'No phrase map found at {PHRASE_MAP_FILE}, starting empty.', flush=True)

# ---------- Helpers for templates and phrase map ----------
def load_templates(path=TEMPLATES_PATH):
    if not path or not os.path.exists(path):
        return {}
    try:
        data = np.load(path, allow_pickle=True)
        templates = {}
        for k in data.files:
            v = np.array(data[k]).astype(np.float32)
            n = np.linalg.norm(v)
            if n > 0:
                templates[k] = v / (n + 1e-9)
            else:
                templates[k] = v
        print(f'Loaded {len(templates)} templates from {path}', flush=True)
        return templates
    except Exception as e:
        print('Failed to load templates:', e, flush=True)
        return {}

def save_templates(templates, path=TEMPLATES_PATH):
    try:
        # ensure path dir exists
        d = os.path.dirname(os.path.abspath(path)) or '.'
        os.makedirs(d, exist_ok=True)
        np.savez_compressed(path, **templates)
        print(f'Saved {len(templates)} templates to {path}', flush=True)
    except Exception as e:
        print('Failed to save templates:', e, flush=True)

def _normalize_map_key(key):
    """Normalized key used in phrase_map storage (lower, no spaces, no underscores)."""
    if key is None:
        return None
    k = str(key).lower().strip()
    k = k.replace(' ', '').replace('_', '').replace('-', '')
    return k

def update_phrase_map(label, phrase, path=PHRASE_MAP_FILE):
    try:
        k = _normalize_map_key(label) or str(label)
        phrase_map[k] = phrase
        # keep a human-friendly copy with original label too (optional)
        # also write the un-normalized label if you want; for now store normalized keys
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(phrase_map, f, ensure_ascii=False, indent=2)
        print(f'Updated phrase_map: "{label}" -> "{phrase}" in {path}', flush=True)
    except Exception as e:
        print('Failed to update phrase map:', e, flush=True)

# ---------- TTS/backends detection ----------
WIN32COM_OK = False
PYTTSX3_OK = False
GTTS_OK = False
_PLAYSOUND_OK = False
try:
    import win32com.client  # type: ignore
    WIN32COM_OK = True
except Exception:
    WIN32COM_OK = False
try:
    import pyttsx3  # type: ignore
    PYTTSX3_OK = True
except Exception:
    PYTTSX3_OK = False
try:
    from gtts import gTTS  # type: ignore
    from playsound import playsound  # type: ignore
    GTTS_OK = True
    _PLAYSOUND_OK = True
except Exception:
    GTTS_OK = False

_tts_queue = queue.Queue()
_worker_thread = None
_worker_started = threading.Event()
_worker_stop = threading.Event()

def _select_female_sapi_voice(sapi):
    try:
        voices = sapi.GetVoices()
        for v in voices:
            try:
                gender = v.GetAttribute('Gender').lower()
            except Exception:
                gender = ''
            try:
                name = v.GetAttribute('Name').lower()
            except Exception:
                name = ''
            if 'female' in gender or any(x in name for x in ('zira','heera','eva','girl')):
                sapi.Voice = v
                return True
    except Exception:
        pass
    return False

def _select_female_pyttsx3_voice(engine):
    try:
        voices = engine.getProperty('voices')
        for v in voices:
            name = getattr(v, 'name', '') or getattr(v, 'id', '')
            if any(x in str(name).lower() for x in ('female','zira','heera','eva','girl')):
                engine.setProperty('voice', v.id)
                return True
    except Exception:
        pass
    return False

def _tts_worker():
    sapi = None
    py_engine = None
    if WIN32COM_OK:
        try:
            sapi = win32com.client.Dispatch('SAPI.SpVoice')  # type: ignore
            try:
                _select_female_sapi_voice(sapi)
            except Exception:
                pass
            try:
                sapi.Rate = -1
            except Exception:
                pass
        except Exception:
            sapi = None
    if sapi is None and PYTTSX3_OK:
        try:
            py_engine = pyttsx3.init()
            try:
                _select_female_pyttsx3_voice(py_engine)
            except Exception:
                pass
        except Exception:
            py_engine = None
    _worker_started.set()
    while not _worker_stop.is_set():
        try:
            text = _tts_queue.get(timeout=0.2)
        except queue.Empty:
            continue
        if text is None:
            break
        spoke = False
        if sapi is not None:
            try:
                sapi.Speak(text)
                spoke = True
            except Exception:
                spoke = False
        if not spoke and py_engine is not None:
            try:
                py_engine.say(text)
                py_engine.runAndWait()
                spoke = True
            except Exception:
                spoke = False
        if not spoke and GTTS_OK and _PLAYSOUND_OK:
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.mp3') as tf:
                    tmp = tf.name
                gTTS(text=text, lang='en').save(tmp)
                playsound(tmp)
                try:
                    os.remove(tmp)
                except Exception:
                    pass
                spoke = True
            except Exception:
                spoke = False
        _tts_queue.task_done()
    try:
        if py_engine is not None:
            py_engine.stop()
    except Exception:
        pass

def start_worker():
    global _worker_thread
    if _worker_thread is None or not _worker_thread.is_alive():
        _worker_thread = threading.Thread(target=_tts_worker, daemon=True)
        _worker_thread.start()
        _worker_started.wait(timeout=2.0)

def stop_worker():
    _worker_stop.set()
    try:
        _tts_queue.put_nowait(None)
    except Exception:
        pass
    try:
        if _worker_thread is not None:
            _worker_thread.join(timeout=1.0)
    except Exception:
        pass

def speak(text):
    start_worker()
    try:
        _tts_queue.put_nowait(text)
    except Exception:
        _tts_queue.put(text)

atexit.register(stop_worker)

# ---------- Gesture helpers ----------
mp_hands = mp.solutions.hands
mp_drawing = mp.solutions.drawing_utils
FINGER_TIPS = [4,8,12,16,20]

def fingers_up(hand_landmarks, handed_label='Right'):
    lm = hand_landmarks.landmark
    res = []
    try:
        if handed_label and handed_label.lower().startswith('r'):
            res.append(lm[4].x < lm[3].x)
        else:
            res.append(lm[4].x > lm[3].x)
    except Exception:
        res.append(False)
    for tip,pip in zip([8,12,16,20],[6,10,14,18]):
        try:
            res.append(lm[tip].y < lm[pip].y)
        except Exception:
            res.append(False)
    return res

def match_pattern(fs):
    if not fs:
        return None
    t,i,m,r,p = fs
    if t and i and (not m) and (not r) and p:
        return 'I LOVE YOU'
    if all(fs):
        return 'HELLO'
    if t and (not i) and (not m) and (not r) and (not p):
        return 'THUMBS UP'
    if (not t) and i and m and (not r) and (not p):
        return 'PEACE'
    if not any(fs):
        return 'FIST'
    return None

def landmarks_to_vector(hand_landmarks):
    lm = hand_landmarks.landmark
    try:
        bx = lm[0].x; by = lm[0].y; bz = lm[0].z
    except Exception:
        bx = by = bz = 0.0
    vec = []
    for i in range(21):
        try:
            vec.extend([lm[i].x - bx, lm[i].y - by, lm[i].z - bz])
        except Exception:
            vec.extend([0.0,0.0,0.0])
    return np.array(vec, dtype=np.float32)

# ---------- Template matching ----------
templates = load_templates(TEMPLATES_PATH)

def match_template(vec, templates_dict):
    if not templates_dict:
        return None, None
    v = vec.astype(np.float32)
    n = np.linalg.norm(v)
    if n == 0:
        return None, None
    v = v / (n + 1e-9)
    best_label, best_score = None, float('inf')
    for lbl,t in templates_dict.items():
        score = 1.0 - float(np.dot(v, t))
        if score < best_score:
            best_label, best_score = lbl, score
    return best_label, best_score

# ---------- capture using existing cap (prompt before capture) ----------
def prompt_text(title, prompt):
    """Return input text using tkinter dialog if available, else terminal input()."""
    if TK_AVAILABLE:
        try:
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            answer = simpledialog.askstring(title, prompt, parent=root)
            root.destroy()
            return answer
        except Exception:
            # fallback to console input
            try:
                return input(f'{prompt} ')
            except Exception:
                return None
    else:
        try:
            return input(f'{prompt} ')
        except Exception:
            return None

def capture_template_for_label_from_cap(cap, label, frames_to_capture=CAPTURE_FRAMES, timeout=6.0):
    """
    Use the existing VideoCapture cap to capture several good frames and average to a template.
    Returns True on success and updates global templates + disk.
    """
    global templates
    if not cap or not cap.isOpened():
        print('Camera not available for capture.', flush=True)
        return False
    print(f'Capturing template for label "{label}" — hold the sign steady...', flush=True)
    collected = []
    t0 = time.time()
    with mp_hands.Hands(max_num_hands=1, min_detection_confidence=0.6) as hands_local:
        while len(collected) < frames_to_capture and (time.time() - t0) < timeout:
            ok, frame = cap.read()
            if not ok or frame is None:
                time.sleep(0.02)
                continue
            frame = cv2.flip(frame, 1)
            rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            res = hands_local.process(rgb)
            if res and hasattr(res, 'multi_hand_landmarks') and res.multi_hand_landmarks:
                lm = res.multi_hand_landmarks[0]
                vec = landmarks_to_vector(lm)
                n = np.linalg.norm(vec)
                if n > 1e-6:
                    vec = vec / (n + 1e-9)
                collected.append(vec)
            # show capture progress overlay
            overlay = frame.copy()
            cv2.putText(overlay, f'Capturing {len(collected)}/{frames_to_capture}', (10,40),
                        cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0,255,255), 2)
            cv2.imshow('Sign-to-Speech (press q to quit)', overlay)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

    if len(collected) < max(3, frames_to_capture // 2):
        print(f'Not enough good frames collected ({len(collected)}), aborting capture.', flush=True)
        return False
    mean_vec = np.mean(np.stack(collected, axis=0), axis=0).astype(np.float32)
    n = np.linalg.norm(mean_vec)
    if n > 0:
        mean_vec = mean_vec / (n + 1e-9)
    templates[label] = mean_vec
    save_templates(templates, TEMPLATES_PATH)
    print(f'Captured template "{label}" with {len(collected)} frames.', flush=True)
    return True

# ---------- BRAILLE: mapping and helper ----------
_braille_map = {
    'a': '\u2801', 'b': '\u2803', 'c': '\u2809', 'd': '\u2819', 'e': '\u2811',
    'f': '\u280B', 'g': '\u281B', 'h': '\u2813', 'i': '\u280A', 'j': '\u281A',
    'k': '\u2805', 'l': '\u2807', 'm': '\u280D', 'n': '\u281D', 'o': '\u2815',
    'p': '\u280F', 'q': '\u281F', 'r': '\u2817', 's': '\u280E', 't': '\u281E',
    'u': '\u2825', 'v': '\u2827', 'w': '\u283A', 'x': '\u282D', 'y': '\u283D',
    'z': '\u2835',
    '0': '\u281A', '1': '\u2801', '2': '\u2803', '3': '\u2809', '4': '\u2819',
    '5': '\u2811', '6': '\u280B', '7': '\u281B', '8': '\u2813', '9': '\u280A',
    ' ': '\u2800',
    '.': '\u2832', ',': '\u2802', '?': '\u2822', '!': '\u2816', "'": '\u2804',
    '-': '\u2824', '/': '\u282C', ':': '\u2812', ';': '\u2836', '@': '\u2800'
}

def text_to_braille(text):
    """Convert ascii text to a string of Unicode Braille characters (grade-1 style)."""
    if not text:
        return ''
    out = []
    for ch in str(text):
        lower = ch.lower()
        if lower in _braille_map:
            out.append(_braille_map[lower])
        else:
            out.append('\u2800')
    return ''.join(out)

# ---------- helpers ----------
def normalize_lookup_key(raw_label):
    if raw_label is None:
        return []
    s = str(raw_label)
    low = s.lower().strip()
    variants = []
    # include a handful of useful normalizations
    variants.append(low)
    variants.append(low.replace(' ', ''))
    variants.append(low.replace('_', ''))
    variants.append(low.replace('-', ''))
    variants.append(low.lstrip('0'))
    seen = set()
    out = []
    for v in variants:
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out

def lookup_phrase_for_label(raw_label):
    if raw_label is None:
        return ''
    # try normalized keys
    for v in normalize_lookup_key(raw_label):
        if v in phrase_map:
            return phrase_map[v]
    # direct fallback
    k = _normalize_map_key(raw_label)
    if k in phrase_map:
        return phrase_map[k]
    return str(raw_label).replace('_',' ').title()

# ---------- Main loop ----------
def main():
    global templates, phrase_map

    cap = cv2.VideoCapture(args.camera)
    if not cap.isOpened():
        print('Camera not available at index', args.camera, flush=True)
        return

    start_worker()

    last_time = 0.0
    DEBOUNCE = args.debounce
    preds = collections.deque(maxlen=7)

    templates = load_templates(args.templates) if args.templates else {}

    spelled_buffer = []
    SPELLED_MAX = 20
    SPELLED_PAUSE = args.spelled_pause
    _last_letter_time = 0.0

    print('--- Sign to Speech Started (one output per {:.1f}s) ---'.format(DEBOUNCE), flush=True)
    print("Interactive keys: 'c' = capture template, 'm' = map phrase, 'v' = view templates/phrases, 't' = test TTS, 'q' = quit", flush=True)

    with mp_hands.Hands(max_num_hands=1, min_detection_confidence=0.6) as hands:
        while True:
            ok, frame = cap.read()
            if not ok or frame is None:
                break

            frame = cv2.flip(frame, 1)
            h, w, _ = frame.shape
            rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            res = hands.process(rgb)

            sign_text = ''
            detected_label = None
            detected_score = None

            if res.multi_hand_landmarks:
                handedness_list = getattr(res, 'multi_handedness', None)

                for idx, hand_landmarks in enumerate(res.multi_hand_landmarks):
                    handed_label = 'Right'
                    if handedness_list and idx < len(handedness_list):
                        try:
                            handed_label = handedness_list[idx].classification[0].label
                        except Exception:
                            pass

                    mp_drawing.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)

                    if templates:
                        vec = landmarks_to_vector(hand_landmarks)
                        detected_label, detected_score = match_template(vec, templates)

                    fs = fingers_up(hand_landmarks, handed_label)
                    rule_match = match_pattern(fs)

                    if detected_label is None and rule_match:
                        preds.append(rule_match)
                    else:
                        if detected_label is not None and detected_score is not None and detected_score < args.thresh:
                            preds.append(detected_label)
                        else:
                            preds.append(None)
            else:
                preds.append(None)

            votes = [p for p in preds if p]
            sign_text = max(set(votes), key=votes.count) if votes else ''

            now = time.time()

            if sign_text:
                cv2.putText(frame, f'Sign: {sign_text}', (10, 30),
                            cv2.FONT_HERSHEY_DUPLEX, 1.0, (0, 255, 0), 2)

                # LETTER ACCUMULATION
                if len(sign_text) == 1 and sign_text.isalpha():
                    if now - _last_letter_time > 0.45:
                        spelled_buffer.append(sign_text)
                        if len(spelled_buffer) > SPELLED_MAX:
                            spelled_buffer.pop(0)
                        _last_letter_time = now

                # FLUSH SPELLED WORD
                if spelled_buffer and (now - _last_letter_time) > SPELLED_PAUSE:
                    word = ''.join(spelled_buffer)
                    phrase = lookup_phrase_for_label(word)

                    print("Spoken:", phrase)
                    print("Braille:", text_to_braille(phrase))

                    speak(phrase)
                    spelled_buffer.clear()
                    last_time = now

                # SPEAK FULL SIGN
                elif (now - last_time) > DEBOUNCE:
                    phrase = lookup_phrase_for_label(sign_text)

                    print("Spoken:", phrase)
                    print("Braille:", text_to_braille(phrase))

                    speak(phrase)
                    last_time = now

            cv2.imshow('Sign-to-Speech (press q to quit)', frame)

            key = cv2.waitKey(1) & 0xFF
            if key == ord('q'):
                break

            elif key == ord('t'):
                speak("Text to speech is working")

            elif key == ord('v'):
                print("Templates:", list(templates.keys()))
                print("Phrase Map:", phrase_map)

            elif key == ord('m'):
                label = prompt_text("Map Label", "Enter label:")
                phrase = prompt_text("Map Phrase", "Enter phrase:")
                if label and phrase:
                    update_phrase_map(label, phrase)

            elif key == ord('c'):
                label = prompt_text("Capture Template", "Enter label:")
                if label:
                    capture_template_for_label_from_cap(cap, label)

    cap.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    main()