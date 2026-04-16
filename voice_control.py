"""
Aero Present - Voice Control Module
Handles voice commands for slide navigation

Uses hybrid approach:
- Google Speech API (when online)
- Vosk (offline fallback)
"""

import speech_recognition as sr
import json, threading, time, re
from queue import Queue

# optional dependencies
try:
    import vosk, pyaudio
    VOSK_AVAILABLE = True
except ImportError:
    VOSK_AVAILABLE = False
    print("[VOICE] Vosk not available, Google-only mode")

try:
    import pyttsx3
    TTS_AVAILABLE = True
except ImportError:
    TTS_AVAILABLE = False

class VoiceController:
    """Voice command processor with semantic navigation"""
    
    def __init__(self, mobile_backend=None, keyword_map=None):
        self.backend = mobile_backend
        self.keyword_map = keyword_map or {}
        
        # recognizers
        self.recognizer = sr.Recognizer()
        self.vosk_rec = None
        
        # settings
        self.enabled = True
        self.listening = False
        self.use_tts = TTS_AVAILABLE

        # Google Speech API failure tracking.
        # After _google_fail_limit consecutive failures, stop calling Google for
        # _google_backoff_secs seconds and use Vosk directly. Resets on any success.
        self._google_failures      = 0
        self._google_fail_limit    = 3
        self._google_backoff_secs  = 60
        self._google_disabled_until = 0.0
        # Say "AERO" to activate an 8-second listening window.
        # Only commands spoken within that window are processed.
        # Set wake_word_mode = False to disable and always listen.
        self.wake_word      = "aero"
        self.wake_word_mode = True
        self._wake_active   = False
        self._wake_expires  = 0.0
        # 8s window: listen_google blocks up to 5s + ~1s TTS "Ready" + 2s margin
        self.wake_window    = 8.0   # seconds to accept commands after wake word
        
        # init vosk if available
        self._pyaudio = None   # initialised inside _init_vosk when Vosk loads
        if VOSK_AVAILABLE:
            self._init_vosk()

        # one-time microphone noise calibration (avoids 0.5s overhead every listen call)
        self._calibrate_microphone()

        # shared Microphone source — opened once at startup, reused by listen_google
        # opening it per-call causes ~100ms device re-enumeration overhead on Windows
        try:
            self._mic = sr.Microphone()
        except Exception as e:
            self._mic = None
            print(f"[VOICE] Microphone unavailable: {e}")
        
        # text-to-speech — runs on a dedicated thread to avoid blocking the listen loop
        # pyttsx3.runAndWait() is not safe to call from an arbitrary daemon thread on Windows
        self._tts_queue  = Queue()
        if TTS_AVAILABLE:
            self.tts = pyttsx3.init()
            self.tts.setProperty('rate', 175)
            self.tts.setProperty('volume', 0.9)
            tts_thread = threading.Thread(target=self._tts_loop, daemon=True)
            tts_thread.start()
        
        print("[VOICE] Initialized")
    
    def _init_vosk(self):
        """Load vosk model and initialise PyAudio once for the session."""
        try:
            model_path = "models/vosk-model-small-en-us-0.15"
            self.vosk_model = vosk.Model(model_path)
            self.vosk_rec = vosk.KaldiRecognizer(self.vosk_model, 16000)
            # Initialise PyAudio once here — creating it per-call caused
            # device re-enumeration overhead (~100ms) and audible pops on Windows
            self._pyaudio = pyaudio.PyAudio()
            print("[VOICE] Vosk model loaded")
        except Exception as e:
            self._pyaudio = None
            print(f"[VOICE] Vosk failed: {e}")
    
    def _calibrate_microphone(self):
        """One-time noise calibration at startup — not repeated every listen call."""
        try:
            with sr.Microphone() as source:
                print("[VOICE] Calibrating microphone for ambient noise...")
                self.recognizer.adjust_for_ambient_noise(source, duration=1.0)
                print("[VOICE] Microphone ready")
        except Exception as e:
            print(f"[VOICE] Mic calibration failed: {e}")

    def listen_google(self, timeout=5):
        """Use free Google Web Speech API.
        Raises sr.RequestError on API failure so listen() can count it.
        Returns None on silence or unrecognised speech — those are NOT failures."""
        if not self._mic:
            return None
        try:
            with self._mic as source:
                audio = self.recognizer.listen(source, timeout=timeout, phrase_time_limit=5)
                text = self.recognizer.recognize_google(audio)
                return text.lower()
        except sr.WaitTimeoutError:
            return None       # silence — not an API failure, don't count it
        except sr.UnknownValueError:
            return None       # speech heard but not recognised — not an API failure
        except sr.RequestError:
            raise             # re-raise so listen() can count a real network failure

    def listen(self):
        """Smart listening — Google first with failure backoff, then Vosk fallback.
        Only genuine sr.RequestError counts toward the circuit breaker.
        Silence and unrecognised speech do NOT trip it."""
        if not self.enabled:
            return None

        google_available = time.time() >= self._google_disabled_until
        text = None

        if google_available:
            try:
                text = self.listen_google()
                # Returned without exception — API is reachable, reset failure counter
                self._google_failures = 0
            except sr.RequestError:
                self._google_failures += 1
                if self._google_failures >= self._google_fail_limit:
                    self._google_disabled_until = time.time() + self._google_backoff_secs
                    print(f"[VOICE] Google API unreliable — switching to Vosk for "
                          f"{self._google_backoff_secs}s")
                    self._google_failures = 0

        # Vosk fallback — used when Google failed or is in backoff
        if text is None and VOSK_AVAILABLE:
            text = self.listen_vosk()

        return text

    def listen_vosk(self):
        """Offline recognition with Vosk — uses shared PyAudio instance."""
        if not VOSK_AVAILABLE or not self.vosk_rec:
            return None
        if not getattr(self, '_pyaudio', None):
            return None

        stream = None
        try:
            stream = self._pyaudio.open(
                format=pyaudio.paInt16,
                channels=1,
                rate=16000,
                input=True,
                frames_per_buffer=8000
            )

            for _ in range(int(16000 / 8000 * 5)):
                data = stream.read(8000, exception_on_overflow=False)
                if self.vosk_rec.AcceptWaveform(data):
                    result = json.loads(self.vosk_rec.Result())
                    text = result.get('text', '')
                    if text:
                        return text.lower()

        except Exception as e:
            print(f"[VOICE] Vosk error: {e}")

        finally:
            # Only close the stream — the PyAudio instance is reused across calls
            try:
                if stream:
                    stream.stop_stream()
                    stream.close()
            except Exception:
                pass

        return None
    
    def process_command(self, text):
        """Process voice command"""
        if not text:
            return False
        
        text = text.lower().strip()
        
        # navigation
        if self._handle_navigation(text):
            return True
        
        # semantic nav (keyword search)
        if self._handle_semantic(text):
            return True
        
        # laser
        if self._handle_laser(text):
            return True
        
        # system
        if self._handle_system(text):
            return True
        
        return False
    
    def _handle_navigation(self, text):
        """Basic navigation commands — all triggers use word boundaries or full phrases
        to avoid accidental fires mid-sentence ('backhand', 'start of story', etc.)."""

        # next — 'next' and 'forward' are safe as-is (rarely appear mid-sentence harmfully)
        if re.search(r'\bnext\b', text) or re.search(r'\bforward\b', text):
            if self.backend:
                self.backend.execute_command("NEXT")
            self.speak("Next")
            return True

        # previous — 'previous' and 'backward' are safe; 'back' must be a standalone word
        # to avoid 'backhand', 'setback', 'feedback' etc. triggering it
        if (re.search(r'\bprevious\b', text) or re.search(r'\bbackward\b', text) or
                re.search(r'\bgo back\b', text)):
            if self.backend:
                self.backend.execute_command("PREVIOUS")
            self.speak("Previous")
            return True

        # first slide — require full phrase; 'start' alone is too common in speech
        if 'first slide' in text or re.search(r'\bbeginning\b', text):
            if self.backend:
                self.backend.execute_command("JUMP", {'slide_number': 1})
            self.speak("First slide")
            return True

        # last slide — require full phrase; 'end' alone fires on "in the end", "end result"
        if 'last slide' in text or re.search(r'\bfinal slide\b', text):
            if self.backend:
                total = self.backend.system_status.get('total_slides', 0)
                if total > 0:
                    self.backend.execute_command("JUMP", {'slide_number': total})
            self.speak("Last slide")   # always confirm — even if PPT not open
            return True

        # go to slide X — 'slide' must be present with a number
        if re.search(r'\bslide\b', text):
            num = self._extract_number(text)
            if num:
                if self.backend:
                    self.backend.execute_command("JUMP", {'slide_number': num})
                self.speak(f"Slide {num}")
                return True

        return False
    
    def _handle_semantic(self, text):
        """Semantic navigation using keywords — always reads live from backend if available."""
        # Use backend's live keyword map if connected, otherwise fall back to local copy.
        # This ensures keywords added via phone after startup are visible immediately.
        keyword_map = (self.backend.slide_keywords
                       if self.backend and hasattr(self.backend, 'slide_keywords')
                       else self.keyword_map)

        if not keyword_map:
            return False

        # check for trigger words
        triggers = ['go to', 'show', 'find', 'jump to', 'open']
        if not any(t in text for t in triggers):
            return False

        # remove triggers
        for trigger in triggers:
            text = text.replace(trigger, '').strip()

        # search keywords
        for slide_num, keywords in keyword_map.items():
            for keyword in keywords:
                if keyword.lower() in text:
                    if self.backend:
                        self.backend.execute_command("JUMP", {'slide_number': slide_num})
                    self.speak(f"Going to {keyword}")
                    return True

        return False
    
    def _handle_laser(self, text):
        """Laser pointer control — routes through execute_command so gesture_control picks it up."""
        if 'laser' not in text:
            return False

        if 'on' in text or 'show' in text:
            if self.backend:
                # Only toggle if currently off — check current state first
                if not self.backend.system_status.get('laser_enabled', False):  # Bug 13 fix: was True — if key absent assumed laser ON, preventing toggle
                    self.backend.execute_command("LASER_TOGGLE")
            self.speak("Laser on")
            return True

        if 'off' in text or 'hide' in text:
            if self.backend:
                # Only toggle if currently on
                if self.backend.system_status.get('laser_enabled', False):
                    self.backend.execute_command("LASER_TOGGLE")
            self.speak("Laser off")
            return True

        return False
    
    def _handle_system(self, text):
        """System control commands — all toggles route through execute_command."""

        # voice control (self-managed, no backend needed)
        if 'disable voice' in text:
            self.enabled = False
            self.speak("Voice disabled")
            return True

        if 'enable voice' in text:
            self.enabled = True
            self.speak("Voice enabled")
            return True

        # gesture control — must go through pending_changes so gesture_control responds
        if 'gesture' in text:
            if 'disable' in text or 'off' in text:
                if self.backend:
                    # Only toggle if currently on
                    if self.backend.system_status.get('gesture_enabled', True):
                        self.backend.execute_command("GESTURE_TOGGLE")
                self.speak("Gestures disabled")
                return True

            if 'enable' in text or 'on' in text:
                if self.backend:
                    # Only toggle if currently off
                    if not self.backend.system_status.get('gesture_enabled', True):
                        self.backend.execute_command("GESTURE_TOGGLE")
                self.speak("Gestures enabled")
                return True

        return False
    
    def _extract_number(self, text):
        """Extract number from text using digit or word form."""
        # try digit sequences first
        matches = re.findall(r'\d+', text)
        if matches:
            return int(matches[0])

        # word-to-number with word boundaries to prevent partial matches
        # e.g. 'stone' contains 'one' but \bone\b won't match it
        numbers = {
            'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5,
            'six': 6, 'seven': 7, 'eight': 8, 'nine': 9, 'ten': 10,
            'eleven': 11, 'twelve': 12, 'thirteen': 13, 'fourteen': 14,
            'fifteen': 15, 'sixteen': 16, 'seventeen': 17, 'eighteen': 18,
            'nineteen': 19, 'twenty': 20,
        }
        for word, num in numbers.items():
            if re.search(rf'\b{word}\b', text):
                return num

        return None
    
    def _tts_loop(self):
        """Dedicated TTS thread — drains the queue and speaks each phrase.
        Keeps pyttsx3 on one consistent thread to avoid COM/Win32 issues."""
        while True:
            text = self._tts_queue.get()   # blocks until something is queued
            try:
                self.tts.say(text)
                self.tts.runAndWait()
            except Exception:
                pass
            finally:
                self._tts_queue.task_done()

    def speak(self, text):
        """Queue a TTS phrase — non-blocking, thread-safe."""
        if not self.use_tts or not TTS_AVAILABLE:
            return
        try:
            self._tts_queue.put_nowait(text)
        except Exception:
            pass

    @property
    def wake_word_active(self) -> bool:
        """
        True for wake_window seconds after the wake word 'AERO' is detected.
        Automatically expires and resets _wake_active when the window closes.
        """
        if self._wake_active and time.time() > self._wake_expires:
            self._wake_active = False
        return self._wake_active
    
    def start_listening(self):
        """Start continuous listening in background"""
        if self.listening:
            return
        
        self.listening = True
        self.thread = threading.Thread(target=self._listen_loop, daemon=True)
        self.thread.start()
        print("[VOICE] Started listening")
    
    def stop_listening(self):
        """Stop listening"""
        self.listening = False
    
    def _listen_loop(self):
        """
        Background listening loop — wake word aware.

        Flow:
          1. Listen for audio (Google → Vosk fallback).
          2. If 'AERO' is in the transcription, activate an 8-second window
             and give TTS confirmation 'Ready'.  If more words follow in the
             same utterance (e.g. 'Aero next'), they are processed immediately.
          3. Only process navigation/control commands while the wake window is
             open.  Commands heard outside the window are silently ignored.

        Set self.wake_word_mode = False to disable the wake requirement and
        always process every heard command (useful during solo development).
        """
        while self.listening:
            if not self.enabled:
                time.sleep(0.5)
                continue

            text = self.listen()

            if not text:
                time.sleep(0.1)
                continue

            text = text.lower().strip()

            # ── Wake word detection ─────────────────────────────────────────
            # Word boundary prevents 'aerodynamics', 'aerospace', 'area' from triggering
            if self.wake_word_mode and re.search(rf'\b{re.escape(self.wake_word)}\b', text):
                self._wake_active  = True
                self._wake_expires = time.time() + self.wake_window
                print(f"[VOICE] Wake word 'AERO' detected — {self.wake_window}s window open")
                self.speak("Ready")

                # Strip wake word and check if a command followed in same breath
                remainder = text.replace(self.wake_word, "").strip()
                if remainder:
                    self.process_command(remainder)

                time.sleep(0.1)
                continue

            # ── Command processing ──────────────────────────────────────────
            if not self.wake_word_mode or self.wake_word_active:
                self.process_command(text)
            # else: wake word required, not yet spoken — silently ignore

            time.sleep(0.1)

# ─────────────────────────────────────────────────────────────────────────────
# Testing
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # test keywords
    keywords = {
        1: ["intro", "title"],
        2: ["outline"],
        4: ["problem"],
        7: ["methodology", "backhand"],
        11: ["demo"],
        12: ["results"]
    }
    
    voice = VoiceController(mobile_backend=None, keyword_map=keywords)
    
    print("\nTry: next slide, go to slide 5, go to results, laser on")
    print("Ctrl+C to stop\n")
    
    while True:
        try:
            text = voice.listen()
            if text:
                print(f"Heard: {text}")
                voice.process_command(text)
        except KeyboardInterrupt:
            break
