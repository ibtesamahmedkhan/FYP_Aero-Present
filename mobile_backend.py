"""
Aero Present - Mobile Backend Server
Auto-detects IP, serves mobile app, handles all commands from phone.

Commands supported:
  NEXT, PREVIOUS, JUMP           — slide navigation
  NEXT_NOTIFY, PREV_NOTIFY       — gesture swipe fast-update (no COM nav)
  ZOOM_IN, ZOOM_OUT, ZOOM_RESET  — zoom overlay control
  GESTURE_TOGGLE, LASER_TOGGLE, VOICE_TOGGLE — feature toggles
  TILT_MOVE, TILT_STOP           — gyro pointer
  PPT_START, PPT_STOP, PPT_POINTER — slideshow control (F5 / Esc / Ctrl+P)
  ANNOTATION_TOGGLE, ANNOTATION_ERASE — draw mode and erase strokes
"""

import json, os, queue as _queue, re, socket, threading, time
from flask import Flask, jsonify, request, Response
from flask_socketio import SocketIO, emit

# ─────────────────────────────────────────────────────────────────────────────
# Session Security Token
# A fresh token is generated each session — printed to console and embedded
# in the QR code URL so only the presenter's phone connects.
# ─────────────────────────────────────────────────────────────────────────────
SESSION_TOKEN = os.urandom(16).hex()

try:
    import win32com.client
    PPT_AVAILABLE = True
except ImportError:
    PPT_AVAILABLE = False

changes_lock = threading.Lock()

# ─────────────────────────────────────────────────────────────────────────────
# Auto IP Detection
# ─────────────────────────────────────────────────────────────────────────────

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

# ─────────────────────────────────────────────────────────────────────────────
# PowerPoint Controller
# One dedicated COM thread owns the STA — all COM calls dispatched via queue.
# ─────────────────────────────────────────────────────────────────────────────

class PowerPointController:
    """
    Thread-safe PowerPoint COM controller.

    win32com.client.Dispatch creates a COM object on a Single-Threaded
    Apartment (STA). Any thread that calls COM methods without owning that
    apartment throws com_error. Fix: one '_com_worker' thread owns the STA
    forever. All operations are dispatched to it via a queue and awaited.
    """

    _STOP = object()

    def __init__(self):
        self._ppt  = None
        self._pres = None
        self._show = None
        self._q    = _queue.Queue()

        if PPT_AVAILABLE:
            t = threading.Thread(target=self._com_worker, daemon=True,
                                 name='ppt-com-thread')
            t.start()

    # ── COM worker ─────────────────────────────────────────────────────────

    def _com_worker(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except ImportError:
            pass

        if PPT_AVAILABLE:
            try:
                self._ppt = win32com.client.Dispatch("PowerPoint.Application")
                print("[PPT] Connected to PowerPoint")
            except Exception as e:
                print(f"[PPT] Connect failed: {e}")

        while True:
            try:
                op, arg, evt, holder = self._q.get(timeout=0.5)
            except _queue.Empty:
                continue

            if op is self._STOP:
                break

            try:
                if   op == 'get_info':       holder[0] = self._do_get_info()
                elif op == 'get_all_titles': holder[0] = self._do_get_all_titles()
                elif op == 'next':           holder[0] = self._do_next()
                elif op == 'prev':           holder[0] = self._do_prev()
                elif op == 'jump':           holder[0] = self._do_jump(arg)
            except Exception:
                holder[0] = self._error_state() if op == 'get_info' else False
            finally:
                evt.set()

        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except ImportError:
            pass

    def _dispatch(self, op, arg=None, timeout=2.0):
        if not PPT_AVAILABLE:
            return None
        holder = [None]
        evt    = threading.Event()
        self._q.put((op, arg, evt, holder))
        evt.wait(timeout=timeout)
        return holder[0]

    # ── Public API — safe to call from any thread ──────────────────────────

    def get_info(self):
        if not PPT_AVAILABLE:
            return self._no_ppt_state()
        result = self._dispatch('get_info')
        if result is None:
            return {**self._no_ppt_state(), 'ppt_message': 'Timeout — PPT busy'}
        return result

    def get_all_slide_titles(self):
        """Returns {slide_number: title_string} for every slide. Used for keyword seeding."""
        if not PPT_AVAILABLE:
            return {}
        result = self._dispatch('get_all_titles', timeout=5.0)
        return result or {}

    def next_slide(self):
        return bool(self._dispatch('next'))

    def previous_slide(self):
        return bool(self._dispatch('prev'))

    def jump_to(self, slide_num):
        return bool(self._dispatch('jump', slide_num))

    # ── COM operations — only called from _com_worker ──────────────────────

    def _error_state(self):
        return {**self._no_ppt_state(), 'ppt_message': 'PPT read error',
                'slide_title': ''}

    def _no_ppt_state(self):
        return {'name': '', 'ppt_ok': False,
                'ppt_message': 'PowerPoint not installed',
                'current_slide': 0, 'total_slides': 0,
                'is_slideshow': False, 'slide_title': ''}

    def _read_slide_title(self, slide_obj):
        """
        Safely read a slide's title text.
        Tries Shapes.Title first (fastest), then iterates placeholders as fallback.
        ppPlaceholderTitle=1, ppPlaceholderCenterTitle=3
        """
        try:
            title_shape = slide_obj.Shapes.Title
            if title_shape and title_shape.HasTextFrame:
                text = title_shape.TextFrame.Text.strip()
                if text:
                    return text
        except Exception:
            pass
        # Fallback: iterate all shapes and check placeholder type
        try:
            for i in range(1, slide_obj.Shapes.Count + 1):
                shape = slide_obj.Shapes(i)
                try:
                    ph_type = shape.PlaceholderFormat.Type
                    if ph_type in (1, 3) and shape.HasTextFrame:
                        text = shape.TextFrame.Text.strip()
                        if text:
                            return text
                except Exception:
                    continue
        except Exception:
            pass
        return ""

    def _do_get_info(self):
        if not self._ppt:
            try:
                self._ppt = win32com.client.Dispatch("PowerPoint.Application")
            except Exception:
                return {**self._no_ppt_state(), 'ppt_message': 'PowerPoint not running'}

        try:
            if self._ppt.Presentations.Count == 0:
                return {**self._no_ppt_state(), 'ppt_message': 'No presentation open'}

            self._pres = self._ppt.ActivePresentation
            total = self._pres.Slides.Count
            name  = self._pres.Name
            for ext in ('.pptx', '.ppt', '.PPTX', '.PPT'):
                if name.endswith(ext):
                    name = name[:-len(ext)]
                    break

            try:
                self._show = self._ppt.SlideShowWindows(1)
                current    = self._show.View.Slide.SlideIndex
                is_show    = True
            except Exception:
                self._show = None
                try:
                    current = self._ppt.ActiveWindow.Selection.SlideRange(1).SlideIndex
                except Exception:
                    current = 1
                is_show = False

            # Read the current slide's title text for display on phone
            slide_title = ""
            try:
                slide_obj   = self._pres.Slides(current)
                slide_title = self._read_slide_title(slide_obj)
            except Exception:
                pass

            return {
                'name':          name,
                'ppt_ok':        True,
                'ppt_message':   '',
                'current_slide': current,
                'total_slides':  total,
                'is_slideshow':  is_show,
                'slide_title':   slide_title,
            }

        except Exception:
            self._pres = None
            self._show = None
            return {**self._no_ppt_state(), 'ppt_message': 'No presentation open'}

    def _do_get_all_titles(self):
        """
        Returns {slide_number: title_string} for ALL slides in the active presentation.
        Called once when a presentation is first detected, to seed the keyword map.
        Runs on the COM thread so no cross-thread issues.
        """
        if not self._ppt or not self._pres:
            # Try to refresh _pres reference
            try:
                if self._ppt and self._ppt.Presentations.Count > 0:
                    self._pres = self._ppt.ActivePresentation
                else:
                    return {}
            except Exception:
                return {}

        titles = {}
        try:
            for i in range(1, self._pres.Slides.Count + 1):
                try:
                    slide = self._pres.Slides(i)
                    titles[i] = self._read_slide_title(slide)
                except Exception:
                    titles[i] = ""
        except Exception:
            pass
        return titles

    def _do_next(self):
        try:
            if self._show:
                self._show.View.Next()
                return True
        except Exception:
            self._show = None
        return False

    def _do_prev(self):
        try:
            if self._show:
                self._show.View.Previous()
                return True
        except Exception:
            self._show = None
        return False

    def _do_jump(self, n):
        try:
            n = int(n)
            if self._show and self._pres and 1 <= n <= self._pres.Slides.Count:
                self._show.View.GotoSlide(n)
                return True
        except Exception:
            self._show = None
        return False


# ─────────────────────────────────────────────────────────────────────────────
# Flask + SocketIO Setup
# ─────────────────────────────────────────────────────────────────────────────

app      = Flask(__name__)
app.config['SECRET_KEY'] = 'aero-present-2026'
socketio = SocketIO(app, cors_allowed_origins="*",
                    async_mode='threading', logger=False,
                    allow_upgrades=False)   # avoids Werkzeug dev-server WS 500 errors

ppt         = PowerPointController()
SERVER_IP   = get_local_ip()
SERVER_PORT = 5000

# ─────────────────────────────────────────────────────────────────────────────
# Mobile Backend — state + command bus
# ─────────────────────────────────────────────────────────────────────────────

class MobileBackend:
    def __init__(self):
        self.clients            = set()
        self.gesture_controller = None

        self.system_status = {
            'current_slide':          0,
            'total_slides':           0,
            'presentation_name':      '',
            'slide_title':            '',   # current slide's title text — NEW
            'ppt_ok':                 False,
            'ppt_message':            'Connecting…',
            'is_slideshow':           False,
            'gesture_enabled':        True,
            'laser_enabled':          True,
            'voice_enabled':          False,
            'hand_detected':          False,
            'back_hand_mode':         False,
            'orientation_confidence': 0.0,
            'zoom_active':            False,
            'zoom_level':             1,
            'draw_mode':              False,  # annotation draw mode state (set by gesture_control)
        }

        self._pending_changes  = {}

        # Keyword map: {slide_number: [keyword, ...]}
        self.slide_keywords = self.load_keywords()

        # Track the last known presentation name so we know when a new one opens
        self._last_presentation_name = ''

        t = threading.Thread(target=self._status_loop, daemon=True)
        t.start()

    # ── Status loop ────────────────────────────────────────────────────────

    def _status_loop(self):
        """
        Polls PowerPoint every 0.5s via COM and broadcasts live state to all phones.

        NEW: When the presentation name changes (new file opened), reads all
        slide titles and seeds the keyword map for slides with no keywords yet.
        This means voice commands like 'go to methodology' work automatically
        without the user manually typing keywords.
        """
        while True:
            info = ppt.get_info()

            new_name = info.get('name', '')
            ppt_ok   = info.get('ppt_ok', False)

            # ── Auto-seed keywords when a new presentation is detected ──────
            if ppt_ok and new_name and new_name != self._last_presentation_name:
                print(f"[KEYWORDS] New presentation: '{new_name}' — seeding from slide titles")
                self._last_presentation_name = new_name
                self._seed_keywords_from_titles()

            self.system_status.update({
                'current_slide':     info['current_slide'],
                'total_slides':      info['total_slides'],
                'presentation_name': new_name,
                'slide_title':       info.get('slide_title', ''),  # NEW
                'ppt_ok':            ppt_ok,
                'ppt_message':       info.get('ppt_message', ''),
                'is_slideshow':      info.get('is_slideshow', False),
            })

            if self.clients:
                socketio.emit('status_update', self.system_status)

            time.sleep(0.5)

    def _seed_keywords_from_titles(self):
        """
        Reads all slide titles from PowerPoint and pre-fills slide_keywords
        for any slide that has NO keywords assigned yet.

        Seeding rules:
          - Slide title "Research Methodology" → keywords: ["research", "methodology"]
          - Words shorter than 3 characters are dropped (articles, 'on', 'of', etc.)
          - Common filler words are stripped (see STOP_WORDS)
          - Never overwrites existing keywords — only fills empty slides
          - Saves to slide_keywords.json and broadcasts to all connected phones

        This lets the presenter say "go to methodology" immediately without
        opening the phone keyword editor.
        """
        STOP_WORDS = {
            'the', 'and', 'for', 'are', 'but', 'not', 'you', 'all', 'can',
            'her', 'was', 'one', 'our', 'out', 'day', 'get', 'has', 'him',
            'his', 'how', 'its', 'may', 'new', 'now', 'old', 'see', 'two',
            'who', 'did', 'any', 'been', 'from', 'had', 'have', 'that',
            'this', 'with', 'your', 'into', 'than', 'then', 'they', 'what',
            'when', 'will', 'more', 'also', 'some', 'would',
        }

        titles = ppt.get_all_slide_titles()
        if not titles:
            print("[KEYWORDS] No titles found — skipping seed")
            return

        seeded = 0
        for slide_num, title in titles.items():
            # Only seed slides that have no keywords at all
            if slide_num in self.slide_keywords and self.slide_keywords[slide_num]:
                continue

            if not title:
                continue

            # Tokenise: split on spaces/punctuation, lowercase, filter
            words = re.findall(r"[a-zA-Z']+", title.lower())
            keywords = [
                w for w in words
                if len(w) >= 3 and w not in STOP_WORDS
            ]

            if keywords:
                self.slide_keywords[slide_num] = keywords[:5]  # cap at 5 per slide
                seeded += 1
                print(f"[KEYWORDS] Slide {slide_num} '{title}' → {keywords[:5]}")

        if seeded:
            self.save_keywords()
            # Broadcast updated keywords to all connected phones immediately
            if self.clients:
                socketio.emit('keywords', self.slide_keywords)
            print(f"[KEYWORDS] Seeded {seeded} slides from slide titles")

    # ── Keyword persistence ────────────────────────────────────────────────

    def load_keywords(self):
        try:
            if os.path.exists("slide_keywords.json"):
                with open("slide_keywords.json") as f:
                    return {int(k): v for k, v in json.load(f).items()}
        except Exception:
            pass
        return {}

    def save_keywords(self):
        try:
            with open("slide_keywords.json", "w") as f:
                json.dump({str(k): v for k, v in self.slide_keywords.items()}, f)
        except Exception as e:
            print(f"[KEYWORDS] Save failed: {e}")

    # ── Command execution ─────────────────────────────────────────────────

    def execute_command(self, cmd, data=None):
        """
        Routes every command from phone or voice_control to the right handler.

        Slide navigation → PowerPoint COM (immediate, no frame delay)
        Toggles / zoom / tilt → _pending_changes (gesture_control picks up next frame)

        After NEXT/PREVIOUS/JUMP, immediately reads updated PPT state and pushes
        it to all phones — reduces perceived lag from 0-500ms to ~50ms.
        """
        data = data or {}

        # ── Slide navigation ───────────────────────────────────────────────
        if cmd == "NEXT":
            ok = ppt.next_slide()
            if ok:
                self._emit_slide_update_soon()
            return ok

        elif cmd == "PREVIOUS":
            ok = ppt.previous_slide()
            if ok:
                self._emit_slide_update_soon()
            return ok

        elif cmd == "JUMP":
            ok = ppt.jump_to(data.get('slide_number', 1))
            if ok:
                self._emit_slide_update_soon()
            return ok

        # ── Gesture-swipe notify — no COM nav, just pushes slide update ─────
        # Called by gesture_control after a Win32 PostMessage swipe so the
        # phone counter updates in ~120ms instead of waiting for the 0.5s poll.
        elif cmd in ("NEXT_NOTIFY", "PREV_NOTIFY"):
            self._emit_slide_update_soon()
            return True

        # ── Zoom ──────────────────────────────────────────────────────────
        elif cmd == "ZOOM_IN":
            current   = self.system_status.get('zoom_level', 1)
            new_level = min(current + 1, 4)
            self._queue_change('zoom_active', True)
            self._queue_change('zoom_level',  new_level)
            self.system_status['zoom_active'] = True
            self.system_status['zoom_level']  = new_level
            socketio.emit('zoom_changed', {'active': True, 'level': new_level})
            print(f"[ZOOM] Level → {new_level}")
            return True

        elif cmd == "ZOOM_OUT":
            current   = self.system_status.get('zoom_level', 1)
            new_level = max(current - 1, 1)
            active    = new_level > 1
            self._queue_change('zoom_active', active)
            self._queue_change('zoom_level',  new_level)
            self.system_status['zoom_active'] = active
            self.system_status['zoom_level']  = new_level
            socketio.emit('zoom_changed', {'active': active, 'level': new_level})
            print(f"[ZOOM] Level → {new_level}")
            return True

        elif cmd == "ZOOM_RESET":
            self._queue_change('zoom_active', False)
            self._queue_change('zoom_level',  1)
            self.system_status['zoom_active'] = False
            self.system_status['zoom_level']  = 1
            socketio.emit('zoom_changed', {'active': False, 'level': 1})
            print("[ZOOM] Reset")
            return True

        # ── Feature toggles ───────────────────────────────────────────────
        elif cmd == "GESTURE_TOGGLE":
            new_state = not self.system_status.get('gesture_enabled', True)
            self._queue_change('gesture_enabled', new_state)
            self.system_status['gesture_enabled'] = new_state
            socketio.emit('toggle_changed', {'feature': 'gesture', 'enabled': new_state})
            print(f"[TOGGLE] Gesture → {'ON' if new_state else 'OFF'}")
            return True

        elif cmd == "LASER_TOGGLE":
            new_state = not self.system_status.get('laser_enabled', True)
            self._queue_change('laser_enabled', new_state)
            self.system_status['laser_enabled'] = new_state
            socketio.emit('toggle_changed', {'feature': 'laser', 'enabled': new_state})
            print(f"[TOGGLE] Laser → {'ON' if new_state else 'OFF'}")
            return True

        elif cmd == "VOICE_TOGGLE":
            new_state = not self.system_status.get('voice_enabled', False)
            self._queue_change('voice_enabled', new_state)
            self.system_status['voice_enabled'] = new_state
            socketio.emit('toggle_changed', {'feature': 'voice', 'enabled': new_state})
            print(f"[TOGGLE] Voice → {'ON' if new_state else 'OFF'}")
            return True

        # ── Tilt pointer ──────────────────────────────────────────────────
        elif cmd == "TILT_MOVE":
            self._queue_change('tilt_x', data.get('x', 0.5))
            self._queue_change('tilt_y', data.get('y', 0.5))
            self._queue_change('tilt_active', True)
            return True

        elif cmd == "TILT_STOP":
            self._queue_change('tilt_active', False)
            return True

        # ── Slideshow control ──────────────────────────────────────────────
        # These queue a one-shot flag that gesture_control picks up next frame
        # and calls the corresponding Win32 function (ppt_start_slideshow, etc.).
        # Using pending_changes (not COM) because gesture_control owns the Win32
        # window handle — sending keys from this thread would steal focus.

        elif cmd == "PPT_START":
            # gesture_control will call ppt_start_slideshow() → sends F5 to PPT
            self._queue_change('ppt_start', True)
            print("[PPT] Start slideshow requested via phone")
            return True

        elif cmd == "PPT_STOP":
            # gesture_control will call ppt_exit_slideshow() → sends Esc to PPT
            self._queue_change('ppt_stop', True)
            print("[PPT] Stop slideshow requested via phone")
            return True

        elif cmd == "PPT_POINTER":
            # gesture_control will call ppt_pointer_mode() → sends Ctrl+P to PPT
            self._queue_change('ppt_pointer', True)
            print("[PPT] Pointer mode requested via phone")
            return True

        # ── Annotation ─────────────────────────────────────────────────────

        elif cmd == "ANNOTATION_TOGGLE":
            # gesture_control toggles draw_mode and calls overlay.set_draw_mode()
            self._queue_change('annotation_toggle', True)
            print("[ANNOTATION] Draw mode toggle requested via phone")
            return True

        elif cmd == "ANNOTATION_ERASE":
            # gesture_control calls overlay.clear_annotations()
            self._queue_change('annotation_erase', True)
            print("[ANNOTATION] Erase all requested via phone")
            return True

        return False

    def _emit_slide_update_soon(self):
        """
        Fires a background thread that waits ~120ms for PowerPoint to process the
        slide change, then reads and broadcasts the new slide state.

        Why 120ms: COM View.Next() sends a WM_COMMAND to the PPT window which
        processes asynchronously. Polling at 0ms reads the OLD slide; 80-150ms
        is reliable across machines. The 0.5s status_loop catches anything missed.
        """
        def _push():
            time.sleep(0.12)
            info = ppt.get_info()
            self.system_status.update({
                'current_slide': info['current_slide'],
                'slide_title':   info.get('slide_title', ''),
                'total_slides':  info['total_slides'],
                'ppt_ok':        info.get('ppt_ok', False),
            })
            if self.clients:
                socketio.emit('status_update', self.system_status)
        threading.Thread(target=_push, daemon=True).start()

    def _queue_change(self, key, value):
        with changes_lock:
            self._pending_changes[key] = value

    def pop_pending_changes(self):
        """Called by gesture_control.py every frame."""
        with changes_lock:
            if not self._pending_changes:
                return {}
            snapshot = dict(self._pending_changes)
            self._pending_changes.clear()
            return snapshot

    def update_system_status(self, **kwargs):
        """Called by gesture_control.py to push hand/orientation data."""
        self.system_status.update(kwargs)
        if self.clients:
            socketio.emit('status_update', self.system_status)

    def start_server(self, host='0.0.0.0', port=5000):
        global SERVER_PORT

        def _run():
            global SERVER_PORT
            actual_port = port
            for attempt in range(5):
                try:
                    SERVER_PORT = actual_port
                    print("=" * 60)
                    print("AERO PRESENT — Mobile Backend")
                    print(f"  PPT    : {'✓ connected' if PPT_AVAILABLE else '✗ not available'}")
                    print(f"  Server : http://{SERVER_IP}:{actual_port}/")
                    print(f"  Phone  : http://{SERVER_IP}:{actual_port}/?token={SESSION_TOKEN}")
                    print(f"  Token  : {SESSION_TOKEN}  (this session only)")
                    print("=" * 60)
                    socketio.run(app, host=host, port=actual_port,
                                 debug=False, use_reloader=False)
                    break
                except OSError as e:
                    if e.errno in (98, 10048):
                        print(f"[SERVER] Port {actual_port} in use, trying {actual_port + 1}…")
                        actual_port += 1
                    else:
                        print(f"[SERVER] Fatal OSError: {e}"); break
                except Exception as e:
                    print(f"[SERVER] Unexpected error: {e}"); break

        t = threading.Thread(target=_run, daemon=True)
        t.start()
        time.sleep(1.0)


# ─────────────────────────────────────────────────────────────────────────────
# Module-level backend instance
# ─────────────────────────────────────────────────────────────────────────────

backend = MobileBackend()

# ─────────────────────────────────────────────────────────────────────────────
# Socket.IO events
# ─────────────────────────────────────────────────────────────────────────────

@socketio.on('connect')
def on_connect():
    token = request.args.get('token')
    if token != SESSION_TOKEN:
        print("[SECURITY] Rejected connection — invalid token")
        return False
    sid = request.sid
    backend.clients.add(sid)
    print(f"[CLIENT] Connected — {len(backend.clients)} total")
    emit('status_update', backend.system_status)
    emit('keywords',      backend.slide_keywords)

@socketio.on('disconnect')
def on_disconnect():
    backend.clients.discard(request.sid)
    print(f"[CLIENT] Disconnected — {len(backend.clients)} remaining")

@socketio.on('command')
def on_command(data):
    cmd     = data.get('command')
    success = backend.execute_command(cmd, data.get('data'))
    emit('command_response', {'command': cmd, 'success': success})

@socketio.on('ping')
def on_ping(data):
    emit('pong', {'timestamp': data.get('timestamp')})

@socketio.on('keyword_add')
def on_keyword_add(data):
    slide = int(data.get('slide', 0))
    word  = str(data.get('keyword', '')).strip().lower()
    if slide > 0 and word:
        backend.slide_keywords.setdefault(slide, [])
        if word not in backend.slide_keywords[slide]:
            if len(backend.slide_keywords[slide]) >= 10:
                emit('keyword_error', {'error': f'Max 10 keywords per slide (slide {slide})'})
                return
            backend.slide_keywords[slide].append(word)
            backend.save_keywords()
            socketio.emit('keywords', backend.slide_keywords)
            print(f"[KEYWORDS] Added '{word}' → slide {slide}")

@socketio.on('keyword_remove')
def on_keyword_remove(data):
    slide = int(data.get('slide', 0))
    word  = str(data.get('keyword', '')).strip().lower()
    if slide in backend.slide_keywords:
        backend.slide_keywords[slide] = [
            k for k in backend.slide_keywords[slide] if k != word
        ]
        backend.save_keywords()
        socketio.emit('keywords', backend.slide_keywords)
        print(f"[KEYWORDS] Removed '{word}' from slide {slide}")

# ─────────────────────────────────────────────────────────────────────────────
# Flask routes
# ─────────────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for fname in ('mobile_app.html', 'Aero_Present.html'):
        for search_dir in (script_dir, os.getcwd()):
            path = os.path.join(search_dir, fname)
            if os.path.exists(path):
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        html = f.read()
                    html = html.replace('{{SERVER_URL}}', f'http://{SERVER_IP}:{SERVER_PORT}')
                    html = html.replace('{{SESSION_TOKEN}}', SESSION_TOKEN)
                    return Response(html, mimetype='text/html')
                except Exception as e:
                    return Response(f"""<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>Aero Present — Read Error</title></head><body>
<h1 style="color:red">File Read Error</h1>
<p>Found {fname} but could not open it: {e}</p></body></html>""",
                        mimetype='text/html', status=500)

    searched = sorted({
        os.path.join(script_dir, 'Aero_Present.html'),
        os.path.join(os.getcwd(), 'Aero_Present.html'),
    })
    paths_html = ''.join(f'<code style="display:block">{p}</code>' for p in searched)
    return Response(f"""<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>Aero Present — Setup Required</title></head><body style="background:#07091a;color:#f9fafb;font-family:sans-serif;padding:32px">
<h1 style="color:#ef4444">Setup Required</h1>
<p>Place <strong>Aero_Present.html</strong> in the same folder as mobile_backend.py</p>
{paths_html}<p>Server: http://{SERVER_IP}:{SERVER_PORT}/</p></body></html>""",
        mimetype='text/html', status=404)

@app.route('/config')
def config():
    return jsonify({
        'server_url':  f'http://{SERVER_IP}:{SERVER_PORT}',
        'server_ip':   SERVER_IP,
        'server_port': SERVER_PORT,
    })

@app.route('/health')
def health():
    return jsonify({
        'status':    'running',
        'ppt':       'available' if PPT_AVAILABLE else 'disabled',
        'clients':   len(backend.clients),
        'server_ip': SERVER_IP,
    })

@app.route('/status')
def status():
    return jsonify(backend.system_status)

@app.route('/keywords', methods=['GET'])
def get_keywords():
    return jsonify(backend.slide_keywords)

@app.route('/keywords', methods=['POST'])
def set_keywords():
    try:
        data = request.get_json()
        backend.slide_keywords = {int(k): v for k, v in data.items()}
        backend.save_keywords()
        socketio.emit('keywords', backend.slide_keywords)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 400

@app.route('/slides')
def get_slide_titles():
    """
    Returns all slide titles as {slide_number: title_string}.
    The phone's keyword editor calls this to show slide names next to slide numbers,
    making it easy to know which slide is which when assigning keywords.
    """
    titles = ppt.get_all_slide_titles()
    return jsonify({str(k): v for k, v in titles.items()})

# ─────────────────────────────────────────────────────────────────────────────
# Entry point for gesture_control.py
# ─────────────────────────────────────────────────────────────────────────────

def initialize_mobile_backend(gesture_controller=None):
    global backend
    backend.gesture_controller = gesture_controller
    return backend

# ─────────────────────────────────────────────────────────────────────────────
# Standalone run
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    backend.start_server()
    while True:
        time.sleep(1)
