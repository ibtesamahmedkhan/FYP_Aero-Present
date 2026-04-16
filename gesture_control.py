"""
Aero Present - Gesture Control System
Multi-modal presentation controller using hand tracking

Controls:
Q - quit | G - toggle gestures | L - toggle laser | H - HUD
D - debug | R - reset hand lock | C - recalibrate
V - toggle voice control | Z - cycle zoom (keyboard)
A - toggle annotation draw mode (LASER traces strokes when ON)
5 - start slideshow (F5)  |  E - exit slideshow (Esc)
P - PPT pointer mode (Ctrl+P)

Gestures:
Index+Middle up       → SWIPE  (slide left/right, 3-frame confirm gate)
Index only up         → LASER  (pointer / draw when A is ON)
Index+Middle+Ring up  → ERASE  (clears all annotation strokes)
All fingers down      → FIST   (hold 0.5s to activate zoom, release to deactivate)
"""

import csv, math, os, struct, sys, threading, time

# QR code display on startup (optional — pip install qrcode pillow)
try:
    import qrcode as _qrcode_lib
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False
import cv2
import mediapipe as mp
from mediapipe.tasks import python
from mediapipe.tasks.python import vision
import pyautogui
import win32api, win32con, win32gui

from utils import (
    Smoother, SwipeDetector, classify_hand, map_to_screen,
    OrientationTracker, MovementValidator, BackHandAccuracyLogger
)

# ── Optional Tier 3: mobile backend ──────────────────────────────────────────
try:
    from mobile_backend import initialize_mobile_backend as _init_mobile
    MOBILE_AVAILABLE = True
except ImportError:
    MOBILE_AVAILABLE = False
    print("[AERO] Running without mobile backend")

# ── Optional Tier 2: voice control ───────────────────────────────────────────
try:
    from voice_control import VoiceController
    VOICE_AVAILABLE = True
except ImportError:
    VOICE_AVAILABLE = False
    print("[AERO] Running without voice control")

pyautogui.FAILSAFE = False
pyautogui.PAUSE    = 0.0

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

WEBCAM_INDEX   = 0
FRAME_WIDTH    = 640
FRAME_HEIGHT   = 480

# Swipe thresholds — tuned through testing sessions
SWIPE_VELOCITY    = 0.028
SWIPE_DISPLACEMENT = 0.10    # min total travel across detection window
SWIPE_COOLDOWN    = 20     # frames between swipes

# Laser pointer
POINTER_RADIUS      = 10
POINTER_COLOR_RGB   = (255, 0, 0)   # red — used by Win32 RGB brush in LaserOverlay._draw()
POINTER_SMOOTH_WINDOW = 8

# MediaPipe confidence — 0.3 works well for varied lighting
MIN_DETECTION_CONF = 0.30
MIN_TRACKING_CONF  = 0.30

# Hand lock — max wrist drift before treating as a different hand
HAND_LOCK_DISTANCE = 0.35

SCREEN_W, SCREEN_H = pyautogui.size()

# Hand skeleton connections (MediaPipe 21-landmark layout)
# Defined once at module level — avoids allocating this list every frame
HAND_CONNECTIONS = [
    (0,1),(1,2),(2,3),(3,4),         # thumb
    (0,5),(5,6),(6,7),(7,8),          # index
    (0,9),(9,10),(10,11),(11,12),     # middle
    (0,13),(13,14),(14,15),(15,16),   # ring
    (0,17),(17,18),(18,19),(19,20),   # pinky
    (5,9),(9,13),(13,17),             # palm cross-connections
]

# 3-frame swipe confirmation — SWIPE gesture must appear for this many
# consecutive frames before SwipeDetector.update() is called.
# Eliminates false triggers from brief SWIPE positions during LASER→SWIPE
# transitions without adding noticeable delay (~100ms at 30fps).
SWIPE_CONFIRM_FRAMES = 2  # 2 consecutive SWIPE frames required before detector fires

# Z-depth threshold scaling — reference wrist-to-middle-MCP distance in metres
# at ~0.5m presenter distance. apply_distance_scale() uses this baseline.
PALM_REFERENCE_METRES = 0.085

# ─────────────────────────────────────────────────────────────────────────────
# Helper: QR code generator
# ─────────────────────────────────────────────────────────────────────────────

def _make_qr_overlay(url, size=180):
    """Return a numpy BGR array containing the QR code for url, or None.
    Shown in the camera preview for the first 6 seconds after server starts
    so the presenter can scan it directly rather than typing the IP address."""
    if not QR_AVAILABLE:
        return None
    try:
        import numpy as np
        from PIL import Image
        qr = _qrcode_lib.QRCode(box_size=4, border=2)
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        img = img.resize((size, size), Image.LANCZOS)
        arr = np.array(img.convert('RGB'))
        return arr[:, :, ::-1].copy()   # RGB → BGR, C-contiguous for cv2
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Helper: erase gesture detector
# ─────────────────────────────────────────────────────────────────────────────

def _is_erase_gesture(lm):
    """True when index+middle+ring are up and pinky is down.
    Distinct from SWIPE (index+middle, no ring) and LASER (index only).
    Called inline each frame — does not touch utils.py."""
    idx  = lm[8].y  < lm[6].y    # index tip above PIP joint
    mid  = lm[12].y < lm[10].y   # middle tip above PIP joint
    ring = lm[16].y < lm[14].y   # ring tip above PIP joint
    pink = lm[20].y < lm[18].y   # pinky
    return idx and mid and ring and not pink


# ─────────────────────────────────────────────────────────────────────────────
# Session Logger
# ─────────────────────────────────────────────────────────────────────────────

class SessionLogger:
    def __init__(self):
        os.makedirs("logs", exist_ok=True)
        timestamp     = time.strftime("%Y%m%d_%H%M%S")
        self.filename = f"logs/session_{timestamp}.csv"
        self.csv_file = open(self.filename, 'w', newline='', encoding='utf-8')
        self.writer   = csv.writer(self.csv_file)
        self.writer.writerow([
            "timestamp", "event_type", "hand_detected", "orientation",
            "confidence", "gesture_type", "is_ghost_frame", "notes"
        ])
        self.csv_file.flush()
        self._last_flush  = time.time()
        self._flush_every = 1.0   # seconds — flush at most once per second
        print(f"[LOG] Session: {self.filename}")

    def log(self, event_type, hand_detected=False, orientation="",
            confidence=0.0, gesture_type="", is_ghost_frame=False, notes=""):
        self.writer.writerow([
            time.time(), event_type, hand_detected, orientation,
            f"{confidence:.2f}", gesture_type, is_ghost_frame, notes
        ])
        # Only flush to disk every _flush_every seconds.
        # Previously flushed every single write (~30/s during laser) causing disk I/O spikes.
        now = time.time()
        if now - self._last_flush >= self._flush_every:
            self.csv_file.flush()
            self._last_flush = now

    def close(self):
        self.csv_file.flush()   # final flush on shutdown
        self.csv_file.close()

# ─────────────────────────────────────────────────────────────────────────────
# Laser Overlay — threaded Win32 transparent window with HUD strip
# Runs on a background thread so it never blocks the camera loop.
# ─────────────────────────────────────────────────────────────────────────────

class LaserOverlay:
    def __init__(self, screen_width, screen_height):
        self.screen_width  = screen_width
        self.screen_height = screen_height
        self._dot_x        = screen_width  // 2
        self._dot_y        = screen_height // 2
        self._dot_visible  = False
        self._hwnd         = None
        self._lock         = threading.Lock()
        self._active       = True

        # HUD state mirrors
        self._hud_gesture_on = True
        self._hud_laser_on   = True
        self._hud_voice_on   = False
        self._hud_visible    = True

        # Wake-word indicator — lights up when "AERO" is detected
        self._wake_active    = False

        # ── Annotation drawing state ──────────────────────────────────────────
        # _strokes      : completed strokes, each a list of (x, y) screen pixels
        # _curr_stroke  : stroke currently being traced by the laser
        # _draw_mode    : True = LASER gesture leaves persistent lines
        # _draw_prev    : previous point, used to draw continuous line segments
        self._strokes      = []
        self._curr_stroke  = []
        self._draw_mode    = False
        self._draw_prev    = None

        t = threading.Thread(target=self._message_loop, daemon=True)
        t.start()
        time.sleep(0.3)  # wait for window to be ready

    def _wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
        elif msg == win32con.WM_PAINT:
            self._draw(hwnd)
            return 0
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def _draw(self, hwnd):
        dc, ps = win32gui.BeginPaint(hwnd)
        try:
            # Double-buffered paint: draw everything into a memory DC first,
            # then BitBlt the completed frame to the window in one atomic operation.
            # This eliminates the blank-frame gap between erase and paint.
            mem_dc  = win32gui.CreateCompatibleDC(dc)
            bmp     = win32gui.CreateCompatibleBitmap(dc, self.screen_width, self.screen_height)
            old_bmp = win32gui.SelectObject(mem_dc, bmp)

            # Fill memory DC with black — colour key makes this transparent on screen
            win32gui.FillRect(mem_dc, (0, 0, self.screen_width, self.screen_height), win32gui.GetStockObject(win32con.BLACK_BRUSH))

            with self._lock:
                # laser dot
                if self._dot_visible:
                    r  = POINTER_RADIUS
                    br = win32gui.CreateSolidBrush(win32api.RGB(*POINTER_COLOR_RGB))
                    ob = win32gui.SelectObject(mem_dc, br)
                    win32gui.Ellipse(mem_dc, self._dot_x-r, self._dot_y-r,self._dot_x+r, self._dot_y+r)
                    win32gui.SelectObject(mem_dc, ob)
                    win32gui.DeleteObject(br)

                # annotation strokes — drawn after dot so strokes appear behind cursor
                # Bright yellow (255, 220, 0) is visible on most slide backgrounds.
                # Each stroke is a list of (x, y) screen pixel points.
                if self._strokes or self._curr_stroke:
                    ann_pen = win32gui.CreatePen(win32con.PS_SOLID, 3, win32api.RGB(255, 220, 0))
                    old_pen = win32gui.SelectObject(mem_dc, ann_pen)
                    for stroke in self._strokes:
                        if len(stroke) >= 2:
                            win32gui.MoveToEx(mem_dc, stroke[0][0], stroke[0][1])
                            for pt in stroke[1:]:
                                win32gui.LineTo(mem_dc, pt[0], pt[1])
                    if len(self._curr_stroke) >= 2:
                        win32gui.MoveToEx(mem_dc, self._curr_stroke[0][0], self._curr_stroke[0][1])
                        for pt in self._curr_stroke[1:]:
                            win32gui.LineTo(mem_dc, pt[0], pt[1])
                    win32gui.SelectObject(mem_dc, old_pen)
                    win32gui.DeleteObject(ann_pen)

                # HUD strip
                if self._hud_visible:
                    H = 36
                    bb = win32gui.CreateSolidBrush(win32api.RGB(20, 20, 30))
                    ob = win32gui.SelectObject(mem_dc, bb)
                    win32gui.Rectangle(mem_dc, 0, 0, self.screen_width, H)
                    win32gui.SelectObject(mem_dc, ob)
                    win32gui.DeleteObject(bb)

                    fs = win32gui.LOGFONT()
                    fs.lfHeight   = 20
                    fs.lfWeight   = win32con.FW_BOLD
                    fs.lfFaceName = "Segoe UI"
                    font = win32gui.CreateFontIndirect(fs)
                    of   = win32gui.SelectObject(mem_dc, font)
                    win32gui.SetBkMode(mem_dc, win32con.TRANSPARENT)

                    def label(text, x, on):
                        col = win32api.RGB(0, 220, 80) if on else win32api.RGB(220, 50, 50)
                        win32gui.SetTextColor(mem_dc, col)
                        win32gui.ExtTextOut(mem_dc, x, 8, 0, None, text, ())

                    label(f"GESTURE: {'ON' if self._hud_gesture_on else 'OFF'}", 10,  self._hud_gesture_on)
                    label(f"LASER: {'ON' if self._hud_laser_on else 'OFF'}",    200, self._hud_laser_on)
                    label(f"VOICE: {'ON' if self._hud_voice_on else 'OFF'}",    360, self._hud_voice_on)

                    # Draw mode indicator — shown when annotation drawing is on
                    if self._draw_mode:
                        win32gui.SetTextColor(mem_dc, win32api.RGB(255, 220, 0))
                        win32gui.ExtTextOut(mem_dc, 510, 8, 0, None, "* DRAW ON", ())

                    # Wake word badge
                    if self._wake_active:
                        badge_br = win32gui.CreateSolidBrush(win32api.RGB(30, 30, 0))
                        ob2 = win32gui.SelectObject(mem_dc, badge_br)
                        win32gui.Rectangle(mem_dc, 510, 3, 720, H - 3)
                        win32gui.SelectObject(mem_dc, ob2)
                        win32gui.DeleteObject(badge_br)
                        win32gui.SetTextColor(mem_dc, win32api.RGB(255, 230, 0))
                        win32gui.ExtTextOut(mem_dc, 518, 8, 0, None, "* AERO LISTENING", ())

                    win32gui.SetTextColor(mem_dc, win32api.RGB(120, 120, 120))
                    win32gui.ExtTextOut(mem_dc, self.screen_width - 180, 8, 0, None, "AERO PRESENT", ())

                    win32gui.SelectObject(mem_dc, of)
                    win32gui.DeleteObject(font)

            # Blit completed frame to window DC in one operation — no visible gap
            win32gui.BitBlt(dc, 0, 0, self.screen_width, self.screen_height,
                            mem_dc, 0, 0, win32con.SRCCOPY)

            win32gui.SelectObject(mem_dc, old_bmp)
            win32gui.DeleteObject(bmp)
            win32gui.DeleteDC(mem_dc)

        finally:
            win32gui.EndPaint(hwnd, ps)

    def _message_loop(self):
        wc               = win32gui.WNDCLASS()
        wc.hInstance     = win32api.GetModuleHandle(None)
        wc.lpszClassName = "AeroPresentOverlay"
        wc.lpfnWndProc   = self._wnd_proc
        wc.hbrBackground = win32gui.GetStockObject(win32con.BLACK_BRUSH)
        win32gui.RegisterClass(wc)

        self._hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT |
            win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW,
            "AeroPresentOverlay", "AeroOverlay", win32con.WS_POPUP,
            0, 0, self.screen_width, self.screen_height,
            None, None, wc.hInstance, None
        )
        win32gui.SetLayeredWindowAttributes(
            self._hwnd, win32api.RGB(0,0,0), 0, win32con.LWA_COLORKEY
        )
        win32gui.ShowWindow(self._hwnd, win32con.SW_SHOW)
        win32gui.UpdateWindow(self._hwnd)

        while self._active:
            win32gui.PumpWaitingMessages()
            time.sleep(0.01)

    def _refresh(self):
        if self._hwnd:
            win32gui.InvalidateRect(self._hwnd, None, False)  # False = don't erase bg before paint

    def move(self, x, y):
        with self._lock:
            self._dot_x, self._dot_y = x, y
        self._refresh()

    def move_and_show(self, x, y):
        """Update position AND make visible in one refresh — avoids double InvalidateRect."""
        with self._lock:
            self._dot_x, self._dot_y = x, y
            self._dot_visible = True
        self._refresh()

    def show(self):
        with self._lock:
            self._dot_visible = True
        self._refresh()

    def hide(self):
        with self._lock:
            self._dot_visible = False
        self._refresh()

    def update_hud(self, gesture_on, laser_on, voice_on=None):
        with self._lock:
            self._hud_gesture_on = gesture_on
            self._hud_laser_on   = laser_on
            if voice_on is not None:
                self._hud_voice_on = voice_on
        self._refresh()

    def show_wake_indicator(self, active: bool):
        """Light up the AERO LISTENING badge in the HUD strip."""
        with self._lock:
            self._wake_active = active
        self._refresh()

    def toggle_hud(self):
        with self._lock:
            self._hud_visible = not self._hud_visible
        self._refresh()

    # ── Annotation drawing methods ────────────────────────────────────────────

    def begin_stroke(self, x, y):
        """Start a new annotation stroke at (x, y)."""
        with self._lock:
            self._curr_stroke = [(x, y)]
            self._draw_prev   = (x, y)

    def extend_stroke(self, x, y):
        """Append a point to the current stroke and redraw."""
        with self._lock:
            if self._draw_prev is not None:
                self._curr_stroke.append((x, y))
                self._draw_prev = (x, y)
        self._refresh()

    def end_stroke(self):
        """Commit the current stroke to completed strokes."""
        with self._lock:
            if len(self._curr_stroke) > 1:
                self._strokes.append(list(self._curr_stroke))
            self._curr_stroke = []
            self._draw_prev   = None
        self._refresh()

    def clear_annotations(self):
        """Erase all annotation strokes and the current stroke."""
        with self._lock:
            self._strokes     = []
            self._curr_stroke = []
            self._draw_prev   = None
        self._refresh()
        print("[DRAW] Annotations cleared")

    def set_draw_mode(self, active: bool):
        """Toggle annotation draw mode. Disabling commits any open stroke."""
        with self._lock:
            self._draw_mode = active
        if not active:
            self.end_stroke()
        print(f"[DRAW] Mode {'ON — LASER now traces strokes' if active else 'OFF'}")

    def destroy(self):
        self._active = False
        if self._hwnd:
            try:
                win32gui.PostMessage(self._hwnd, win32con.WM_CLOSE, 0, 0)
            except Exception:
                pass

# ─────────────────────────────────────────────────────────────────────────────
# PowerPoint control — Win32 PostMessage (no focus steal)
# ─────────────────────────────────────────────────────────────────────────────

def _find_ppt_window():
    found = {"hwnd": None}
    def check(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title      = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            if class_name == "screenClass" or "Slide Show" in title:
                found["hwnd"] = hwnd
            elif not found["hwnd"] and ("PowerPoint" in title or title.endswith(".pptx")):
                found["hwnd"] = hwnd
    win32gui.EnumWindows(check, None)
    return found["hwnd"]

def _send_key(hwnd, vk):
    win32api.PostMessage(hwnd, 0x0100, vk, 0)   # WM_KEYDOWN
    time.sleep(0.01)
    win32api.PostMessage(hwnd, 0x0101, vk, 0)   # WM_KEYUP

def next_slide():
    hwnd = _find_ppt_window()
    if hwnd:
        _send_key(hwnd, 0x27)   # VK_RIGHT — no focus steal
        print("[AERO] Next slide")
    else:
        pyautogui.press('right')   # fallback if PPT window not found

def prev_slide():
    hwnd = _find_ppt_window()
    if hwnd:
        _send_key(hwnd, 0x25)   # VK_LEFT — no focus steal
        print("[AERO] Prev slide")
    else:
        pyautogui.press('left')    # fallback


def ppt_start_slideshow():
    """F5 — start slideshow from slide 1."""
    hwnd = _find_ppt_window()
    if hwnd:
        _send_key(hwnd, win32con.VK_F5)
    else:
        pyautogui.press('f5')
    print("[AERO] Start slideshow (F5)")


def ppt_exit_slideshow():
    """Escape — exit slideshow and return to edit view."""
    hwnd = _find_ppt_window()
    if hwnd:
        _send_key(hwnd, win32con.VK_ESCAPE)
    else:
        pyautogui.press('escape')
    print("[AERO] Exit slideshow (Esc)")


def ppt_pointer_mode():
    """Ctrl+P — activate PowerPoint built-in pen/pointer mode."""
    pyautogui.hotkey('ctrl', 'p')
    print("[AERO] PPT pointer mode (Ctrl+P)")


# ─────────────────────────────────────────────────────────────────────────────
# Calibration — 5-second detection rate measurement
# ─────────────────────────────────────────────────────────────────────────────

def calibrate_confidence(detector, cap):
    print("\n╔══════════════════════════════════════╗")
    print("║   CALIBRATION — wave your hand       ║")
    print("║          for 5 seconds...            ║")
    print("╚══════════════════════════════════════╝\n")

    total    = 0
    detected = 0
    target   = 150  # ~5 seconds at 30fps

    while total < target:
        ok, frame = cap.read()
        if not ok:
            continue

        frame    = cv2.flip(frame, 1)
        rgb      = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        mp_img   = mp.Image(image_format=mp.ImageFormat.SRGB, data=rgb)
        result   = detector.detect(mp_img)

        bar_w = int((total / target) * FRAME_WIDTH)
        cv2.rectangle(frame, (0, FRAME_HEIGHT-20), (bar_w, FRAME_HEIGHT), (0,255,0), -1)
        cv2.putText(frame, "Calibrating...", (10,30),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,255,0), 2)

        if result.hand_landmarks:
            detected += 1

        cv2.imshow("Calibration", frame)
        cv2.waitKey(1)
        total += 1

    cv2.destroyWindow("Calibration")
    rate = detected / target
    print(f"[CALIBRATION] Detection rate: {rate:.1%}")

    if rate > 0.85:
        det, trk = 0.60, 0.60
        print("[CALIBRATION] Excellent — strict thresholds")
    elif rate > 0.70:
        det, trk = 0.50, 0.50
        print("[CALIBRATION] Good — balanced thresholds")
    elif rate > 0.50:
        det, trk = 0.40, 0.40
        print("[CALIBRATION] Moderate — relaxed thresholds")
    else:
        det, trk = 0.30, 0.30
        print("[CALIBRATION] Poor lighting — very relaxed")

    return det, trk

# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "="*60)
    print("AERO PRESENT - Gesture Control System")
    print("="*60 + "\n")

    # ── Preflight checks ──────────────────────────────────────────────────────
    MODEL_FILE = 'hand_landmarker.task'
    if not os.path.exists(MODEL_FILE):
        print("╔══════════════════════════════════════════════════════════╗")
        print("║  ERROR: hand_landmarker.task not found                  ║")
        print("║  Download it from:                                      ║")
        print("║  https://ai.google.dev/edge/mediapipe/solutions/        ║")
        print("║          vision/hand_landmarker                         ║")
        print("║  Place the file in the same folder as gesture_control.py║")
        print("╚══════════════════════════════════════════════════════════╝")
        sys.exit(1)

    # camera
    cap = cv2.VideoCapture(WEBCAM_INDEX)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH,  FRAME_WIDTH)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, FRAME_HEIGHT)
    if not cap.isOpened():
        print("[ERROR] Can't open camera")
        sys.exit(1)

    # mediapipe — IMAGE mode for calibration, VIDEO mode for main loop
    base_options = python.BaseOptions(model_asset_path='hand_landmarker.task')

    def make_detector(det_conf, trk_conf, mode):
        opts = vision.HandLandmarkerOptions(
            base_options                  = base_options,
            running_mode                  = mode,
            num_hands                     = 1,
            min_hand_detection_confidence = det_conf,
            min_hand_presence_confidence  = trk_conf,
            min_tracking_confidence       = trk_conf,
        )
        return vision.HandLandmarker.create_from_options(opts)

    det_conf = MIN_DETECTION_CONF
    trk_conf = MIN_TRACKING_CONF

    # calibration prompt (3-second window)
    print("╔══════════════════════════════════════╗")
    print("║  Press C to calibrate (5 sec)        ║")
    print("║  or wait to use default settings     ║")
    print("╚══════════════════════════════════════╝")

    calib_detector  = make_detector(det_conf, trk_conf, vision.RunningMode.IMAGE)
    want_calibration = False
    deadline         = time.time() + 3

    while time.time() < deadline:
        # Show live camera feed with countdown so user knows the app is running
        ok, frame = cap.read()
        if ok:
            frame     = cv2.flip(frame, 1)
            remaining = max(0, int(deadline - time.time()) + 1)
            cv2.putText(frame, f"Press C to calibrate ({remaining}s)...",
                        (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
            cv2.putText(frame, "or wait to start with defaults",
                        (10, 58), cv2.FONT_HERSHEY_SIMPLEX, 0.55, (150, 150, 150), 1)
            cv2.imshow("Aero Present - Camera", frame)
        k = cv2.waitKey(10) & 0xFF
        if k in (ord('c'), ord('C')):
            want_calibration = True
            break

    if want_calibration:
        det_conf, trk_conf = calibrate_confidence(calib_detector, cap)
    else:
        print("[AERO] Using default thresholds")

    calib_detector.close()

    # main VIDEO detector
    detector = make_detector(det_conf, trk_conf, vision.RunningMode.VIDEO)

    # logger + utils
    logger          = SessionLogger()
    logger.log("SESSION_START", notes=f"det={det_conf} trk={trk_conf}")
    orient_tracker  = OrientationTracker(transition_frames_required=5)
    move_validator  = MovementValidator(history_frames=10)
    backhand_logger = BackHandAccuracyLogger()

    swipe_detector = SwipeDetector(
        palm_velocity_threshold     = SWIPE_VELOCITY,
        palm_displacement_threshold = SWIPE_DISPLACEMENT,
        back_velocity_threshold     = 0.032,
        back_displacement_threshold = 0.12,
        cooldown_frames             = SWIPE_COOLDOWN,
    )

    laser_smooth_x = Smoother(window_size=POINTER_SMOOTH_WINDOW)
    laser_smooth_y = Smoother(window_size=POINTER_SMOOTH_WINDOW)
    overlay        = LaserOverlay(SCREEN_W, SCREEN_H)

    # ── Tier 3: mobile ────────────────────────────────────────────────────────
    mobile_server = None
    if MOBILE_AVAILABLE:
        try:
            mobile_server = _init_mobile()
            mobile_server.start_server(host='0.0.0.0', port=5000)
            # SERVER_PORT may have been bumped to 5001-5004 if 5000 was in use
            from mobile_backend import SERVER_PORT as _actual_port, SERVER_IP as _actual_ip, SESSION_TOKEN as _token
            print(f"[MOBILE] Server started on port {_actual_port}")
            # QR code — shown in the camera preview for ~6 seconds on startup
            _phone_url  = f"http://{_actual_ip}:{_actual_port}/?token={_token}"
            _qr_img     = _make_qr_overlay(_phone_url, size=180)
            qr_overlay     = _qr_img
            qr_show_frames = 180 if _qr_img is not None else 0
            if _qr_img is not None:
                print(f"[QR] Showing QR code for 6s — or scan: {_phone_url}")
            else:
                print(f"[QR] pip install qrcode pillow  to enable QR display")
                print(f"[QR] Phone URL: {_phone_url}")
        except Exception as e:
            print(f"[MOBILE] Failed to start: {e}")
            mobile_server  = None
            qr_overlay     = None
            qr_show_frames = 0
    else:
        qr_overlay     = None
        qr_show_frames = 0

    # ── Tier 2: voice ─────────────────────────────────────────────────────────
    voice = None
    if VOICE_AVAILABLE:
        try:
            kw_map = mobile_server.slide_keywords if mobile_server else {}
            voice  = VoiceController(mobile_backend=mobile_server, keyword_map=kw_map)
            voice.start_listening()
            print("[VOICE] Listening started")
        except Exception as e:
            print(f"[VOICE] Failed to start: {e}")
            voice = None

    # ── Runtime state ─────────────────────────────────────────────────────────
    gesture_enabled        = True
    laser_enabled          = True
    voice_enabled          = voice is not None
    show_hud               = True
    debug_mode             = False

    # ghost hand state
    last_good_landmarks        = None   # most recent valid 21-point set
    ghost_frames_remaining     = 0      # frames left in ghost window
    in_ghost_mode              = False  # True while showing ghost (orange)
    # Ghost hand: if MediaPipe briefly drops the hand (e.g. fast swipe motion blur),
    # we re-use the last known landmarks for up to 10 frames rather than
    # instantly losing tracking. Shown in orange so you know it is buffered.

    # hand lock state
    locked_wrist     = None    # wrist position of the first detected hand
    lock_set         = False   # True once the lock reference exists
    last_orientation = "palm" 

    # laser persistence — keeps dot for 5 frames after pose changes
    laser_persist    = 0

    # zoom state (controlled by phone, Z key, or fist-hold gesture)


    # tilt pointer state (phone gyro controls laser when active)
    tilt_active      = False
    tilt_x           = 0.5
    tilt_y           = 0.5

    # 3-frame swipe confirmation state
    _swipe_gesture_frames = 0   # consecutive frames of SWIPE gesture detected

    # annotation drawing state
    draw_mode            = False   # True = LASER gesture traces strokes
    _ann_was_laser       = False   # True when last frame was LASER (for end_stroke on exit)
    _erase_cooldown      = 0       # frames until erase gesture can fire again (prevents repeat)

    # gesture visual feedback — keeps the NEXT/PREV banner visible for ~1.5s after firing
    # without persistence it disappears after one frame, too fast to read on camera
    _feedback_text       = ""
    _feedback_color      = (0, 255, 0)
    _feedback_frames     = 0       # counts down from FEEDBACK_HOLD_FRAMES to 0
    FEEDBACK_HOLD_FRAMES = 45      # ~1.5s at 30fps

    # mobile status mirrors (updated inside SWIPE branch)
    _back_hand_mode  = False
    _orient_conf     = 0.0

    # fps
    frame_count = 0
    fps_start   = time.time()
    fps         = 0

    overlay.update_hud(gesture_enabled, laser_enabled, voice_enabled)

    print("\n╔══════════════════════════════════════╗")
    print("║        AERO PRESENT — ACTIVE         ║")
    print("╠══════════════════════════════════════╣")
    print("║  Q - quit  | G - gestures            ║")
    print("║  L - laser | H - HUD                 ║")
    print("║  D - debug | R - reset lock          ║")
    print("║  C - recalibrate | V - voice         ║")
    print("║  Z - cycle zoom (2× → 3× → 4× → off)║")
    print("╚══════════════════════════════════════╝\n")

    try:
        while True:
            # ── mobile pending changes ─────────────────────────────────────────
            if mobile_server:
                changes = mobile_server.pop_pending_changes()
                # Validate: pop_pending_changes() should return a dict
                if changes and isinstance(changes, dict):
                    if 'gesture_enabled' in changes:
                        gesture_enabled = changes['gesture_enabled']
                    if 'laser_enabled' in changes:
                        laser_enabled = changes['laser_enabled']
                        if not laser_enabled:
                            overlay.hide()
                    if 'voice_enabled' in changes and voice:
                        voice_enabled        = changes['voice_enabled']
                        voice.enabled        = voice_enabled
                        if voice_enabled:
                            voice.start_listening()
                        else:
                            voice.stop_listening()
                    if 'tilt_active' in changes:
                        tilt_active = changes['tilt_active']
                        if not tilt_active:
                            overlay.hide()
                    if 'tilt_x' in changes:
                        tilt_x = changes['tilt_x']
                    if 'tilt_y' in changes:
                        tilt_y = changes['tilt_y']
                    # ── Slideshow control from phone ─────────────────────────
                    if 'ppt_start' in changes and changes['ppt_start']:
                        ppt_start_slideshow()
                        logger.log("PPT_START")
                    if 'ppt_stop' in changes and changes['ppt_stop']:
                        ppt_exit_slideshow()
                        logger.log("PPT_STOP")
                    if 'ppt_pointer' in changes and changes['ppt_pointer']:
                        ppt_pointer_mode()
                        logger.log("PPT_POINTER")
                    # ── Annotation from phone ─────────────────────────────────
                    if 'annotation_toggle' in changes and changes['annotation_toggle']:
                        draw_mode = not draw_mode
                        overlay.set_draw_mode(draw_mode)
                        logger.log(f"DRAW_MODE_{'ON' if draw_mode else 'OFF'}")
                    if 'annotation_erase' in changes and changes['annotation_erase']:
                        overlay.clear_annotations()
                        logger.log("ANNOTATION_ERASE")
                    overlay.update_hud(gesture_enabled, laser_enabled, voice_enabled)

            # ── Wake word indicator ────────────────────────────────────────────
            # Polls VoiceController.wake_word_active and lights the HUD badge
            if voice and hasattr(voice, 'wake_word_active'):
                overlay.show_wake_indicator(voice.wake_word_active)

            # ── camera read ────────────────────────────────────────────────────
            ret, frame = cap.read()
            if not ret:
                print("[ERROR] Camera read failed")
                break

            frame       = cv2.flip(frame, 1)
            frame_count += 1
            ts_ms        = int(time.time() * 1000)

            if frame_count % 30 == 0:
                fps       = 30 / (time.time() - fps_start)
                fps_start = time.time()

            # ── MediaPipe detection ────────────────────────────────────────────
            rgb    = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            mp_img = mp.Image(image_format=mp.ImageFormat.SRGB, data=rgb)
            result = detector.detect_for_video(mp_img, ts_ms)

            hand_landmarks = None
            is_ghost       = False

            # ── Case 1: live hand ──────────────────────────────────────────────
            if result.hand_landmarks:
                raw = result.hand_landmarks[0]

                # hand lock
                wrist_pos = (raw[0].x, raw[0].y)
                if not lock_set:
                    locked_wrist = wrist_pos
                    lock_set     = True
                    logger.log("LOCK_SET", hand_detected=True, notes=f"wrist=({wrist_pos[0]:.3f},{wrist_pos[1]:.3f})")
                    print("[AERO] Hand lock SET")
                else:
                    dist = math.sqrt((wrist_pos[0]-locked_wrist[0])**2 + (wrist_pos[1]-locked_wrist[1])**2)
                    if dist > HAND_LOCK_DISTANCE:
                        # different hand — ignore
                        raw = None

                if raw is not None:
                    hand_landmarks        = raw
                    last_good_landmarks   = raw
                    locked_wrist          = wrist_pos   # follow normal presenter movement
                    ghost_frames_remaining = 10

                    if in_ghost_mode:
                        in_ghost_mode = False
                        logger.log("GHOST_END", hand_detected=True)

                    # green skeleton
                    for a, b in HAND_CONNECTIONS:
                        ax = int(raw[a].x * FRAME_WIDTH);  ay = int(raw[a].y * FRAME_HEIGHT)
                        bx = int(raw[b].x * FRAME_WIDTH);  by = int(raw[b].y * FRAME_HEIGHT)
                        cv2.line(frame, (ax,ay), (bx,by), (0,255,0), 2)

            # ── Case 2: ghost buffer ───────────────────────────────────────────
            elif last_good_landmarks is not None and ghost_frames_remaining > 0:
                ghost_frames_remaining -= 1
                hand_landmarks          = last_good_landmarks
                is_ghost                = True

                if not in_ghost_mode:
                    in_ghost_mode = True
                    logger.log("GHOST_START", hand_detected=True,notes=f"frames_left={ghost_frames_remaining}")

                # orange skeleton to show ghost mode
                for a, b in HAND_CONNECTIONS:
                    ax = int(hand_landmarks[a].x * FRAME_WIDTH)
                    ay = int(hand_landmarks[a].y * FRAME_HEIGHT)
                    bx = int(hand_landmarks[b].x * FRAME_WIDTH)
                    by = int(hand_landmarks[b].y * FRAME_HEIGHT)
                    cv2.line(frame, (ax,ay), (bx,by), (0,150,255), 2)

                cv2.putText(frame, "GHOST MODE", (10,90),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0,150,255), 2)

            # ── Case 3: no hand ────────────────────────────────────────────────
            else:
                if last_good_landmarks is not None:
                    logger.log("DROPOUT", notes="beyond ghost window")
                    in_ghost_mode = False
                last_good_landmarks    = None
                ghost_frames_remaining = 0

            # ── Gesture processing ─────────────────────────────────────────────
            if hand_landmarks and gesture_enabled:

                # ── Z-depth dynamic threshold scaling ─────────────────────────
                # World landmarks give real-world metric distances (metres).
                # Wrist-to-middle-MCP distance scales with presenter's distance from camera.
                # Far-away hands → smaller palm → looser thresholds (less travel needed).
                # Close hands   → bigger palm  → stricter thresholds (more travel needed).
                # apply_distance_scale() was written and waiting in utils.py; this wires it.
                try:
                    if result.hand_world_landmarks:
                        w      = result.hand_world_landmarks[0]
                        wrist  = w[0]
                        mid_mcp = w[9]
                        palm_m = math.sqrt(
                            (wrist.x - mid_mcp.x)**2 +
                            (wrist.y - mid_mcp.y)**2 +
                            (wrist.z - mid_mcp.z)**2
                        )
                        scale = max(0.5, min(2.0, palm_m / PALM_REFERENCE_METRES))
                        swipe_detector.apply_distance_scale(
                            scale,
                            SWIPE_VELOCITY, SWIPE_DISPLACEMENT,
                            0.032, 0.12
                        )
                except Exception as e:
                    if debug_mode:
                        print(f"[DEBUG] World landmarks skipped: {e}")

                # ── Erase gesture (index+middle+ring up, pinky down) ──────────
                # Checked before gesture classification so it can fire even while
                # transitioning between other gestures.  30-frame cooldown (~1s)
                # prevents accidental repeated clears.
                if _erase_cooldown == 0 and _is_erase_gesture(hand_landmarks):
                    overlay.clear_annotations()
                    _erase_cooldown = 30
                    logger.log("ERASE_GESTURE", True)
                if _erase_cooldown > 0:
                    _erase_cooldown -= 1
                orientation, confidence = orient_tracker.update_orientation(hand_landmarks)

                if orientation != last_orientation:
                    # ── Master Reset on orientation change ────────────────────
                    # "I don't care what the layers think — we have a new hand
                    #  type now. Start the math from zero." (Fix 3)
                    # Everything stops for 1 frame so the new hand settles before
                    # any swipe detection can fire.
                    _swipe_gesture_frames = 0   # drop confirmation counter
                    swipe_detector.reset()       # wipe velocity/position buffer
                    move_validator.reset()       # wipe wrist movement history
                    lock_set      = False        # force re-acquire in Case 1
                    locked_wrist  = None
                    last_orientation = orientation
                    if debug_mode:
                        print(f"[AERO] Orientation → {orientation.upper()} | hard reset")
                # Classify gesture — fist checked first to prevent misfire during
                # transitions, then SWIPE/LASER/NONE
                gesture = classify_hand(hand_landmarks)

                _back_hand_mode = (orientation == "back-hand")
                _orient_conf    = confidence

                if mobile_server:
                    mobile_server.update_system_status(
                        gesture_enabled    = gesture_enabled,
                        hand_detected      = True,
                        back_hand_mode     = _back_hand_mode,
                        orientation_confidence = _orient_conf,
                    )

                # ── SWIPE ──────────────────────────────────────────────────────
                if gesture == "SWIPE":
                    # 3-frame confirmation: gesture must be stable before detector fires.
                    # Stops false triggers from passing through SWIPE on the way to LASER.
                    if not is_ghost:   # ghost frames are visuals only — not confirmation data
                        _swipe_gesture_frames += 1

                    ix = hand_landmarks[8].x
                    iy = hand_landmarks[8].y
                    mx = hand_landmarks[12].x
                    my = hand_landmarks[12].y
                    mid_x = (ix + mx) / 2
                    mid_y = (iy + my) / 2

                    hand_moving = move_validator.check_if_hand_is_moving(
                        hand_landmarks[0].x, hand_landmarks[0].y
                    )

                    direction = None
                    if _swipe_gesture_frames >= SWIPE_CONFIRM_FRAMES:
                        # ── Fix 1: Ghost hand = visuals only, not physics ─────
                        # Ghost landmarks are the last REAL position frozen in
                        # place.  If the real hand reappears even 5px away that
                        # looks like a massive instant velocity to the detector
                        # and fires a wrong swipe.  Only real frames do physics.
                        if not is_ghost:
                            direction = swipe_detector.update(mid_x, mid_y,orientation=orientation,is_ghost=is_ghost)
                            # backup: catch fast short swipes
                            if not direction:
                                direction = swipe_detector.trigger_if_displaced(
                                    threshold=0.08, orientation=orientation
                                )
                    else:
                        # still accumulating — draw a subtle confirmation counter
                        cv2.putText(frame,
                                    f"({_swipe_gesture_frames}/{SWIPE_CONFIRM_FRAMES})",
                                    (10, FRAME_HEIGHT - 55),
                                    cv2.FONT_HERSHEY_SIMPLEX, 0.4, (100, 100, 255), 1)

                    # end any open annotation stroke — leaving LASER for SWIPE
                    if _ann_was_laser:
                        overlay.end_stroke()
                    _ann_was_laser = False

                    if direction and hand_moving:
                        # invert for back-hand (physically mirrored)
                        # confidence gate: skip inversion when hand is edge-on (ambiguous)
                        if _back_hand_mode and confidence >= 0.40:
                            direction = "LEFT" if direction == "RIGHT" else "RIGHT"
                            conf_pct  = int(confidence * 100)
                            cv2.putText(frame, f"BACK-HAND ({conf_pct}%)", (10,120),
                                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255,150,0), 2)

                        if direction == "RIGHT":
                            next_slide()
                            if mobile_server: mobile_server.execute_command("NEXT_NOTIFY")  # immediate phone update
                            logger.log("SWIPE_RIGHT", True, orientation, confidence,"SWIPE", is_ghost)
                            backhand_logger.log_swipe(orientation, confidence, "RIGHT",is_ghost_frame_active=is_ghost)
                            # set persistent feedback banner — visible for ~1.5s
                            _feedback_text   = ">>> NEXT"
                            _feedback_color  = (0, 255, 0)
                            _feedback_frames = FEEDBACK_HOLD_FRAMES
                        elif direction == "LEFT":
                            prev_slide()
                            if mobile_server: mobile_server.execute_command("PREV_NOTIFY")  # immediate phone update
                            logger.log("SWIPE_LEFT", True, orientation, confidence,"SWIPE", is_ghost)
                            backhand_logger.log_swipe(orientation, confidence, "LEFT",is_ghost_frame_active=is_ghost)
                            _feedback_text   = "<<< PREV"
                            _feedback_color  = (0, 100, 255)
                            _feedback_frames = FEEDBACK_HOLD_FRAMES

                    overlay.hide()
                    laser_persist = 0
                elif gesture == "LASER" and laser_enabled:
                    swipe_detector.reset()
                    _swipe_gesture_frames = 0   # leaving SWIPE — reset confirmation counter
                    laser_persist = 5   # keep dot visible for 5 frames after pose changes

                    tip_x = laser_smooth_x.update(hand_landmarks[8].x)
                    tip_y = laser_smooth_y.update(hand_landmarks[8].y)
                    sx, sy = map_to_screen(tip_x, tip_y, SCREEN_W, SCREEN_H)
                    overlay.move_and_show(sx, sy)   # single refresh — no flicker
                    logger.log("LASER_FRAME", True, orientation, confidence, "LASER", is_ghost)

                    # ── Annotation drawing ─────────────────────────────────────
                    # When draw mode is ON, the laser tip leaves a persistent yellow
                    # stroke on the overlay.  begin_stroke starts a new stroke the
                    # first frame LASER appears; extend_stroke appends on subsequent
                    # frames.  end_stroke is called when LASER is lost.
                    if draw_mode:
                        if not _ann_was_laser:
                            overlay.begin_stroke(sx, sy)
                        else:
                            overlay.extend_stroke(sx, sy)
                    _ann_was_laser = True

                # ── FIST — no action (zoom removed) ────────────────────────────
                elif gesture == "FIST":
                    _swipe_gesture_frames = 0
                    if _ann_was_laser:
                        overlay.end_stroke()
                    _ann_was_laser = False
                    # fist is a rest pose — hide laser, log for accuracy data
                    if laser_persist > 0:
                        laser_persist -= 1
                    else:
                        overlay.hide()

                # ── NO GESTURE / NONE ───────────────────────────────────────────
                else:
                    _swipe_gesture_frames = 0
                    if _ann_was_laser:
                        overlay.end_stroke()
                    _ann_was_laser = False
                    if laser_persist > 0:
                        laser_persist -= 1
                    else:
                        overlay.hide()

            else:
                # no hand — reset all gesture state so it is clean on re-detection
                _swipe_gesture_frames = 0
                if _ann_was_laser:
                    overlay.end_stroke()
                _ann_was_laser   = False
                swipe_detector.reset()    # clears waiting_for_centre + cooldown
                move_validator.reset()    # clears stale wrist history
                orient_tracker.frames_counted_toward_switch = 0  # reset hysteresis
                last_orientation = "palm"  # reset so next detection does not fire master-reset
                if mobile_server:
                    mobile_server.update_system_status(hand_detected=False)
                if laser_persist > 0:
                    laser_persist -= 1
                else:
                    overlay.hide()

            # ── Tilt pointer (phone gyro moves laser) ─────────────────────────
            if tilt_active and laser_enabled:
                sx = int(tilt_x * SCREEN_W)
                sy = int(tilt_y * SCREEN_H)
                overlay.move_and_show(sx, sy)   # single refresh — no flicker
                # Visual feedback: log tilt control is active so presenter knows
                if debug_mode:
                    cv2.putText(frame, "[TILT CONTROL ACTIVE]", (10, 150),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 165, 0), 2)

            # ── HUD ───────────────────────────────────────────────────────────
            if show_hud:
                lines = [
                    f"FPS: {fps:.1f}",
                    f"Gestures: {'ON' if gesture_enabled else 'OFF'}",
                    f"Laser:    {'ON' if laser_enabled    else 'OFF'}",
                    f"Voice:    {'ON' if voice_enabled    else 'OFF'}",
                    f"Draw:     {'ON — tracing strokes' if draw_mode else 'OFF'}",
                    f"Hand:     {'GHOST' if is_ghost else ('YES' if hand_landmarks else 'NO')}",
                ]
                if debug_mode and hand_landmarks:
                    lines.append(f"Orient: {_back_hand_mode and 'back' or 'palm'} {_orient_conf:.2f}")
                y = 30
                for txt in lines:
                    cv2.putText(frame, txt, (10, y),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0,255,0), 2)
                    y += 22

            # ── QR code overlay (shown for first 6 seconds after server starts) ─
            if qr_show_frames > 0 and qr_overlay is not None:
                qh, qw = qr_overlay.shape[:2]
                pad = 8
                # place bottom-right corner, 8px from edge, with a dark background pad
                x1 = FRAME_WIDTH  - qw - pad
                y1 = FRAME_HEIGHT - qh - pad
                # draw dark background pad so QR is readable on any slide colour
                cv2.rectangle(frame, (x1 - 4, y1 - 4),(x1 + qw + 4, y1 + qh + 20), (20, 20, 20), -1)
                frame[y1:y1+qh, x1:x1+qw] = qr_overlay
                cv2.putText(frame, "Scan to connect",
                            (x1, y1 + qh + 14),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.4, (180, 220, 255), 1)
                qr_show_frames -= 1

            # persistent gesture feedback banner — large text stays on screen for
            # ~1.5s after each swipe fires so it's clearly visible during testing.
            # Shadow pass first for readability over any slide background colour.
            if _feedback_frames > 0:
                cv2.putText(frame, _feedback_text, (18, 78),
                            cv2.FONT_HERSHEY_SIMPLEX, 1.8, (0, 0, 0), 6)
                cv2.putText(frame, _feedback_text, (20, 80),
                            cv2.FONT_HERSHEY_SIMPLEX, 1.8, _feedback_color, 4)
                _feedback_frames -= 1

            cv2.imshow("Aero Present - Camera", frame)

            # ── Keyboard controls ─────────────────────────────────────────────
            key = cv2.waitKey(1) & 0xFF

            if key == ord('q'):
                break

            elif key == ord('g'):
                gesture_enabled = not gesture_enabled
                print(f"[AERO] Gestures {'ON' if gesture_enabled else 'OFF'}")
                overlay.update_hud(gesture_enabled, laser_enabled, voice_enabled)
                logger.log(f"GESTURE_{'ON' if gesture_enabled else 'OFF'}")

            elif key == ord('l'):
                laser_enabled = not laser_enabled
                print(f"[AERO] Laser {'ON' if laser_enabled else 'OFF'}")
                if not laser_enabled:
                    overlay.hide()
                overlay.update_hud(gesture_enabled, laser_enabled, voice_enabled)
                logger.log(f"LASER_{'ON' if laser_enabled else 'OFF'}")

            elif key == ord('v'):
                if voice:
                    voice_enabled = not voice_enabled
                    if voice_enabled:
                        voice.enabled = True
                        voice.start_listening()
                        print("[AERO] Voice ON")
                    else:
                        voice.enabled = False
                        voice.stop_listening()
                        print("[AERO] Voice OFF")
                    overlay.update_hud(gesture_enabled, laser_enabled, voice_enabled)
                else:
                    print("[AERO] Voice not available")

            elif key == ord('h'):
                show_hud = not show_hud
                overlay.toggle_hud()

            elif key == ord('d'):
                debug_mode = not debug_mode

            elif key == ord('r'):
                locked_wrist = None
                lock_set     = False
                print("[AERO] Hand lock RESET")
                logger.log("LOCK_RESET", notes="manual R key")

            elif key == ord('z'):
                # zoom removed — key reserved for future use
                pass

            elif key in (ord('c'), ord('C')):
                print("[AERO] Recalibrating...")
                logger.log("RECALIBRATE_START")
                calib_d = make_detector(det_conf, trk_conf, vision.RunningMode.IMAGE)
                det_conf, trk_conf = calibrate_confidence(calib_d, cap)
                calib_d.close()
                detector.close()
                detector = make_detector(det_conf, trk_conf, vision.RunningMode.VIDEO)
                logger.log("RECALIBRATE_END", notes=f"det={det_conf} trk={trk_conf}")
                print("[AERO] Recalibration done\n")

            elif key == ord('a'):
                # toggle annotation draw mode
                draw_mode = not draw_mode
                overlay.set_draw_mode(draw_mode)
                logger.log(f"DRAW_MODE_{'ON' if draw_mode else 'OFF'}")

            elif key == ord('5'):
                # F5 — start slideshow from current slide
                ppt_start_slideshow()
                logger.log("SHORTCUT_F5")

            elif key in (ord('e'), ord('E')):
                # Escape — exit slideshow, return to edit view
                ppt_exit_slideshow()
                logger.log("SHORTCUT_ESC")

            elif key == ord('p'):
                # Ctrl+P — activate PowerPoint built-in pen/pointer mode during slideshow
                ppt_pointer_mode()
                logger.log("SHORTCUT_CTRLP")

    finally:
        print("\n[AERO] Shutting down...")
        logger.log("SESSION_END", notes="clean shutdown")
        logger.close()
        backhand_logger.close()
        if voice:
            voice.stop_listening()
        cap.release()
        cv2.destroyAllWindows()
        overlay.destroy()
        detector.close()
        print("[AERO] Goodbye.")


if __name__ == '__main__':
    main()
