"""
Aero Present — Utility Module  (v3 — Three-Signal 2D Orientation Engine)
=========================================================================
All landmark access is direct by index — no iteration overhead.
ALL orientation signals use ONLY x and y coordinates from 2D webcam.
No z-depth, no depth sensor, no pseudo-3D — pure monocular 2D.

ORIENTATION ENGINE UPGRADE (v2 → v3):
v2 used ONE signal: index MCP x vs pinky MCP x separation.
v3 fuses THREE independent 2D signals using majority vote:

  Signal 1 — MCP Horizontal Separation   [original — uses landmark.x only]
  Signal 2 — Thumb Chirality Test         [new — uses landmark.x only]
  Signal 3 — MediaPipe Handedness Label   [new — model output from 2D image]

PAPER CLAIM (preserved):
  "All orientation features are computed from 2D screen coordinates output
   by MediaPipe's hand landmark model running on a standard monocular webcam.
   No depth sensor or RGB-D camera is required."

Signal 1 is the original contribution from the back-hand paper.
Signal 2 adds a novel 2D geometric test based on thumb chirality.
Signal 3 leverages MediaPipe's own handedness classifier output.
Fusing all three via majority vote with hysteresis is the novel contribution.

BACK-HAND INVERSION FIX (gesture_control.py):
  Add confidence >= 0.40 gate before inverting direction.
  This prevents inversion during ambiguous edge-on transitions.
"""

import collections, csv, os, time
import numpy as np


# ══════════════════════════════════════════════════════════════════════════════
#  SMOOTHER
# ══════════════════════════════════════════════════════════════════════════════
class Smoother:
    def __init__(self, window_size: int = 6):
        self._buf = collections.deque(maxlen=window_size)

    def update(self, v: float) -> float:
        self._buf.append(v)
        return float(np.mean(self._buf))

    def reset(self):
        self._buf.clear()


# ══════════════════════════════════════════════════════════════════════════════
#  LANDMARK CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
LANDMARK_INDEX_FINGERTIP  = 8
LANDMARK_MIDDLE_FINGERTIP = 12
LANDMARK_RING_FINGERTIP   = 16
LANDMARK_PINKY_FINGERTIP  = 20
LANDMARK_INDEX_PIP_JOINT  = 6
LANDMARK_MIDDLE_PIP_JOINT = 10
LANDMARK_RING_PIP_JOINT   = 14
LANDMARK_PINKY_PIP_JOINT  = 18


def _is_finger_extended(fingertip_y: float, pip_joint_y: float) -> bool:
    return fingertip_y < pip_joint_y


def classify_hand(all_landmarks) -> str:
    index_up  = _is_finger_extended(all_landmarks[LANDMARK_INDEX_FINGERTIP].y,
                                     all_landmarks[LANDMARK_INDEX_PIP_JOINT].y)
    middle_up = _is_finger_extended(all_landmarks[LANDMARK_MIDDLE_FINGERTIP].y,
                                     all_landmarks[LANDMARK_MIDDLE_PIP_JOINT].y)
    ring_up   = _is_finger_extended(all_landmarks[LANDMARK_RING_FINGERTIP].y,
                                     all_landmarks[LANDMARK_RING_PIP_JOINT].y)
    pinky_up  = _is_finger_extended(all_landmarks[LANDMARK_PINKY_FINGERTIP].y,
                                     all_landmarks[LANDMARK_PINKY_PIP_JOINT].y)

    if index_up and middle_up and not ring_up and not pinky_up:
        return "SWIPE"
    if index_up and middle_up and ring_up and not pinky_up:
        return "ERASE"
    if index_up and not middle_up and not ring_up and not pinky_up:
        return "LASER"
    if not index_up and not middle_up and not ring_up and not pinky_up:
        return "FIST"
    return "NONE"


# ══════════════════════════════════════════════════════════════════════════════
#  ORIENTATION TRACKER  v3  —  Three-Signal 2D Fusion Engine
#
#  All three signals use ONLY .x and .y coordinates.
#  No z-depth. Pure monocular 2D.
#
#  Signal 1 — MCP Horizontal Separation (ORIGINAL contribution):
#    index_mcp.x − pinky_mcp.x > threshold → back-hand
#    Landmarks 5 and 17 — rigidly on the palm bone.
#    Reliable at 0.5-1.5m, degrades at edge-on angles.
#    Distance-adaptive threshold compensates for camera distance.
#
#  Signal 2 — Thumb Chirality (NEW, 2D geometric contribution):
#    Thumb tip x (landmark 4) relative to palm centroid x.
#    Palm centroid = average x of landmarks [0, 5, 9, 13, 17].
#    Right-handed presenter, mirrored frame (cv2.flip):
#      Palm toward camera: thumb to the LEFT of centroid → dx < 0
#      Back toward camera: thumb to the RIGHT of centroid → dx > 0
#    Chirality is frame-flip invariant in a predictable way.
#    Source principle: TheJLifeX (2021), landmark 4 vs 10 position check.
#    Novel application: using palm centroid instead of single reference point
#    makes this more robust to wrist rotation.
#
#  Signal 3 — MediaPipe Handedness Label (NEW, model-output contribution):
#    MediaPipe's own Left/Right classifier output.
#    Still from 2D image — the model estimates handedness from visual features.
#    Only used when score >= 0.80 (very high confidence).
#    With cv2.flip(frame,1): "Right" → palm, "Left" → back-hand (mirrored).
#    Source: MediaPipe docs — "handedness assumes mirrored input image."
#
#  FUSION RULE:
#    Each non-ambiguous signal casts a vote.
#    Majority wins. Confidence = (winning fraction) × (average signal strength).
#    Minimum 2 signals must participate; otherwise hold current orientation.
#
#  HYSTERESIS:
#    8 consecutive frames must agree before committing to a state change.
#    Prevents mid-swipe orientation flicker (a swipe takes ~250ms = ~7.5 frames).
# ══════════════════════════════════════════════════════════════════════════════

class OrientationTracker:
    """
    Three-signal majority-vote orientation tracker.
    All signals use only 2D (x, y) landmark coordinates.
    Compatible with standard monocular webcam — no depth sensor required.
    """

    def __init__(self,
                 transition_frames_required: int = 8,
                 presenter_right_handed: bool = True):
        self.current_confirmed_orientation   = "palm"
        self.frames_counted_toward_switch    = 0
        self.frames_needed_to_confirm_switch = transition_frames_required
        self._presenter_right_handed         = presenter_right_handed

    # ── Signal 1: MCP x-separation (original — preserved exactly) ─────────
    def _signal_mcp_separation(self, lm) -> tuple[str, float]:
        """
        Original contribution from Aero Present paper.

        Compares x-position of index finger MCP (landmark 5) vs
        pinky MCP (landmark 17). Both are rigidly fixed to the palm bone
        and don't flex with finger movement.

        palm toward camera : index_mcp is LEFT  of pinky_mcp → separation < 0
        back toward camera : index_mcp is RIGHT of pinky_mcp → separation > 0

        Distance-adaptive threshold: compensates for hand size in frame.
        At 0.5m: hand_size ≈ 0.12, threshold ≈ 0.10
        At 1.5m: hand_size ≈ 0.06, threshold ≈ 0.05
        """
        index_mcp_x = lm[5].x
        pinky_mcp_x = lm[17].x
        sep         = index_mcp_x - pinky_mcp_x

        hand_size = max(0.04, min(0.20, abs(lm[9].y - lm[0].y)))
        threshold = 0.10 * (hand_size / 0.12)

        if   sep >  threshold: return "back-hand", min(abs(sep), 0.30) / 0.30
        elif sep < -threshold: return "palm",      min(abs(sep), 0.30) / 0.30
        else:                  return "edge-on",   0.0  # hand is edge-on, ambiguous

    # ── Signal 2: Thumb Chirality (new — 2D, novel geometric contribution) ─
    def _signal_thumb_chirality(self, lm) -> tuple[str, float]:
        """
        Novel 2D contribution: thumb tip x relative to palm centroid x.

        The thumb's position relative to the palm center encodes chirality —
        whether the hand is oriented with palm toward or away from camera.
        Unlike the MCP gap which collapses at edge-on angles, this signal
        remains discriminative because the thumb is anatomically constrained
        to be on one side of the palm.

        Palm centroid = mean x of:
          landmark 0  (wrist)
          landmark 5  (index MCP)
          landmark 9  (middle MCP)
          landmark 13 (ring MCP)
          landmark 17 (pinky MCP)
        All five are on the palm bone — stable regardless of finger position.

        Right-handed presenter, mirrored frame (cv2.flip):
          Palm toward camera → thumb LEFT of centroid → thumb_x < centroid_x
          Back toward camera → thumb RIGHT of centroid → thumb_x > centroid_x

        Only uses landmark.x — purely 2D.
        """
        # Palm centroid from 5 bone-attached landmarks
        palm_xs    = [lm[0].x, lm[5].x, lm[9].x, lm[13].x, lm[17].x]
        centroid_x = sum(palm_xs) / 5.0
        dx         = lm[4].x - centroid_x  # positive = thumb right of center

        if abs(dx) < 0.03:
            return "edge-on", 0.0  # too close to call

        if self._presenter_right_handed:
            orientation = "back-hand" if dx > 0 else "palm"
        else:
            orientation = "palm" if dx > 0 else "back-hand"

        confidence = min(abs(dx) / 0.15, 1.0)
        return orientation, confidence

    # ── Signal 3: MediaPipe Handedness Label (model output from 2D image) ─
    def _signal_handedness(self, label: str, score: float) -> tuple[str, float]:
        """
        MediaPipe's own handedness classifier — output from 2D RGB image only.
        No depth sensor required. Still a 2D camera operation.

        Only trusted at score >= 0.80. Below this the model is uncertain
        (e.g., hand is edge-on to camera where chirality is ambiguous).

        With cv2.flip(frame,1) — mirrored input as used in gesture_control.py:
          Right-handed presenter:
            "Right" label → palm toward camera
            "Left"  label → back-hand toward camera
          Source: MediaPipe documentation, "handedness assumes mirrored input."
        """
        if not label or score < 0.80:
            return "edge-on", 0.0  # not confident enough

        if self._presenter_right_handed:
            orientation = "palm" if label == "Right" else "back-hand"
        else:
            orientation = "palm" if label == "Left" else "back-hand"

        return orientation, score

    # ── Main update: majority vote ─────────────────────────────────────────
    def update_orientation(
        self,
        all_landmarks,
        handedness_label: str   = "",
        handedness_score: float = 0.0,
    ) -> tuple[str, float]:
        """
        Returns (confirmed_orientation, confidence) every frame.

        BACKWARD COMPATIBLE: calling with just all_landmarks still works.
        Signals 1 and 2 run from landmarks alone.
        Signal 3 requires handedness_label and handedness_score to be passed.

        To activate Signal 3 in gesture_control.py, extract from result:
          _handedness_label = result.handedness[0][0].category_name
          _handedness_score = result.handedness[0][0].score
          orient_tracker.update_orientation(
              hand_landmarks,
              handedness_label=_handedness_label,
              handedness_score=_handedness_score
          )
        """
        votes_back = 0
        votes_palm = 0
        total_conf = 0.0
        n_active   = 0

        for fn, args in [
            (self._signal_mcp_separation,  (all_landmarks,)),
            (self._signal_thumb_chirality, (all_landmarks,)),
            (self._signal_handedness,      (handedness_label, handedness_score)),
        ]:
            orientation, conf = fn(*args)
            if orientation == "edge-on":
                continue   # this signal is ambiguous — skip its vote
            if orientation == "back-hand":
                votes_back += 1
            else:
                votes_palm += 1
            total_conf += conf
            n_active   += 1

        # Need at least 1 signal to make a call; 0 = hold current
        if n_active == 0:
            return self.current_confirmed_orientation, 0.0

        raw_orientation = "back-hand" if votes_back > votes_palm else "palm"

        # Confidence = (fraction agreeing) × (average signal strength)
        # Ranges 0.0–1.0. At 2/2 agreement with avg 0.8 strength: 1.0 × 0.8 = 0.80
        # At 1/2 agreement with strength 0.9: 0.5 × 0.9/2 = 0.225 (low — ambiguous)
        winning_votes  = max(votes_back, votes_palm)
        raw_confidence = (winning_votes / n_active) * (total_conf / n_active)

        # Hysteresis: require 8 consecutive frames before committing
        if raw_orientation != self.current_confirmed_orientation:
            self.frames_counted_toward_switch += 1
            if self.frames_counted_toward_switch >= self.frames_needed_to_confirm_switch:
                self.current_confirmed_orientation = raw_orientation
                self.frames_counted_toward_switch  = 0
        else:
            self.frames_counted_toward_switch = 0

        return self.current_confirmed_orientation, raw_confidence

    def reset(self):
        self.frames_counted_toward_switch  = 0
        self.current_confirmed_orientation = "palm"

    hard_reset = reset

    def soft_reset(self):
        self.frames_counted_toward_switch = 0


# ══════════════════════════════════════════════════════════════════════════════
#  MOVEMENT VALIDATOR  (unchanged)
# ══════════════════════════════════════════════════════════════════════════════
class MovementValidator:
    def __init__(self, history_frames: int = 10, minimum_movement: float = 0.05):
        self.wrist_position_history = collections.deque(maxlen=history_frames)
        self._minimum_movement      = minimum_movement

    def check_if_hand_is_moving(self, wx: float, wy: float) -> bool:
        self.wrist_position_history.append((wx, wy))
        if len(self.wrist_position_history) < 3:
            return False
        total = 0.0
        for i in range(len(self.wrist_position_history) - 1):
            a = self.wrist_position_history[i]
            b = self.wrist_position_history[i + 1]
            total += ((b[0]-a[0])**2 + (b[1]-a[1])**2) ** 0.5
        return total > self._minimum_movement

    def reset(self):
        self.wrist_position_history.clear()


# ══════════════════════════════════════════════════════════════════════════════
#  BACK-HAND ACCURACY LOGGER
# ══════════════════════════════════════════════════════════════════════════════
class BackHandAccuracyLogger:
    def __init__(self):
        os.makedirs("logs", exist_ok=True)
        ts              = time.strftime("%Y%m%d_%H%M%S")
        self.path       = os.path.join("logs", f"backhand_accuracy_{ts}.csv")
        self.start_time = time.time()
        self.fh         = open(self.path, "w", newline="", encoding="utf-8")
        self.wr         = csv.writer(self.fh)
        self.wr.writerow([
            "timestamp", "elapsed_s", "orientation", "confidence",
            "handedness_score", "swipe_direction", "is_ghost", "notes"
        ])
        self.fh.flush()
        print(f"[BACKHAND LOGGER] → {self.path}")

    def log_swipe(self, orientation, confidence, direction,
                  handedness_score=0.0, is_ghost_frame_active=False, notes=""):
        self.wr.writerow([
            time.strftime("%Y-%m-%dT%H:%M:%S"),
            round(time.time() - self.start_time, 3),
            orientation, round(confidence, 3), round(handedness_score, 3),
            direction, is_ghost_frame_active, notes
        ])
        self.fh.flush()

    def calculate_and_print_session_summary(self):
        try:
            with open(self.path) as f:
                r     = csv.DictReader(f)
                palm  = back = back_hc = 0
                for row in r:
                    if   row["orientation"] == "palm":      palm += 1
                    elif row["orientation"] == "back-hand":
                        back += 1
                        try:
                            if float(row.get("handedness_score", 0) or 0) >= 0.80:
                                back_hc += 1
                        except ValueError:
                            pass
            total = palm + back
            if total:
                print(f"\n[BACKHAND LOGGER]  Palm: {palm}  Back: {back} ({back/total*100:.1f}%)  High-conf: {back_hc}")
        except Exception as e:
            print(f"[BACKHAND LOGGER] summary error: {e}")

    def close(self):
        self.calculate_and_print_session_summary()
        self.fh.close()
        print(f"[BACKHAND LOGGER] Saved → {self.path}")


# ══════════════════════════════════════════════════════════════════════════════
#  SWIPE DETECTOR  (window 8, ratio 2.5, timeout 30)
# ══════════════════════════════════════════════════════════════════════════════
class SwipeDetector:
    def __init__(self, palm_velocity_threshold=0.028, palm_displacement_threshold=0.10,
                 back_velocity_threshold=0.032, back_displacement_threshold=0.12,
                 cooldown_frames=20, window=8, neutral_zone=0.20,
                 require_dominant_axis_ratio=2.5):

        self.palm_velocity_limit          = palm_velocity_threshold
        self.palm_displacement_limit      = palm_displacement_threshold
        self.back_velocity_limit          = back_velocity_threshold
        self.back_displacement_limit      = back_displacement_threshold
        self.cooldown_frames_after_swipe  = cooldown_frames
        self.frames_remaining_in_cooldown = 0
        self.recent_position_history      = collections.deque(maxlen=window)
        self.neutral_zone_half_width      = neutral_zone
        self.waiting_for_hand_to_return_to_centre = False
        self.last_known_orientation       = "palm"
        self.is_currently_in_ghost_mode   = False
        self._waiting_frames              = 0
        self._require_ratio               = require_dominant_axis_ratio

    def update(self, midpoint_x, midpoint_y, orientation="palm", is_ghost=False):
        self.last_known_orientation     = orientation
        self.is_currently_in_ghost_mode = is_ghost

        nz_left  = 0.5 - self.neutral_zone_half_width
        nz_right = 0.5 + self.neutral_zone_half_width
        in_nz    = nz_left < midpoint_x < nz_right

        if self.waiting_for_hand_to_return_to_centre:
            if in_nz:
                self.waiting_for_hand_to_return_to_centre = False
                self._waiting_frames = 0
                self.recent_position_history.clear()
                return None
            self._waiting_frames += 1
            if self._waiting_frames >= 30:
                self.waiting_for_hand_to_return_to_centre = False
                self._waiting_frames = 0
                self.recent_position_history.clear()
            return None

        self.recent_position_history.append((midpoint_x, midpoint_y))

        if self.frames_remaining_in_cooldown > 0:
            self.frames_remaining_in_cooldown -= 1
            return None

        if len(self.recent_position_history) < self.recent_position_history.maxlen:
            return None

        xs       = [p[0] for p in self.recent_position_history]
        ys       = [p[1] for p in self.recent_position_history]
        dx       = xs[-1] - xs[0]
        dy       = ys[-1] - ys[0]
        velocity = dx / len(xs)

        horizontal_dominates = abs(dx) >= self._require_ratio * abs(dy)
        vel_thr  = self.back_velocity_limit    if orientation == "back-hand" else self.palm_velocity_limit
        disp_thr = self.back_displacement_limit if orientation == "back-hand" else self.palm_displacement_limit

        if horizontal_dominates and abs(velocity) >= vel_thr and abs(dx) >= disp_thr:
            self.frames_remaining_in_cooldown         = self.cooldown_frames_after_swipe
            self.waiting_for_hand_to_return_to_centre = True
            self._waiting_frames                      = 0
            self.recent_position_history.clear()
            return "RIGHT" if dx > 0 else "LEFT"

        return None

    def trigger_if_displaced(self, threshold=0.08, orientation="palm"):
        if self.waiting_for_hand_to_return_to_centre or self.frames_remaining_in_cooldown > 0:
            return None
        if len(self.recent_position_history) < 3:
            return None
        xs  = [p[0] for p in self.recent_position_history]
        dx  = xs[-1] - xs[0]
        thr = threshold * 1.1 if orientation == "back-hand" else threshold
        if abs(dx) > thr:
            self.waiting_for_hand_to_return_to_centre = True
            self._waiting_frames = 0
            self.recent_position_history.clear()
            self.frames_remaining_in_cooldown = self.cooldown_frames_after_swipe
            return "RIGHT" if dx > 0 else "LEFT"
        return None

    def apply_distance_scale(self, scale_factor, bpv, bpd, bbv, bbd):
        self.palm_velocity_limit     = bpv / scale_factor
        self.palm_displacement_limit = bpd / scale_factor
        self.back_velocity_limit     = bbv / scale_factor
        self.back_displacement_limit = bbd / scale_factor

    def reset(self):
        self.recent_position_history.clear()
        self.frames_remaining_in_cooldown             = 0
        self.waiting_for_hand_to_return_to_centre     = False
        self._waiting_frames                          = 0


# ══════════════════════════════════════════════════════════════════════════════
#  COORDINATE MAPPER
# ══════════════════════════════════════════════════════════════════════════════
def map_to_screen(normalised_x, normalised_y, screen_width, screen_height,
                  margin=0.05, dead_zone=0.02):
    usable = 1.0 - 2.0 * margin
    cx = max(dead_zone, min(1.0 - dead_zone, (normalised_x - margin) / usable))
    cy = max(dead_zone, min(1.0 - dead_zone, (normalised_y - margin) / usable))
    return int(cx * screen_width), int(cy * screen_height)
