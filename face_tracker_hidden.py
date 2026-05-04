import requests
import cv2
import serial
import time
import threading
import os
import tempfile
import re
from difflib import SequenceMatcher
from serial.tools import list_ports

try:
    import mediapipe as mp
    MEDIAPIPE_AVAILABLE = True
except Exception:
    mp = None
    MEDIAPIPE_AVAILABLE = False

try:
    import pyttsx3
    PYTTSX3_AVAILABLE = True
except Exception:
    pyttsx3 = None
    PYTTSX3_AVAILABLE = False

try:
    import edge_tts
    import asyncio
    EDGE_TTS_AVAILABLE = True
except Exception:
    edge_tts = None
    asyncio = None
    EDGE_TTS_AVAILABLE = False

try:
    import pygame
    PYGAME_AVAILABLE = True
except Exception:
    pygame = None
    PYGAME_AVAILABLE = False

try:
    import speech_recognition as sr
    SPEECH_RECOGNITION_AVAILABLE = True
except Exception:
    sr = None
    SPEECH_RECOGNITION_AVAILABLE = False

SERVER_SCAN_URL = "http://127.0.0.1:5000/api/attendance/scan"
HTTP_TIMEOUT = 0.35
POST_SCAN_TO_SERVER = True
POST_SCAN_COOLDOWN = 1.0
TEST_BARCODE = "TEST123456"
BACKEND_ATTENDANCE_URL = "http://127.0.0.1:5000/api/attendance"
BACKEND_POLL_INTERVAL = 2.0
ENABLE_BACKEND_ATTENDANCE_WATCH = False

last_posted_barcode = ""
last_post_time = 0.0
last_backend_poll_time = 0.0
last_seen_backend_signature = None
recent_handled_signatures = {}
http_session = requests.Session()

SERIAL_PORT = "COM5"
BAUD_RATE = 115200
SERIAL_OPEN_DELAY = 0.5
RECONNECT_INTERVAL = 2.0
ENABLE_SERVO = True
READ_BARCODE_FROM_ESP32 = True

CAMERA_INDEX = 1
CAMERA_CANDIDATES = [1]
CAMERA_RECONNECT_DELAY = 0.5
CAMERA_OPEN_DELAY = 0.02
CAMERA_BACKENDS = [cv2.CAP_DSHOW]
SHOW_CAMERA = True
HEADLESS_STATUS_INTERVAL = 2.0

CAMERA_WIDTH = 640
CAMERA_HEIGHT = 360
PROCESS_WIDTH = 320
PROCESS_HEIGHT = 180
DISPLAY_WINDOW_WIDTH = 960
DISPLAY_WINDOW_HEIGHT = 540
WINDOW_NAME = "3DOF Face Tracker - Integrated Camera"

ENABLE_UNIFORM_DETECTION = True
UNIFORM_TEXT_COOLDOWN = 1.5
UNIFORM_MIN_SHIRT_BRIGHT_RATIO = 0.52
UNIFORM_MIN_GREEN_RATIO = 0.018
UNIFORM_MIN_GOLD_RATIO = 0.010
UNIFORM_MIN_COMBINED_SCORE = 0.58

ENABLE_FACE_WELCOME_AUDIO = True
FACE_AUDIO_MESSAGE = "Hello, Welcome to the Library Please scan a barcode to proceed"
FACE_AUDIO_RATE = 150

PREFER_EDGE_TTS = True
EDGE_TTS_VOICE = "en-US-GuyNeural"
EDGE_TTS_RATE = "-5%"
EDGE_TTS_PITCH = "-1Hz"
EDGE_TTS_VOLUME = "+0%"

ENABLE_VOICE_INTERACTION = True

VOICE_TRIGGER_WORDS = [
    "hello",
    "track"
]

VOICE_TRIGGER_ALIASES = [
    "hello", "helo", "halo", "hullo",
    "track", "trak", "trax", "trk", "trek", "traq"
]

VOICE_TRIGGER_FIRST_WORDS = [
    "hello"
]

VOICE_TRIGGER_SECOND_WORDS = [
    "track", "trak", "trax", "trk", "trek", "traq"
]

VOICE_TRIGGER_SIMILARITY = 0.72
VOICE_REPLY_TEXT = "Hello. Welcome to the library. How may I assist you?"
VOICE_FOLLOWUP_REGISTER_TEXT = "Please proceed to the administrator for registration, and wait for your barcode to be generated."

VOICE_REGISTER_WORDS = [
    "register",
    "registration",
    "register me",
    "registered",
    "registar",
    "registrar",
    "regster",
    "regis ter",
    "regitr",
    "registor",
    "rejister",
    "rejester",
    "ragister",
    "ragester",
    "rekister",
    "rigister",
    "rigester",
    "resister",
    "register now",
    "i want to register",
    "i want register",
    "can i register",
    "sign up",
    "signup",
    "sign-up",
    "sign me up",
    "signing up",
    "sinup",
    "sine up",
    "sayn up",
    "enroll",
    "enrol",
    "enroll me",
    "enrol me"
]

VOICE_WAIT_FOR_REGISTER_TIMEOUT = 8.0
VOICE_COMMAND_COOLDOWN = 1.0    
VOICE_LISTEN_TIMEOUT = 0.8
VOICE_PHRASE_TIME_LIMIT = 2.5

DETECT_EVERY_N_FRAMES = 3

NO_FACE_TIMEOUT = 0.5

INVERT_CAMERA = True
INVERT_MODE = 1
REVERSE_BASE = False
REVERSE_SHOULDER = False

HOME_BASE = 90
HOME_SHOULDER = 70
HOME_ELBOW = 150

BASE_MIN, BASE_MAX = 10, 170
SHOULDER_MIN, SHOULDER_MAX = 25, 155
ELBOW_MIN, ELBOW_MAX = 70, 175

CENTER_DEADZONE_X = 20
CENTER_DEADZONE_Y = 16
SIZE_DEADZONE = 14

FACE_CENTER_ALPHA = 0.34
FACE_SIZE_ALPHA = 0.22
SEARCH_MARGIN = 220

NO_FACE_HOME_ALPHA_BASE = 0.25
NO_FACE_HOME_ALPHA_SHOULDER = 0.25
NO_FACE_HOME_ALPHA_ELBOW = 0.18

BASE_FINE_KP = 0.13
SHOULDER_FINE_KP = 0.10
ELBOW_FINE_KP = 0.06
BASE_STEP_MIN = 0.7
BASE_STEP_MAX = 4.0
SHOULDER_STEP_MIN = 0.6
SHOULDER_STEP_MAX = 3.6
ELBOW_STEP_MIN = 0.5
ELBOW_STEP_MAX = 2.8

SERVO_SEND_INTERVAL = 0.03
ANGLE_SEND_DEADBAND = 3

OLED_SEND_INTERVAL = 0.05
OLED_EYE_DEADBAND = 1
OLED_X_MIN = -12
OLED_X_MAX = 12
OLED_Y_MIN = -4
OLED_Y_MAX = 4

base_angle = float(HOME_BASE)
shoulder_angle = float(HOME_SHOULDER)
elbow_angle = float(HOME_ELBOW)

target_base = float(HOME_BASE)
target_shoulder = float(HOME_SHOULDER)
target_elbow = float(HOME_ELBOW)

filtered_target_base = float(HOME_BASE)
filtered_target_shoulder = float(HOME_SHOULDER)
filtered_target_elbow = float(HOME_ELBOW)

last_sent = {"BASE": None, "SHOULDER": None, "ELBOW": None}
locked_face_center = None
locked_face_width = None
last_barcode = ""
last_serial_message = ""
frame_count = 0
last_serial_attempt = 0.0
last_send_time = 0.0
last_oled_send_time = 0.0
last_headless_status_time = 0.0
last_face_seen_time = 0.0
last_face_lost_time = 0.0
home_command_sent = False
last_detection_faces = []
fps_time = 0.0
fps_frames = 0
current_fps = 0.0
last_uniform_detect_time = 0.0
last_uniform_score = 0.0
last_uniform_detected = False
last_face_present = False
face_detected_since = 0.0

audio_busy = False
audio_lock = threading.Lock()
tts_thread = None

face_audio_armed = True
last_audio_play_time = 0.0
AUDIO_REARM_DELAY = 2.0
AUDIO_MIN_PLAY_GAP = 3.0
FACE_STABLE_SPEAK_DELAY = 0.8
FACE_LOST_REARM_DELAY = 1.5

oled_face_state = None
oled_talk_state = None
oled_eye_x = None
oled_eye_y = None

ESP32_KEYWORDS = ["cp210", "ch340", "usb serial", "silicon labs", "wch", "esp32", "uart", "ftdi"]

voice_thread = None
voice_running = False
last_voice_command_time = 0.0
waiting_for_register_command = False
wait_for_register_until = 0.0

ser = None


class CameraStream:
    def __init__(self, index=1):
        self.index = index
        self.cap = None
        self.frame = None
        self.lock = threading.Lock()
        self.running = False
        self.thread = None

    def _try_open_with_backend(self, backend):
        try:
            cap = cv2.VideoCapture(self.index, backend)
        except Exception:
            cap = None

        if not cap or not cap.isOpened():
            if cap:
                cap.release()
            return None

        try:
            cap.set(cv2.CAP_PROP_FRAME_WIDTH, CAMERA_WIDTH)
            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, CAMERA_HEIGHT)
            cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
            cap.set(cv2.CAP_PROP_FPS, 30)
            cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*"MJPG"))
        except Exception:
            pass

        time.sleep(CAMERA_OPEN_DELAY)

        ok, test_frame = cap.read()
        if not ok or test_frame is None:
            cap.release()
            return None

        return cap

    def _open(self):
        global last_serial_message

        for backend in CAMERA_BACKENDS:
            cap = self._try_open_with_backend(backend)
            if cap is not None:
                last_serial_message = f"CAMERA CONNECTED: index={self.index} backend={backend}"
                return cap

        last_serial_message = f"CAMERA ERROR: integrated camera index {self.index} not opened"
        return None

    def start(self):
        self.cap = self._open()
        if self.cap is None:
            raise RuntimeError(
                f"Integrated camera open failed on index {self.index}. "
                "If your laptop maps the integrated camera differently, change CAMERA_INDEX to 0."
            )
        self.running = True
        self.thread = threading.Thread(target=self._reader, daemon=True)
        self.thread.start()
        return self

    def _reader(self):
        while self.running:
            if self.cap is None:
                time.sleep(0.1)
                continue

            ok, frame = self.cap.read()
            if not ok or frame is None:
                try:
                    self.cap.release()
                except Exception:
                    pass
                time.sleep(CAMERA_RECONNECT_DELAY)
                self.cap = self._open()
                continue

            if INVERT_CAMERA:
                frame = cv2.flip(frame, INVERT_MODE)

            if frame.shape[1] != PROCESS_WIDTH or frame.shape[0] != PROCESS_HEIGHT:
                frame = cv2.resize(frame, (PROCESS_WIDTH, PROCESS_HEIGHT), interpolation=cv2.INTER_LINEAR)

            with self.lock:
                self.frame = frame

    def read(self):
        with self.lock:
            if self.frame is None:
                return False, None
            return True, self.frame.copy()

    def stop(self):
        self.running = False
        if self.thread:
            self.thread.join(timeout=1.0)
        if self.cap:
            self.cap.release()


def clamp(v, mn, mx):
    return max(mn, min(mx, v))


def map_range(val, in_min, in_max, out_min, out_max):
    val = clamp(val, in_min, in_max)
    if in_max == in_min:
        return out_min
    return out_min + (float(val - in_min) / float(in_max - in_min)) * (out_max - out_min)


def low_pass_filter(current, target, alpha):
    return current + (target - current) * alpha


def smooth_value_or_init(current, target, alpha):
    if current is None:
        return float(target)
    return low_pass_filter(float(current), float(target), alpha)


def adaptive_step(error, min_step, max_step, max_error):
    ratio = clamp(error / max_error, 0.0, 1.0)
    return min_step + (max_step - min_step) * ratio


def smooth_move(curr, target, step):
    if abs(curr - target) <= step:
        return target
    return curr + step if curr < target else curr - step


def list_available_ports():
    return list(list_ports.comports())


def port_text(port_info):
    return f"{port_info.device} {port_info.description} {port_info.hwid}".lower()


def is_likely_esp32_port(port_info):
    text = port_text(port_info)
    return any(keyword in text for keyword in ESP32_KEYWORDS)


def score_port(port_info):
    text = port_text(port_info)
    score = 0
    for kw in ESP32_KEYWORDS:
        if kw in text:
            score += 10
    if "bluetooth" in text:
        score -= 50
    if "standard serial over bluetooth" in text:
        score -= 100
    return score


def try_open_serial(port_name):
    ser = serial.Serial(
        port=port_name,
        baudrate=BAUD_RATE,
        timeout=0.001,
        write_timeout=0.01,
        rtscts=False,
        dsrdtr=False,
        xonxoff=False,
    )
    try:
        ser.setDTR(False)
        ser.setRTS(False)
    except Exception:
        pass
    time.sleep(SERIAL_OPEN_DELAY)
    ser.reset_input_buffer()
    ser.reset_output_buffer()
    return ser


def connect_serial(preferred_port=None):
    errors = []
    ports = list_available_ports()

    if preferred_port:
        try:
            ser = try_open_serial(preferred_port)
            return ser, preferred_port, f"CONNECTED ({preferred_port})"
        except Exception as e:
            errors.append(f"{preferred_port}: {e}")

    likely_ports = sorted([p for p in ports if is_likely_esp32_port(p)], key=score_port, reverse=True)
    fallback_ports = sorted([p for p in ports if not is_likely_esp32_port(p)], key=score_port, reverse=True)
    candidates = likely_ports if likely_ports else fallback_ports[:1]

    for p in candidates:
        if preferred_port and p.device == preferred_port:
            continue
        try:
            ser = try_open_serial(p.device)
            return ser, p.device, f"CONNECTED ({p.device})"
        except Exception as e:
            errors.append(f"{p.device}: {e}")

    if not ports:
        return None, None, "DISCONNECTED - No COM ports detected"
    if not candidates:
        return None, None, "DISCONNECTED - COM ports found, but none looked like ESP32 USB serial"
    return None, None, "DISCONNECTED - " + " | ".join(errors)


def send_joint_if_changed(ser, joint, value):
    global last_serial_message
    value = int(round(value))
    old = last_sent.get(joint)
    if old is None or abs(value - old) >= ANGLE_SEND_DEADBAND:
        ser.write(f"{joint},{value}\n".encode())
        last_sent[joint] = value
        last_serial_message = f"TX: {joint}={value}"


def send_all_servos(ser, base, shoulder, elbow):
    b = int(round(clamp(base, BASE_MIN, BASE_MAX)))
    s = int(round(clamp(shoulder, SHOULDER_MIN, SHOULDER_MAX)))
    e = int(round(clamp(elbow, ELBOW_MIN, ELBOW_MAX)))
    send_joint_if_changed(ser, "BASE", b)
    send_joint_if_changed(ser, "SHOULDER", s)
    send_joint_if_changed(ser, "ELBOW", e)


def send_oled_face_state(ser, detected):
    global oled_face_state
    wanted = "DETECTED" if detected else "IDLE"
    if oled_face_state == wanted:
        return False
    ok = send_serial_command(ser, f"FACE:{wanted}")
    if ok:
        oled_face_state = wanted
    return ok


def send_oled_talk_state(ser, talking):
    global oled_talk_state
    wanted = "TALK:1" if talking else "TALK:0"
    if oled_talk_state == wanted:
        return False
    ok = send_serial_command(ser, wanted)
    if ok:
        oled_talk_state = wanted
    return ok


def send_oled_eye_position(ser, eye_x, eye_y):
    global oled_eye_x, oled_eye_y

    eye_x = int(round(clamp(eye_x, OLED_X_MIN, OLED_X_MAX)))
    eye_y = int(round(clamp(eye_y, OLED_Y_MIN, OLED_Y_MAX)))

    if (
        oled_eye_x is not None and
        oled_eye_y is not None and
        abs(eye_x - oled_eye_x) < OLED_EYE_DEADBAND and
        abs(eye_y - oled_eye_y) < OLED_EYE_DEADBAND
    ):
        return False

    ok = send_serial_command(ser, f"EYE:{eye_x},{eye_y}")
    if ok:
        oled_eye_x = eye_x
        oled_eye_y = eye_y
    return ok


def send_oled_tracking(ser, base, shoulder, face_detected):
    global last_oled_send_time

    if ser is None or not ser.is_open:
        return False

    now = time.time()
    if (now - last_oled_send_time) < OLED_SEND_INTERVAL:
        return False

    send_oled_face_state(ser, face_detected)

    if face_detected:
        eye_x = map_range(base, BASE_MIN, BASE_MAX, OLED_X_MIN, OLED_X_MAX)
        eye_y = map_range(shoulder, SHOULDER_MIN, SHOULDER_MAX, OLED_Y_MIN, OLED_Y_MAX)
    else:
        eye_x = 0
        eye_y = 0

    changed = send_oled_eye_position(ser, eye_x, eye_y)
    last_oled_send_time = now
    return changed


def send_serial_command(ser, text):
    global last_serial_message
    try:
        if ser and ser.is_open:
            cmd = text.strip() + "\n"
            ser.write(cmd.encode())
            last_serial_message = f"TX: {text.strip()}"
            return True
    except Exception as e:
        last_serial_message = f"SEND ERROR: {e}"
    return False


def post_barcode_to_server(barcode, ser=None):
    global last_posted_barcode, last_post_time, last_serial_message
    barcode = str(barcode or "").strip()
    if not barcode or not POST_SCAN_TO_SERVER:
        return False

    now = time.time()
    if barcode == last_posted_barcode and (now - last_post_time) < POST_SCAN_COOLDOWN:
        return False

    try:
        response = http_session.post(
            SERVER_SCAN_URL,
            json={"barcode": barcode},
            timeout=HTTP_TIMEOUT
        )
        last_posted_barcode = barcode
        last_post_time = now
        last_serial_message = f"SERVER {response.status_code}"
        return response.status_code == 200
    except Exception as e:
        last_serial_message = f"POST ERROR: {e}"
        return False


def read_serial_feedback(ser):
    global last_barcode, last_serial_message
    if not ser:
        return
    try:
        max_reads = 3
        reads = 0
        while ser.in_waiting and reads < max_reads:
            line = ser.readline().decode(errors="ignore").strip()
            reads += 1
            if not line:
                continue
            last_serial_message = line
            if READ_BARCODE_FROM_ESP32 and line.startswith("SCAN:"):
                last_barcode = line.replace("SCAN:", "", 1).strip()
                threading.Thread(target=post_barcode_to_server, args=(last_barcode, ser), daemon=True).start()
    except Exception as e:
        last_serial_message = f"READ ERROR: {e}"


def init_face_detector():
    if MEDIAPIPE_AVAILABLE:
        detector = mp.solutions.face_detection.FaceDetection(
            model_selection=0,
            min_detection_confidence=0.5,
        )
        return "mediapipe", detector

    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    return "haar", face_cascade


def detect_faces(frame, detector_type, detector):
    if detector_type == "mediapipe":
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        results = detector.process(rgb)
        faces = []
        if results.detections:
            h, w = frame.shape[:2]
            for det in results.detections:
                bbox = det.location_data.relative_bounding_box
                x = int(bbox.xmin * w)
                y = int(bbox.ymin * h)
                bw = int(bbox.width * w)
                bh = int(bbox.height * h)
                x = max(0, x)
                y = max(0, y)
                bw = max(1, min(bw, w - x))
                bh = max(1, min(bh, h - y))
                faces.append((x, y, bw, bh))
        return faces

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = detector.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(22, 22))
    return list(faces)


def choose_best_face(faces, previous_center):
    if len(faces) == 0:
        return None
    if previous_center is None:
        return max(faces, key=lambda f: f[2] * f[3])

    lock_x, lock_y = previous_center
    nearby_faces = []
    for f in faces:
        fx, fy, fw, fh = f
        cx_full = fx + fw / 2
        cy_full = fy + fh / 2
        if abs(cx_full - lock_x) <= SEARCH_MARGIN and abs(cy_full - lock_y) <= SEARCH_MARGIN:
            nearby_faces.append(f)
    candidate_faces = nearby_faces if nearby_faces else faces

    def face_score(face):
        fx, fy, fw, fh = face
        cx_full = fx + fw / 2
        cy_full = fy + fh / 2
        dist = (cx_full - lock_x) ** 2 + (cy_full - lock_y) ** 2
        area_bonus = fw * fh * 0.010
        return dist - area_bonus

    return min(candidate_faces, key=face_score)


def update_fps(now):
    global fps_time, fps_frames, current_fps
    if fps_time == 0.0:
        fps_time = now
    fps_frames += 1
    elapsed = now - fps_time
    if elapsed >= 0.5:
        current_fps = fps_frames / elapsed
        fps_frames = 0
        fps_time = now


def safe_roi(frame, x1, y1, x2, y2):
    h, w = frame.shape[:2]
    x1 = max(0, min(w, int(x1)))
    y1 = max(0, min(h, int(y1)))
    x2 = max(0, min(w, int(x2)))
    y2 = max(0, min(h, int(y2)))
    if x2 <= x1 or y2 <= y1:
        return None
    return frame[y1:y2, x1:x2]


def init_tts_engine():
    if not PYTTSX3_AVAILABLE:
        return None

    engine = pyttsx3.init()
    engine.setProperty("rate", FACE_AUDIO_RATE)

    try:
        voices = engine.getProperty("voices")
        preferred = None
        for voice in voices:
            name = (getattr(voice, "name", "") or "").lower()
            vid = (getattr(voice, "id", "") or "").lower()
            if any(keyword in name or keyword in vid for keyword in ["david", "mark", "guy", "male"]):
                preferred = voice.id
                break
        if preferred:
            engine.setProperty("voice", preferred)
    except Exception:
        pass

    return engine


def init_hidden_audio_player():
    if not PYGAME_AVAILABLE:
        return False
    try:
        if not pygame.mixer.get_init():
            pygame.mixer.pre_init(frequency=24000, size=-16, channels=1, buffer=4096)
            pygame.mixer.init()
        return True
    except Exception:
        return False


async def save_edge_tts_to_file(text, output_file):
    communicate = edge_tts.Communicate(
        text=text,
        voice=EDGE_TTS_VOICE,
        rate=EDGE_TTS_RATE,
        pitch=EDGE_TTS_PITCH,
        volume=EDGE_TTS_VOLUME,
    )
    await communicate.save(output_file)


def speak_with_edge_tts(text):
    if not EDGE_TTS_AVAILABLE:
        return False

    if not init_hidden_audio_player():
        return False

    temp_file = os.path.join(tempfile.gettempdir(), "jarvis_voice_output.mp3")

    try:
        if pygame.mixer.music.get_busy():
            pygame.mixer.music.stop()
            time.sleep(0.03)

        try:
            if hasattr(pygame.mixer.music, "unload"):
                pygame.mixer.music.unload()
        except Exception:
            pass

        asyncio.run(save_edge_tts_to_file(text, temp_file))

        pygame.mixer.music.load(temp_file)
        pygame.mixer.music.set_volume(1.0)
        pygame.mixer.music.play()

        while pygame.mixer.music.get_busy():
            time.sleep(0.02)

        try:
            pygame.mixer.music.stop()
        except Exception:
            pass

        try:
            if hasattr(pygame.mixer.music, "unload"):
                pygame.mixer.music.unload()
        except Exception:
            pass

        try:
            if os.path.exists(temp_file):
                os.remove(temp_file)
        except Exception:
            pass

        return True

    except Exception as e:
        print(f"EDGE TTS ERROR: {e}")
        return False


def speak_with_pyttsx3(text):
    engine = None
    try:
        engine = init_tts_engine()
        if engine is None:
            return False

        try:
            engine.stop()
        except Exception:
            pass

        engine.say(text)
        engine.runAndWait()
        return True

    except Exception:
        return False

    finally:
        try:
            if engine is not None:
                engine.stop()
        except Exception:
            pass


def speak_text(text):
    global audio_busy, last_serial_message, last_audio_play_time

    with audio_lock:
        if audio_busy:
            return False
        audio_busy = True

    talk_started = False

    try:
        spoken = False

        global ser
        active_ser = globals().get("ser")
        if active_ser and getattr(active_ser, "is_open", False):
            send_oled_face_state(active_ser, True)
            talk_started = send_oled_talk_state(active_ser, True)

        if PREFER_EDGE_TTS and EDGE_TTS_AVAILABLE:
            spoken = speak_with_edge_tts(text)
            if spoken:
                last_audio_play_time = time.time()
                last_serial_message = f'AUDIO: "{text}" [{EDGE_TTS_VOICE}]'

        if not spoken and PYTTSX3_AVAILABLE:
            spoken = speak_with_pyttsx3(text)
            if spoken:
                last_audio_play_time = time.time()
                last_serial_message = f'AUDIO: "{text}" [pyttsx3]'

        if not spoken:
            last_serial_message = "AUDIO ERROR: no working TTS engine found"

        return spoken

    except Exception as e:
        last_serial_message = f"AUDIO ERROR: {e}"
        return False

    finally:
        active_ser = globals().get("ser")
        if active_ser and getattr(active_ser, "is_open", False):
            send_oled_talk_state(active_ser, False)
            send_oled_face_state(active_ser, last_face_present)

        with audio_lock:
            audio_busy = False


def speak_welcome_message():
    if not ENABLE_FACE_WELCOME_AUDIO:
        return
    speak_text(FACE_AUDIO_MESSAGE)


def trigger_face_audio():
    global tts_thread, face_audio_armed, last_audio_play_time

    if not ENABLE_FACE_WELCOME_AUDIO:
        return

    now = time.time()

    if not face_audio_armed:
        return

    if (now - last_audio_play_time) < AUDIO_MIN_PLAY_GAP:
        return

    if tts_thread is not None and tts_thread.is_alive():
        return

    face_audio_armed = False
    tts_thread = threading.Thread(target=speak_welcome_message, daemon=True)
    tts_thread.start()


def should_welcome_face():
    if not face_audio_armed:
        return False
    if face_detected_since <= 0:
        return False
    return (time.time() - face_detected_since) >= FACE_STABLE_SPEAK_DELAY


def normalize_voice_text(text):
    text = (text or "").lower().strip()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    replacements = {
        "hiya": "hi",
        "heya": "hey",
        "hullo": "hello",
        "halo": "hello",
        "helo": "hello",
        "hay": "hey",
        "gud": "good",
        "morneng": "morning",
        "afternun": "afternoon",
        "evning": "evening",

        "trak": "track",
        "trax": "track",
        "trk": "track",
        "trek": "track",
        "traq": "track",

        "libary": "library",
        "librery": "library",
        "liberian": "librarian",

        "regster": "register",
        "regis": "register",
        "ter": "register",
        "registar": "register",
        "registrar": "register",
        "regitr": "register",
        "registor": "register",
        "rejister": "register",
        "rejester": "register",
        "ragister": "register",
        "ragester": "register",
        "rekister": "register",
        "rigister": "register",
        "rigester": "register",

        "sinup": "sign up",
        "sine": "sign",
        "signup": "sign up",
        "sign": "sign",
        "up": "up",
        "enrol": "enroll"
    }

    words = text.split()

    normalized_words = []
    i = 0
    while i < len(words):
        current = words[i]

        if i < len(words) - 1:
            pair = f"{words[i]} {words[i+1]}"
            if pair in ["regis ter", "sign up", "sine up", "sayn up", "sign me", "enroll me", "enrol me"]:
                if pair == "regis ter":
                    normalized_words.append("register")
                elif pair in ["sign up", "sine up", "sayn up"]:
                    normalized_words.append("sign up")
                else:
                    normalized_words.append(pair)
                i += 2
                continue

        normalized_words.append(replacements.get(current, current))
        i += 1

    return " ".join(normalized_words)


def fuzzy_ratio(a, b):
    return SequenceMatcher(None, a, b).ratio()


def is_trigger_match(heard_text):
    heard = normalize_voice_text(heard_text)
    if not heard:
        return False

    trigger_targets = set(normalize_voice_text(x) for x in (VOICE_TRIGGER_WORDS + VOICE_TRIGGER_ALIASES))
    heard_words = heard.split()

    for word in heard_words:
        if word in trigger_targets:
            return True

        for target in trigger_targets:
            if fuzzy_ratio(word, target) >= VOICE_TRIGGER_SIMILARITY:
                return True

    if heard in trigger_targets:
        return True

    for target in trigger_targets:
        if fuzzy_ratio(heard, target) >= VOICE_TRIGGER_SIMILARITY:
            return True

    return False


def is_register_match(heard_text):
    heard = normalize_voice_text(heard_text)
    if not heard:
        return False

    register_targets = [normalize_voice_text(x) for x in VOICE_REGISTER_WORDS]
    heard_words = heard.split()

    direct_phrases = [
        "register",
        "registration",
        "register me",
        "i want to register",
        "i want register",
        "can i register",
        "register now",
        "sign up",
        "sign me up",
        "signing up",
        "enroll",
        "enroll me"
    ]

    for phrase in direct_phrases:
        if phrase in heard:
            return True

    if heard in register_targets:
        return True

    for word in heard_words:
        for target in register_targets:
            if word == target:
                return True
            if fuzzy_ratio(word, target) >= 0.74:
                return True

    for target in register_targets:
        if target and target in heard:
            return True
        if fuzzy_ratio(heard, target) >= 0.72:
            return True

    return False


def speak_intro_and_arm_register():
    global waiting_for_register_command, wait_for_register_until
    global face_audio_armed, last_audio_play_time

    face_audio_armed = False
    last_audio_play_time = time.time()

    intro_spoken = speak_text(VOICE_REPLY_TEXT)
    if intro_spoken:
        waiting_for_register_command = True
        wait_for_register_until = time.time() + VOICE_WAIT_FOR_REGISTER_TIMEOUT


def voice_listener_loop():
    global voice_running, last_voice_command_time, last_serial_message
    global waiting_for_register_command, wait_for_register_until

    if not SPEECH_RECOGNITION_AVAILABLE:
        last_serial_message = "VOICE ERROR: speech_recognition not installed"
        return

    recognizer = sr.Recognizer()
    recognizer.energy_threshold = 180
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold = 0.35
    recognizer.non_speaking_duration = 0.15
    recognizer.phrase_threshold = 0.10
    recognizer.operation_timeout = 3

    try:
        mic = sr.Microphone(sample_rate=16000)
    except Exception as e:
        last_serial_message = f"VOICE ERROR: microphone not found ({e})"
        return

    try:
        with mic as source:
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
    except Exception as e:
        last_serial_message = f"VOICE ERROR: mic init failed ({e})"
        return

    while voice_running:
        try:
            with mic as source:
                audio = recognizer.listen(
                    source,
                    timeout=VOICE_LISTEN_TIMEOUT,
                    phrase_time_limit=VOICE_PHRASE_TIME_LIMIT
                )
        except sr.WaitTimeoutError:
            continue
        except Exception as e:
            last_serial_message = f"VOICE LISTEN ERROR: {e}"
            time.sleep(0.1)
            continue

        try:
            heard_text = recognizer.recognize_google(audio, language="en-US").lower().strip()
            normalized_heard = normalize_voice_text(heard_text)
            last_serial_message = f'VOICE HEARD: "{heard_text}" -> "{normalized_heard}"'
        except sr.UnknownValueError:
            continue
        except sr.RequestError as e:
            last_serial_message = f"VOICE API ERROR: {e}"
            time.sleep(0.2)
            continue
        except Exception as e:
            last_serial_message = f"VOICE ERROR: {e}"
            time.sleep(0.1)
            continue

        now = time.time()

        if waiting_for_register_command and now > wait_for_register_until:
            waiting_for_register_command = False

        if waiting_for_register_command and is_register_match(heard_text):
            waiting_for_register_command = False
            last_voice_command_time = now
            threading.Thread(
                target=speak_text,
                args=(VOICE_FOLLOWUP_REGISTER_TEXT,),
                daemon=True
            ).start()
            continue

        if (now - last_voice_command_time) < VOICE_COMMAND_COOLDOWN:
            continue

        if is_trigger_match(heard_text):
            last_voice_command_time = now

            if not audio_busy:
                threading.Thread(
                    target=speak_intro_and_arm_register,
                    daemon=True
                ).start()
            continue


def start_voice_listener():
    global voice_thread, voice_running

    if not ENABLE_VOICE_INTERACTION:
        return

    if not SPEECH_RECOGNITION_AVAILABLE:
        return

    if voice_thread is not None and voice_thread.is_alive():
        return

    voice_running = True
    voice_thread = threading.Thread(target=voice_listener_loop, daemon=True)
    voice_thread.start()


def stop_voice_listener():
    global voice_running, voice_thread
    voice_running = False
    if voice_thread is not None:
        voice_thread.join(timeout=1.0)


def detect_uniform_from_face(frame, face_box):
    if face_box is None or not ENABLE_UNIFORM_DETECTION:
        return False, 0.0, None

    x, y, w, h = face_box
    h_img, w_img = frame.shape[:2]

    torso_x1 = x - int(w * 0.60)
    torso_x2 = x + w + int(w * 0.60)
    torso_y1 = y + int(h * 0.90)
    torso_y2 = y + int(h * 4.60)
    torso = safe_roi(frame, torso_x1, torso_y1, torso_x2, torso_y2)
    if torso is None or torso.size == 0:
        return False, 0.0, None

    hsv = cv2.cvtColor(torso, cv2.COLOR_BGR2HSV)

    shirt_mask = cv2.inRange(hsv, (0, 0, 150), (180, 65, 255))
    green_mask = cv2.inRange(hsv, (35, 45, 35), (95, 255, 210))
    gold_mask = cv2.inRange(hsv, (12, 55, 70), (42, 255, 255))

    total_pixels = float(torso.shape[0] * torso.shape[1])
    if total_pixels <= 0:
        return False, 0.0, None

    shirt_ratio = cv2.countNonZero(shirt_mask) / total_pixels
    green_ratio = cv2.countNonZero(green_mask) / total_pixels
    gold_ratio = cv2.countNonZero(gold_mask) / total_pixels

    cx1 = int(torso.shape[1] * 0.32)
    cx2 = int(torso.shape[1] * 0.68)
    cy2 = int(torso.shape[0] * 0.78)
    center_strip = safe_roi(torso, cx1, 0, cx2, cy2)

    green_center_ratio = 0.0
    gold_center_ratio = 0.0
    if center_strip is not None and center_strip.size > 0:
        center_hsv = cv2.cvtColor(center_strip, cv2.COLOR_BGR2HSV)
        center_total = float(center_strip.shape[0] * center_strip.shape[1])
        green_center_ratio = cv2.countNonZero(cv2.inRange(center_hsv, (35, 45, 35), (95, 255, 210))) / center_total
        gold_center_ratio = cv2.countNonZero(cv2.inRange(center_hsv, (12, 55, 70), (42, 255, 255))) / center_total

    green_hit = max(green_ratio, green_center_ratio)
    gold_hit = max(gold_ratio, gold_center_ratio)
    score = (shirt_ratio * 0.72) + (green_hit * 7.0) + (gold_hit * 5.5)

    uniform_ok = (
        shirt_ratio >= UNIFORM_MIN_SHIRT_BRIGHT_RATIO and
        green_hit >= UNIFORM_MIN_GREEN_RATIO and
        gold_hit >= UNIFORM_MIN_GOLD_RATIO and
        score >= UNIFORM_MIN_COMBINED_SCORE
    )

    roi_box = (
        max(0, torso_x1),
        max(0, torso_y1),
        min(w_img, torso_x2),
        min(h_img, torso_y2),
    )
    return uniform_ok, score, roi_box


def main():
    global ser
    global base_angle, shoulder_angle, elbow_angle
    global target_base, target_shoulder, target_elbow
    global filtered_target_base, filtered_target_shoulder, filtered_target_elbow
    global locked_face_center, locked_face_width, frame_count
    global last_serial_attempt, last_send_time, last_oled_send_time, last_headless_status_time
    global last_face_seen_time, home_command_sent, last_detection_faces
    global last_uniform_detect_time, last_uniform_score, last_uniform_detected
    global last_face_present, last_face_lost_time, face_audio_armed, face_detected_since

    ser = None
    serial_status = "CONNECTING..."
    connected_port = None

    ser, connected_port, serial_status = connect_serial(SERIAL_PORT)
    last_serial_attempt = time.time()

    detector_type, detector = init_face_detector()
    camera = CameraStream(CAMERA_INDEX).start()

    start_voice_listener()

    if SHOW_CAMERA:
        cv2.namedWindow(WINDOW_NAME, cv2.WINDOW_NORMAL)
        cv2.resizeWindow(WINDOW_NAME, DISPLAY_WINDOW_WIDTH, DISPLAY_WINDOW_HEIGHT)

    last_face_seen_time = time.time()

    while True:
        now = time.time()
        update_fps(now)

        if ser is None and (now - last_serial_attempt) >= RECONNECT_INTERVAL:
            ser, connected_port, serial_status = connect_serial(SERIAL_PORT)
            last_serial_attempt = now
            if ser:
                home_command_sent = False

        ret, frame = camera.read()
        if not ret:
            continue

        h, w = frame.shape[:2]
        tracking_status = "SEARCHING"

        if frame_count % DETECT_EVERY_N_FRAMES == 0:
            last_detection_faces = detect_faces(frame, detector_type, detector)
        faces = last_detection_faces
        frame_count += 1

        best_face = choose_best_face(faces, locked_face_center)
        uniform_detected = False
        uniform_score = 0.0
        uniform_box = None

        if best_face is not None:
            last_face_seen_time = now
            home_command_sent = False

            xf, yf, wf, hf = best_face
            raw_cx = xf + wf / 2
            raw_cy = yf + hf / 2

            smooth_cx = smooth_value_or_init(locked_face_center[0] if locked_face_center else None, raw_cx, FACE_CENTER_ALPHA)
            smooth_cy = smooth_value_or_init(locked_face_center[1] if locked_face_center else None, raw_cy, FACE_CENTER_ALPHA)
            smooth_w = smooth_value_or_init(locked_face_width, wf, FACE_SIZE_ALPHA)

            cx, cy = int(round(smooth_cx)), int(round(smooth_cy))
            locked_face_center = (cx, cy)
            locked_face_width = float(smooth_w)
            tracking_status = "LOCKED"

            dx = cx - (w // 2)
            dy = cy - (h // 2)

            if REVERSE_BASE:
                absolute_base = map_range(cx, 0, w, BASE_MAX, BASE_MIN)
                correction_base = filtered_target_base - (dx * BASE_FINE_KP)
            else:
                absolute_base = map_range(cx, 0, w, BASE_MIN, BASE_MAX)
                correction_base = filtered_target_base + (dx * BASE_FINE_KP)

            if REVERSE_SHOULDER:
                absolute_shoulder = map_range(cy, 0, h, SHOULDER_MAX, SHOULDER_MIN)
                correction_shoulder = filtered_target_shoulder - (dy * SHOULDER_FINE_KP)
            else:
                absolute_shoulder = map_range(cy, 0, h, SHOULDER_MIN, SHOULDER_MAX)
                correction_shoulder = filtered_target_shoulder + (dy * SHOULDER_FINE_KP)

            desired_face_width = 58
            size_error = desired_face_width - locked_face_width
            absolute_elbow = map_range(locked_face_width, 20, 110, ELBOW_MAX, ELBOW_MIN)
            correction_elbow = filtered_target_elbow - (size_error * ELBOW_FINE_KP)

            target_base = clamp(
                filtered_target_base if abs(dx) <= CENTER_DEADZONE_X
                else (absolute_base * 0.72 + correction_base * 0.28),
                BASE_MIN, BASE_MAX
            )
            target_shoulder = clamp(
                filtered_target_shoulder if abs(dy) <= CENTER_DEADZONE_Y
                else (absolute_shoulder * 0.70 + correction_shoulder * 0.30),
                SHOULDER_MIN, SHOULDER_MAX
            )
            target_elbow = clamp(
                filtered_target_elbow if abs(size_error) <= SIZE_DEADZONE
                else (absolute_elbow * 0.82 + correction_elbow * 0.18),
                ELBOW_MIN, ELBOW_MAX
            )

            filtered_target_base = clamp(low_pass_filter(filtered_target_base, target_base, 0.16), BASE_MIN, BASE_MAX)
            filtered_target_shoulder = clamp(low_pass_filter(filtered_target_shoulder, target_shoulder, 0.14), SHOULDER_MIN, SHOULDER_MAX)
            filtered_target_elbow = clamp(low_pass_filter(filtered_target_elbow, target_elbow, 0.10), ELBOW_MIN, ELBOW_MAX)

            uniform_detected, uniform_score, uniform_box = detect_uniform_from_face(frame, best_face)
            last_uniform_score = uniform_score

            if not last_face_present:
                face_detected_since = now

            if should_welcome_face():
                trigger_face_audio()

            if ser:
                send_oled_tracking(ser, base_angle, shoulder_angle, True)

            last_face_present = True

            if uniform_detected:
                last_uniform_detect_time = now
                last_uniform_detected = True
            elif (now - last_uniform_detect_time) > UNIFORM_TEXT_COOLDOWN:
                last_uniform_detected = False

        else:
            no_face_elapsed = now - last_face_seen_time

            if last_face_present:
                last_face_lost_time = now
                last_face_present = False
                face_detected_since = 0.0
                locked_face_center = None
                locked_face_width = None

            if (not face_audio_armed) and last_face_lost_time and (now - last_face_lost_time) >= FACE_LOST_REARM_DELAY:
                face_audio_armed = True

            if no_face_elapsed < NO_FACE_TIMEOUT:
                tracking_status = "SEARCHING"
            else:
                tracking_status = "RESETTING TO START"

                filtered_target_base = low_pass_filter(filtered_target_base, HOME_BASE, NO_FACE_HOME_ALPHA_BASE)
                filtered_target_shoulder = low_pass_filter(filtered_target_shoulder, HOME_SHOULDER, NO_FACE_HOME_ALPHA_SHOULDER)
                filtered_target_elbow = low_pass_filter(filtered_target_elbow, HOME_ELBOW, NO_FACE_HOME_ALPHA_ELBOW)

                target_base = HOME_BASE
                target_shoulder = HOME_SHOULDER
                target_elbow = HOME_ELBOW

                if ser:
                    send_oled_tracking(ser, HOME_BASE, HOME_SHOULDER, False)

                if (now - last_uniform_detect_time) > UNIFORM_TEXT_COOLDOWN:
                    last_uniform_detected = False
                    last_uniform_score = 0.0

        base_step = adaptive_step(abs(filtered_target_base - base_angle), BASE_STEP_MIN, BASE_STEP_MAX, 24)
        shoulder_step = adaptive_step(abs(filtered_target_shoulder - shoulder_angle), SHOULDER_STEP_MIN, SHOULDER_STEP_MAX, 20)
        elbow_step = adaptive_step(abs(filtered_target_elbow - elbow_angle), ELBOW_STEP_MIN, ELBOW_STEP_MAX, 22)

        base_angle = smooth_move(base_angle, filtered_target_base, base_step)
        shoulder_angle = smooth_move(shoulder_angle, filtered_target_shoulder, shoulder_step)
        elbow_angle = smooth_move(elbow_angle, filtered_target_elbow, elbow_step)

        if ser:
            try:
                if ENABLE_SERVO and (now - last_send_time) >= SERVO_SEND_INTERVAL:
                    send_all_servos(ser, base_angle, shoulder_angle, elbow_angle)
                    last_send_time = now

                send_oled_tracking(ser, base_angle, shoulder_angle, best_face is not None)
                read_serial_feedback(ser)
                serial_status = f"CONNECTED ({connected_port})"
            except Exception as e:
                serial_status = f"DISCONNECTED - {e}"
                try:
                    ser.close()
                except Exception:
                    pass
                ser = None
                connected_port = None
                last_serial_attempt = now

        if SHOW_CAMERA:
            display = cv2.resize(frame, (DISPLAY_WINDOW_WIDTH, DISPLAY_WINDOW_HEIGHT), interpolation=cv2.INTER_NEAREST)

            if best_face is not None:
                sx = DISPLAY_WINDOW_WIDTH / float(w)
                sy = DISPLAY_WINDOW_HEIGHT / float(h)
                rx1 = int(xf * sx)
                ry1 = int(yf * sy)
                rx2 = int((xf + wf) * sx)
                ry2 = int((yf + hf) * sy)
                cv2.rectangle(display, (rx1, ry1), (rx2, ry2), (0, 255, 0), 2)

                if uniform_box is not None:
                    ux1, uy1, ux2, uy2 = uniform_box
                    ux1 = int(ux1 * sx)
                    uy1 = int(uy1 * sy)
                    ux2 = int(ux2 * sx)
                    uy2 = int(uy2 * sy)
                    cv2.rectangle(display, (ux1, uy1), (ux2, uy2), (255, 220, 0), 2)

            cv2.putText(display, f"{tracking_status} | FPS {current_fps:.1f}", (12, 28), cv2.FONT_HERSHEY_SIMPLEX, 0.65, (255, 255, 255), 2)
            uniform_text = f"UNIFORM: {'DETECTED' if last_uniform_detected else 'NOT MATCHED'} ({last_uniform_score:.2f})"
            cv2.putText(display, uniform_text, (12, 58), cv2.FONT_HERSHEY_SIMPLEX, 0.62, (0, 255, 255) if last_uniform_detected else (180, 180, 180), 2)

            if ENABLE_VOICE_INTERACTION and SPEECH_RECOGNITION_AVAILABLE:
                voice_state = "VOICE CMD: ON"
            elif ENABLE_VOICE_INTERACTION:
                voice_state = "VOICE CMD: MISSING speech_recognition"
            else:
                voice_state = "VOICE CMD: OFF"

            cv2.putText(display, voice_state, (12, 88), cv2.FONT_HERSHEY_SIMPLEX, 0.62, (255, 200, 0), 2)

            if PREFER_EDGE_TTS and EDGE_TTS_AVAILABLE:
                audio_state = f"VOICE OUT: {EDGE_TTS_VOICE}"
            elif PYTTSX3_AVAILABLE:
                audio_state = "VOICE OUT: pyttsx3"
            else:
                audio_state = "VOICE OUT: OFF"

            cv2.putText(display, audio_state, (12, 118), cv2.FONT_HERSHEY_SIMPLEX, 0.62, (0, 255, 0), 2)
            cv2.putText(display, f"CAMERA: INTEGRATED ({CAMERA_INDEX})", (12, 148), cv2.FONT_HERSHEY_SIMPLEX, 0.62, (255, 255, 0), 2)

            cv2.imshow(WINDOW_NAME, display)

            key = cv2.waitKey(1) & 0xFF
            if key == ord("q"):
                break

    stop_voice_listener()
    camera.stop()
    if SHOW_CAMERA:
        cv2.destroyAllWindows()
    if ser:
        ser.close()


if __name__ == "__main__":
    main()