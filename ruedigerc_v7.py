from __future__ import annotations
import time
import socket
import threading
import math
import logging
from dataclasses import dataclass
from typing import Optional, Callable
import serial
import serial.tools.list_ports
import mido
from pythonosc import osc_message_builder, osc_message

# --- LOGGING SETUP ---
logging.basicConfig(
    level=logging.INFO,  # Zurück auf INFO (weniger Spam)
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# --- KONSTANTEN ---
WING_IP = "192.168.1.4"
WING_PORT = 2223
LOCAL_PORT = 10024

# Fader-Einstellungen
FADER_MIN = 0
FADER_MAX = 1000
FADER_SILENT_THRESHOLD = 5
FADER_DEADBAND = 15  # Hysterese gegen Jitter (1.5%)

# Gain-Einstellungen
GAIN_MIN = -2.5    # dB
GAIN_MAX = 45.0   # dB
GAIN_STEP = 1.0   # dB pro Encoder-Tick

# Sync-Einstellungen
SYNC_INTERVAL = 0.5  # Sekunden zwischen Sync-Checks
MAX_SYNC_RETRIES = 3
CONNECTION_CHECK_INTERVAL = 60  # Sekunden


# --- DATENKLASSEN ---
@dataclass
class ChannelState:
    """Speichert den Zustand eines Kanals"""
    fader_value: int = 0
    last_stable_value: int = 0
    is_muted: bool = False
    is_touched: bool = False
    gain_db: float = 0.0  # Aktueller Gain-Wert


# --- WING STEUERUNG ---
class WingControl:
    """Kommunikation mit Behringer Wing über OSC"""
    
    def __init__(self, ip: str, port: int, local_port: int = 10024):
        self.ip = ip
        self.port = port
        self.local_port = local_port
        
        # Koeffizienten für Fader-zu-dB Konvertierung
        self.a, self.b, self.c = -6e-05, 0.1359, -62.895
        
        # Persistenter Socket für Queries (verhindert TIME_WAIT)
        self._query_socket = None
        self._socket_lock = threading.Lock()
        
        self._connection_ok = True
        self._failed_queries = 0
        
        # Merke dir welcher Query-Pfad funktioniert
        self._working_query_path = None
        
        self._init_query_socket()
        logger.info(f"Wing Control initialisiert für {ip}:{port}")
    
    def _init_query_socket(self):
        """Initialisiert den persistenten Query-Socket"""
        try:
            if self._query_socket:
                try:
                    self._query_socket.close()
                except:
                    pass
            
            self._query_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            self._query_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self._query_socket.bind(('', self.local_port))
            self._query_socket.settimeout(0.3)
            logger.debug("Query Socket initialisiert")
        except Exception as e:
            logger.error(f"Fehler beim Initialisieren von Query-Socket: {e}")
            raise
    
    def _query(self, address: str) -> Optional[float]:
        """
        Sendet OSC Query und wartet auf Antwort
        Verwendet persistenten Socket (kein TIME_WAIT Problem)
        """
        try:
            with self._socket_lock:  # Thread-sicher
                msg = osc_message_builder.OscMessageBuilder(address=address).build()
                
                logger.debug(f"Sende OSC Query: {address} → {self.ip}:{self.port}")
                logger.debug(f"OSC Bytes: {msg.dgram[:20]}...")  # Erste 20 Bytes
                
                self._query_socket.sendto(msg.dgram, (self.ip, self.port))
                
                # Versuche Antwort zu empfangen
                try:
                    data, addr = self._query_socket.recvfrom(4096)
                    logger.debug(f"Antwort empfangen von {addr}, {len(data)} bytes")
                    
                    result = osc_message.OscMessage(data).params[0]
                    logger.debug(f"OSC Result: {result}")
                    
                    self._failed_queries = 0
                    self._connection_ok = True
                    return result
                except socket.timeout:
                    logger.debug(f"Timeout bei Query: {address} (keine Antwort nach 0.3s)")
                    self._failed_queries += 1
                    return None
                
        except socket.timeout:
            logger.debug(f"Timeout bei Query: {address}")
            self._failed_queries += 1
            return None
        except Exception as e:
            logger.error(f"Fehler bei Query {address}: {e}")
            self._failed_queries += 1
            return None
    
    def _send_osc(self, address: str, value: float) -> bool:
        """
        Sendet OSC Message (fire-and-forget)
        Erstellt temporären Socket (kein bind() nötig für Sends)
        """
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
                sock.settimeout(1.0)
                builder = osc_message_builder.OscMessageBuilder(address=address)
                builder.add_arg(float(value))
                sock.sendto(builder.build().dgram, (self.ip, self.port))
                return True
        except Exception as e:
            logger.error(f"Fehler beim Senden an {address}: {e}")
            return False
    
    def check_connection(self) -> bool:
        """Prüft ob Wing erreichbar ist"""
        # Verwende den Pfad der beim letzten Mal funktioniert hat
        if self._working_query_path:
            result = self._query(self._working_query_path)
            if result is not None:
                self._connection_ok = True
                self._failed_queries = 0
                return True
        
        # Sonst: teste verschiedene Pfade
        logger.debug("Teste Wing Verbindung...")
        
        test_paths = ["/", "/info/name", "/info", "/status"]
        
        for path in test_paths:
            result = self._query(path)
            if result is not None:
                logger.info(f"✓ Wing antwortet auf {path}")
                self._working_query_path = path  # Merke dir diesen Pfad
                self._connection_ok = True
                self._failed_queries = 0
                return True
        
        # Keine Antwort
        logger.error("Keine OSC-Antwort vom Wing")
        logger.error("Mögliche Ursachen:")
        logger.error("  1. Wing OSC ist nicht aktiviert (Check: Setup → Network → OSC)")
        logger.error("  2. Firewall blockiert Port 2223")
        logger.error("  3. Wing lauscht auf anderem Port")
        logger.error(f"  4. Local Port {self.local_port} ist bereits belegt")
        
        self._connection_ok = False
        
        # Versuche Socket neu zu initialisieren
        try:
            logger.info("Versuche Socket-Neuinitialisierung...")
            self._init_query_socket()
        except Exception as e:
            logger.error(f"Socket-Neuinitialisierung fehlgeschlagen: {e}")
        
        return False
    
    def setFader(self, target_type: str, number: int, fader_1000: int) -> bool:
        """
        Setzt Fader-Position
        
        Args:
            target_type: "ch" für Channel, "mtx" für Matrix
            number: Kanalnummer (1-basiert)
            fader_1000: Faderwert 0-1000
            
        Returns:
            True bei Erfolg
        """
        # Validierung
        if not FADER_MIN <= fader_1000 <= FADER_MAX:
            logger.warning(f"Fader Wert {fader_1000} außerhalb {FADER_MIN}-{FADER_MAX}")
            fader_1000 = max(FADER_MIN, min(FADER_MAX, fader_1000))
        
        # Konvertierung zu dB
        if fader_1000 <= FADER_SILENT_THRESHOLD:
            db_val = -144.0
        else:
            db_val = (self.a * (fader_1000**2) + 
                     self.b * fader_1000 + 
                     self.c)
        
        address = f"/{target_type}/{number}/fdr"
        success = self._send_osc(address, db_val)
        
        if success:
            logger.debug(f"Fader {target_type}/{number} → {db_val:.1f}dB ({fader_1000}/1000)")
        
        return success
    
    def getFader(self, target_type: str, number: int) -> Optional[int]:
        """
        Liest Fader-Position
        
        Returns:
            Wert 0-1000 oder None bei Fehler
        """
        address = f"/{target_type}/{number}/fdr"
        raw_val = self._query(address)
        
        if raw_val is None:
            return None
        
        try:
            # -oo String behandeln
            if str(raw_val) == '-oo':
                return 0
            
            y = float(raw_val)
            
            # Sehr leise = 0
            if y <= -90.0:
                return 0
            
            # Quadratische Gleichung lösen
            discriminant = (self.b**2) - (4 * self.a * (self.c - y))
            if discriminant < 0:
                return 0
            
            result = (-self.b + math.sqrt(discriminant)) / (2 * self.a)
            return int(max(FADER_MIN, min(FADER_MAX, result)))
            
        except (ValueError, TypeError) as e:
            logger.error(f"Fehler beim Parsen von Fader-Wert: {e}")
            return None
    
    def getMute(self, target_type: str, number: int) -> bool:
        """
        Liest Mute-Status
        
        Returns:
            True wenn gemuted, False sonst
        """
        address = f"/{target_type}/{number}/mute"
        val = self._query(address)
        
        if val is None:
            return False
        
        try:
            status = int(float(val))
            return status > 0
        except (ValueError, TypeError):
            # String-Antworten behandeln
            val_str = str(val).upper()
            return val_str in ["ON", "MUTE", "TRUE", "1"]
    
    def setGain(self, source_group: str, source_num: int, gain_db: float) -> bool:
        """
        Setzt Gain einer INPUT-QUELLE (nicht Channel!)
        
        Args:
            source_group: "LCL", "AUX", "A", "B", "C", "SC", "USB", etc.
            source_num: Quellnummer (1-basiert)
            gain_db: Gain in dB (0 bis +60dB für Preamps)
            
        Returns:
            True bei Erfolg
        """
        # Validierung (Wing Input Gain: 0 bis +60dB)
        gain_db = max(GAIN_MIN, min(GAIN_MAX, gain_db))
        
        address = f"/io/in/{source_group}/{source_num}/g"
        success = self._send_osc(address, gain_db)
        
        if success:
            logger.debug(f"Input Gain {source_group}/{source_num} → {gain_db:.1f}dB")
        
        return success
    
    def getChannelInputSource(self, channel: int) -> Optional[tuple]:
        """
        Findet heraus, welche Input-Quelle einem Channel zugeordnet ist
        
        Args:
            channel: Channel-Nummer (1-40)
            
        Returns:
            Tuple (group, number) z.B. ("LCL", 1) oder None
        """
        # Lese Routing-Info
        grp_address = f"/ch/{channel}/in/conn/grp"
        in_address = f"/ch/{channel}/in/conn/in"
        
        grp = self._query(grp_address)
        in_num = self._query(in_address)
        
        if grp is None or in_num is None:
            return None
        
        try:
            return (str(grp), int(float(in_num)))
        except (ValueError, TypeError):
            return None
    
    def getGain(self, source_group: str, source_num: int) -> Optional[float]:
        """
        Liest Gain einer INPUT-QUELLE
        
        Args:
            source_group: "LCL", "AUX", "A", "B", "C", etc.
            source_num: Quellnummer (1-basiert)
            
        Returns:
            Gain in dB (0-60) oder None bei Fehler
        """
        address = f"/io/in/{source_group}/{source_num}/g"
        raw_val = self._query(address)
        
        if raw_val is None:
            return None
        
        try:
            gain = float(raw_val)
            logger.debug(f"Input Gain gelesen: {source_group}/{source_num} = {gain:.1f}dB")
            return max(GAIN_MIN, min(GAIN_MAX, gain))
        except (ValueError, TypeError) as e:
            logger.error(f"Fehler beim Parsen von Gain: {e}")
            return None
    
    def setMute(self, target_type: str, number: int, state: bool) -> bool:
        """
        Setzt Mute-Status
        
        Args:
            state: True = mute, False = unmute
        """
        address = f"/{target_type}/{number}/mute"
        success = self._send_osc(address, 1.0 if state else 0.0)
        
        if success:
            logger.debug(f"Mute {target_type}/{number} → {'ON' if state else 'OFF'}")
        
        return success
    
    def close(self):
        """Räumt Ressourcen auf"""
        try:
            if self._query_socket:
                self._query_socket.close()
                logger.info("Query Socket geschlossen")
        except Exception as e:
            logger.error(f"Fehler beim Schließen: {e}")


# --- X-TOUCH KLASSE ---
class XTouchExtender:
    """Kommunikation mit Behringer X-Touch Extender über MIDI"""
    
class XTouchExtender:
    def __init__(self):
        self.inport: Optional[mido.ports.BaseInput] = None
        self.outport: Optional[mido.ports.BaseOutput] = None
        self.selected_page = 0
        
        # Channel States (1-8)
        self.channels = {i: ChannelState() for i in range(1, 9)}
        
        # Callbacks
        self.on_fader_move: Callable[[int, int], None] = lambda ch, val: None
        self.on_button_press: Callable[[int, str, bool], None] = lambda ch, btn, state: None
        self.on_encoder_turn: Callable[[int, int], None] = lambda ch, delta: None
        
        # ====== NEU: Heartbeat tracking ======
        self._last_midi_received = time.time()
        self._connected = False
        # =====================================
        
        logger.info("X-Touch Extender initialisiert")

    def is_connected(self) -> bool:
        """Prüft ob X-Touch verbunden ist und Messages empfängt"""
        if time.time() - self._last_midi_received > 5.0:
            return False
        return self._connected and self.inport is not None and self.outport is not None
    

    def open(self, fragment: str = "xtouch") -> bool:
        """
        Öffnet MIDI-Verbindung zum X-Touch
        """
        try:
            # ====== NEU: Schließe alte Verbindungen ======
            self.close()
            # =============================================
            
            names = mido.get_output_names()
            logger.info(f"Verfügbare MIDI Ports: {names}")
            
            target = next((p for p in names if "X-TOUCH-" in p), None)
            
            if not target:
                target = next((p for p in names if fragment.lower() in p.lower()), None)
            
            if not target:
                logger.error(f"Kein Port mit '{fragment}' gefunden")
                logger.info("Verfügbare Ports:")
                for name in names:
                    logger.info(f"  - {name}")
                # ====== NEU ======
                self._connected = False
                # =================
                return False
            
            logger.info(f"Verwende MIDI Port: {target}")
            
            self.outport = mido.open_output(target)
            self.inport = mido.open_input(target, callback=self._input_callback)
            
            # ====== NEU: Status tracking ======
            self._connected = True
            self._last_midi_received = time.time()
            # ==================================
            
            # Initialisierung
            self._init_display()
            
            return True
            
        except Exception as e:
            logger.error(f"Fehler beim Öffnen von MIDI: {e}")
            # ====== NEU ======
            self._connected = False
            # =================
            return False

    def reconnect(self, fragment: str = "xtouch") -> bool:
        """Versucht Verbindung wiederherzustellen"""
        logger.info("Versuche X-Touch Reconnect...")
        success = self.open(fragment)
        if success:
            logger.info("✓ X-Touch Reconnect erfolgreich!")
            # Stelle aktuellen Zustand wieder her
            self.update_page_display(self.selected_page)
            self.set_led(self.selected_page + 1, "select", True)
        return success

    def _init_display(self):
        """Initialisiert Display und LEDs"""
        logger.info("Initialisiere X-Touch Display und LEDs...")
        
        # Alle LEDs aus
        for ch in range(1, 9):
            for btn in ["rec", "solo", "mute", "select"]:
                self.set_led(ch, btn, False)
        
        # Page 0 als Standard
        self.set_led(1, "select", True)
        self.update_page_display(0)
        
        logger.info("X-Touch initialisiert")
    
    def _input_callback(self, msg):
        """Verarbeitet eingehende MIDI Messages"""
        try:
            self._last_midi_received = time.time()
            if msg.type == 'pitchwheel':
                ch = msg.channel + 1
                val = int((msg.pitch + 8192) / 16383.0 * FADER_MAX)
                
                # Nur bei Touch oder spezielle Kanäle
                if self.channels[ch].is_touched or ch in [5, 7]:
                    self.on_fader_move(ch, val)
            
            elif msg.type == 'control_change':
                # Encoder (Drehregler) - CC 16-23 für Kanäle 1-8
                if 16 <= msg.control <= 23:
                    ch = msg.control - 15  # CC 16 = Kanal 1
                    
                    # Relative Encoder im 2's complement Format
                    # Werte 1-63: Rechtsdrehung (positiv)
                    # Werte 65-127: Linksdrehung (negativ in 2's complement)
                    if msg.value > 64:
                        # Linksdrehung: konvertiere von 2's complement
                        delta = msg.value - 128
                    elif msg.value < 64:
                        # Rechtsdrehung
                        delta = msg.value
                    else:
                        delta = 0  # 64 = kein Delta
                    
                    logger.debug(f"Encoder CH{ch}: raw={msg.value}, delta={delta}")
                    self.on_encoder_turn(ch, delta)
            
            elif msg.type in ['note_on', 'note_off']:
                note = msg.note
                state = (msg.type == 'note_on' and msg.velocity > 0)
                
                # Fader Touch (Note 104-111)
                if 104 <= note <= 111:
                    ch = note - 103
                    self.channels[ch].is_touched = state
                    logger.debug(f"Fader {ch} touch: {state}")
                
                # Buttons
                else:
                    btn_type = None
                    ch = None
                    
                    if 0 <= note <= 7:
                        btn_type, ch = "rec", note + 1
                    elif 8 <= note <= 15:
                        btn_type, ch = "solo", note - 7
                    elif 16 <= note <= 23:
                        btn_type, ch = "mute", note - 15
                    elif 24 <= note <= 31:
                        btn_type, ch = "select", note - 23
                    
                    if btn_type and ch:
                        self.on_button_press(ch, btn_type, state)
        
        except Exception as e:
            logger.error(f"Fehler in MIDI Callback: {e}")
    
    def set_text(self, ch: int, text: str):
        """Setzt Display-Text für Kanal"""
        if not 1 <= ch <= 8:
            logger.warning(f"Ungültiger Channel für Text: {ch}")
            return
        
        try:
            data = [0x00, 0x00, 0x66, 0x15, 0x12, (ch-1)*7]
            data += [ord(c) for c in text.ljust(7)[:7]]
            self.outport.send(mido.Message('sysex', data=data))
        except Exception as e:
            logger.error(f"Fehler beim Setzen von Text: {e}")
    
    def set_fader(self, ch: int, val: int):
        """Setzt Fader Position"""
        if not 1 <= ch <= 8:
            return
        
        try:
            val = max(FADER_MIN, min(FADER_MAX, val))
            pitch = int((val / FADER_MAX) * 16383) - 8192
            self.outport.send(mido.Message('pitchwheel', channel=ch-1, pitch=pitch))
        except Exception as e:
            logger.error(f"Fehler beim Setzen von Fader {ch}: {e}")
    
    def set_encoder_ring(self, ch: int, value: int, mode: int = 1):
        """
        Setzt LED-Ring um den Encoder
        
        Args:
            ch: Kanal 1-8
            value: Position 0-11 (12 LEDs im Ring) oder 0-13 für verschiedene Modi
            mode: 0=Single, 1=Pan (Mitte), 2=Fan, 3=Spread
        """
        if not 1 <= ch <= 8:
            return
        
        try:
            # CC 48-55 für Encoder-Ringe
            cc_num = 47 + ch
            
            # Value encoding: bits 0-3 = position, bits 4-5 = mode
            encoded_value = (mode << 4) | (value & 0x0F)
            
            self.outport.send(mido.Message('control_change', 
                                          control=cc_num, 
                                          value=encoded_value))
        except Exception as e:
            logger.error(f"Fehler beim Setzen von Encoder-Ring: {e}")
    
    def set_led(self, ch: int, btn: str, state: bool):
        """Setzt Button LED"""
        if not 1 <= ch <= 8:
            return
        
        notes = {"rec": 0, "solo": 8, "mute": 16, "select": 24}
        
        if btn not in notes:
            return
        
        try:
            note = notes[btn] + ch - 1
            velocity = 127 if state else 0
            self.outport.send(mido.Message('note_on', note=note, velocity=velocity))
        except Exception as e:
            logger.error(f"Fehler beim Setzen von LED: {e}")
    
    def set_color(self, ch: int, color_name: str):
        """
        Setzt Display-Farbe für einen Kanal
        
        Args:
            ch: Kanal 1-8
            color_name: "black", "red", "green", "yellow", "blue", "magenta", "cyan", "white"
        """
        if not 1 <= ch <= 8:
            logger.warning(f"Ungültiger Channel für Farbe: {ch}")
            return
        
        colors = {
            "black": 0,
            "red": 1,
            "green": 2,
            "yellow": 3,
            "blue": 4,
            "magenta": 5,
            "cyan": 6,
            "white": 7
        }
        
        color_idx = colors.get(color_name.lower(), 7)  # Default: white
        
        try:
            # X-Touch Extender benötigt möglicherweise ein anderes Format
            # Format 1: Standard (dein Original)
            data = [0x00, 0x00, 0x66, 0x15, 0x72, ch-1, color_idx]
            self.outport.send(mido.Message('sysex', data=data))
            time.sleep(0.01)  # Kurze Pause
            logger.debug(f"Farbe für Kanal {ch} → {color_name} (ID {color_idx})")
        except Exception as e:
            logger.error(f"Fehler beim Setzen von Farbe: {e}")
    
    def test_color_formats(self, ch: int = 1):
        """
        Testet verschiedene SysEx-Formate für Farben
        Rufe diese Funktion auf um herauszufinden welches Format funktioniert
        """
        logger.info(f"=== Teste Farb-Formate für Kanal {ch} ===")
        
        # Test mit ROT (color_idx = 1)
        color_idx = 1
        
        formats = {
            "Format 1 (Standard)": [0x00, 0x00, 0x66, 0x15, 0x72, ch-1, color_idx],
            "Format 2 (Alternative Manufacturer ID)": [0x00, 0x00, 0x66, 0x14, 0x72, ch-1, color_idx],
            "Format 3 (Kurz)": [0x00, 0x00, 0x66, 0x72, ch-1, color_idx],
            "Format 4 (Mit Command 0x58)": [0x00, 0x00, 0x66, 0x58, ch-1, color_idx],
            "Format 5 (Offset anders)": [0x00, 0x00, 0x66, 0x15, 0x72, (ch-1)*7, color_idx],
        }
        
        for name, data in formats.items():
            logger.info(f"Teste {name}: {data}")
            try:
                self.outport.send(mido.Message('sysex', data=data))
                logger.info(f"  → Gesendet. Wurde Display rot?")
                time.sleep(2)  # 2 Sekunden warten um Effekt zu sehen
            except Exception as e:
                logger.error(f"  → Fehler: {e}")
        
        logger.info("=== Test abgeschlossen ===")
        logger.info("Falls ein Format funktioniert hat, notiere welches!")
    
    def update_page_display(self, page: int):
        """Aktualisiert Display-Texte und Farben für Seite"""
        texts = {
            0: ["Mic 1", "Mic 2", "Beamer", "BT", "Saal", "Foyer", "PA-Ext", "Licht"],
            1: ["In 1", "In 2", "In 3", "In 4", "In 5", "In 6", "In 7", "In 8"],
            2: ["In 9", "In 10", "In 11", "In 12", "In 13", "In 14", "In 15", "In 16"],
            3: ["Licht L1", "Licht L2", "Licht R1", "Licht R2", "DMX 5", "DMX 6", "DMX 7", "DMX 8"]
        }
        
        # Farben pro Seite definieren
        colors = {
            0: ["cyan", "cyan", "green", "blue", "white", "yellow", "yellow", "magenta"],  # Inputs cyan, Outputs yellow, Licht magenta
            1: ["red", "red", "red", "red", "red", "red", "red", "red"],  # Alle rot für Inputs 9-16
            2: ["green", "green", "green", "green", "green", "green", "green", "green"],  # Alle grün für Inputs 17-24
            3: ["magenta", "magenta", "magenta", "magenta", "magenta", "magenta", "magenta", "magenta"]  # DMX in magenta
        }
        
        page_texts = texts.get(page, [""] * 8)
        page_colors = colors.get(page, ["white"] * 8)
        
        for i, (text, color) in enumerate(zip(page_texts, page_colors), start=1):
            self.set_text(i, text)
            time.sleep(0.01)  # Kurze Pause zwischen Commands
            self.set_color(i, color)
    
    def update_page_mutes(self, page: int, wing: WingControl):
        """Aktualisiert Mute-LEDs für aktuelle Seite"""
        try:
            if page == 0:
                # Inputs 1-4
                for ch in range(1, 5):
                    muted = wing.getMute("ch", ch)
                    self.set_led(ch, 'mute', muted)
                
                # Kanal 5 leer
                self.set_led(5, 'mute', False)
                
                # Outputs (Matrix 1-2)
                for ch in [5, 6, 7]:
                    muted = wing.getMute("mtx", ch - 4)
                    self.set_led(ch, 'mute', muted)
                
                # Kanal 8 (Licht)
                self.set_led(8, 'mute', False)
            
            elif page == 1:
                # Channels 9-16
                for ch in range(1, 9):
                    muted = wing.getMute("ch", ch + 8)
                    self.set_led(ch, 'mute', muted)
            
            elif page == 2:
                # Channels 17-24
                for ch in range(1, 9):
                    muted = wing.getMute("ch", ch + 16)
                    self.set_led(ch, 'mute', muted)
            
            else:
                # DMX oder andere
                for ch in range(1, 9):
                    self.set_led(ch, 'mute', False)
        
        except Exception as e:
            logger.error(f"Fehler beim Aktualisieren von Mutes: {e}")
    
    def update_page_faders(self, page: int, wing: WingControl, dmx: DMXController):
        """Aktualisiert Fader für aktuelle Seite"""
        logger.info(f"Aktualisiere Fader für Seite {page}")
        try:
            if page == 0:
                # Inputs 1-4
                for ch in range(1, 5):
                    val = wing.getFader("ch", ch)
                    if val is not None:
                        logger.info(f"Setze Fader {ch} auf {val}")
                        self.set_fader(ch, val)
                        self.channels[ch].fader_value = val
                        self.channels[ch].last_stable_value = val
                    time.sleep(0.02)
                
                # Kanal 5 leer
                self.set_fader(5, 0)
                
                # Outputs (Matrix 1-2)
                for ch, mtx_num in [(5, 1), (6, 2), (7, 3)]:
                    val = wing.getFader("mtx", mtx_num)
                    if val is not None:
                        logger.info(f"Setze Matrix-Fader {ch} (MTX {mtx_num}) auf {val}")
                        self.set_fader(ch, val)
                        self.channels[ch].fader_value = val
                        self.channels[ch].last_stable_value = val
                    time.sleep(0.02)
                
                # Kanal 8 (Licht)
                self.set_fader(8, dmx.getSaallicht())

            elif page == 1:
                # Channels 9-16
                for ch in range(1, 9):
                    val = wing.getFader("ch", ch + 8)
                    if val is not None:
                        self.set_fader(ch, val)
                        self.channels[ch].fader_value = val
                        self.channels[ch].last_stable_value = val
                    time.sleep(0.02)  # Kleine Pause zwischen Fadern
            
            elif page == 2:
                # Channels 17-24
                for ch in range(1, 9):
                    val = wing.getFader("ch", ch + 16)
                    if val is not None:
                        self.set_fader(ch, val)
                        self.channels[ch].fader_value = val
                        self.channels[ch].last_stable_value = val
                    time.sleep(0.02)
            
            elif page == 3:
                for ch in range(1, 9):
                    print(dmx.getDMX(ch))
                    self.set_fader(ch, dmx.getDMX(ch))
                    time.sleep(0.02)
        
        except Exception as e:
            logger.error(f"Fehler beim Aktualisieren von Fadern: {e}")
    
    def close(self):
        """Schließt MIDI-Verbindung"""
        try:
            if self.inport:
                self.inport.close()
                # ====== NEU ======
                self.inport = None
                # =================
            if self.outport:
                self.outport.close()
                # ====== NEU ======
                self.outport = None
                # =================
            # ====== NEU ======
            self._connected = False
            # =================
            logger.info("X-Touch Verbindung geschlossen")
        except Exception as e:
            logger.error(f"Fehler beim Schließen: {e}")
# --- DMX KLASSE ---
class DMXController:

    def __init__(self, baudrate=9600):
        self.baudrate = baudrate
        self.ser = None
        self.dmx_stored_values = [0] * 513

    def connect(self):
        """
            Verbindet sich per USB mit einem DMX interface

        """
        ports = list(serial.tools.list_ports.comports())
        target_port = None
        for p in ports:
            # Speziell für M4 Mac / CH340 / FTDI
            if any(x in p.description.upper() for x in ["USB", "CH340", "SERIAL", "FT232"]):
                target_port = p.device
                break
        
        if target_port:
            try:
                self.ser = serial.Serial(target_port, self.baudrate, timeout=1)
                logger.info("DMX ist verbunden mit {target_port}")
                time.sleep(2) # Wichtig für Arduino-Reboot
                return True
            except Exception as e:
                logger.info("DMX Fehler")
        return False

    def sendDMX(self, channel: int, value_1000: int):
        """
        Übergibt Werte von 0-1000 und rechnet sie auf 0-255 um.

        args:
        ch (int): DMX Kanal
        vlaue_1000 (int): Wert zwischen 0 und 1000
        """
        if self.ser and self.ser.is_open:
            # Umrechnung und Begrenzung (Clamping)
            val_255 = int(max(0, min(1000, value_1000)) * 255 / 1000)
            self.dmx_stored_values[channel] = value_1000
            msg = f"{channel},{val_255}\n"
            self.ser.write(msg.encode('utf-8'))
        else:
            logger.info("DMX nicht verbunden!")

    def getDMX(self, channel: int):
        """
        Meldet zurück, auf welchen Wert ein DMX Kanal gesetzt wurde
        
        Args:
            channel (int): Abzufragender Kanal
            
        Returns:
            int: Wert vom Kanal (0-100)
        """
        if 0 <= self.dmx_stored_values[channel]<= 1000:
            return(self.dmx_stored_values[channel])
        else:
            logger.warning("Abgefragter DMX-Wert außerhalb des zulässigen Rahmens")
            return(0)

    def getSaallicht(self):
        dmx_mittel = 0
        for i in range(1, 5, 1):
            dmx_mittel = dmx_mittel + self.dmx_stored_values[i]
        dmx_mittel = int(dmx_mittel/4)
        return dmx_mittel

    def close(self):
        if self.ser: self.ser.close()


# --- SYNC LOOP ---
def sync_loop(device: XTouchExtender, wing: WingControl, dmx: DMXController):
    """
    Synchronisiert Wing-Status mit X-Touch
    Läuft in eigenem Thread
    """
    logger.info("Sync Loop gestartet")
    consecutive_errors = 0
    
    while True:
        try:
            if not device.is_connected():
                logger.debug("X-Touch nicht verbunden, überspringe Sync")
                time.sleep(SYNC_INTERVAL)
                continue
            page = device.selected_page
            
            # Gain-Synchronisation für Input-Kanäle (alle Seiten)
            sync_gain_for_page(device, wing, page)
            
            # Page 0: Inputs 1-4 + Matrix 1-2
            if page == 0:
                # Inputs
                for ch in range(1, 5):
                    if not device.channels[ch].is_touched:
                        wing_fader = wing.getFader("ch", ch)
                        if wing_fader is not None:
                            current = device.channels[ch].fader_value
                            if abs(wing_fader - current) > FADER_DEADBAND:
                                device.set_fader(ch, wing_fader)
                            device.channels[ch].fader_value = wing_fader
                            device.channels[ch].last_stable_value = wing_fader
                    
                    # Mute Status
                    muted = wing.getMute("ch", ch)
                    device.set_led(ch, 'mute', muted)
                
                # Matrix Outputs
                for ch, mtx_num in [(5, 1), (6, 2), (7, 3)]:
                    if not device.channels[ch].is_touched:
                        wing_fader = wing.getFader("mtx", mtx_num)
                        if wing_fader is not None:
                            current = device.channels[ch].fader_value
                            if abs(wing_fader - current) > FADER_DEADBAND:
                                device.set_fader(ch, wing_fader)
                            device.channels[ch].fader_value = wing_fader
                            device.channels[ch].last_stable_value = wing_fader
                    
                    # Mute Status
                    muted = wing.getMute("mtx", mtx_num)
                    device.set_led(ch, 'mute', muted)

                # Licht 
                device.set_fader(8, dmx.getSaallicht())
            
            # Page 1: Channels 9-16
            elif page == 1:
                for ch in range(1, 9):
                    wing_ch = ch + 8
                    
                    if not device.channels[ch].is_touched:
                        wing_fader = wing.getFader("ch", wing_ch)
                        if wing_fader is not None:
                            current = device.channels[ch].fader_value
                            if abs(wing_fader - current) > FADER_DEADBAND:
                                device.set_fader(ch, wing_fader)
                            device.channels[ch].fader_value = wing_fader
                            device.channels[ch].last_stable_value = wing_fader
                    
                    # Mute Status
                    muted = wing.getMute("ch", wing_ch)
                    device.set_led(ch, 'mute', muted)
            
            # Page 2: Channels 17-24
            elif page == 2:
                for ch in range(1, 9):
                    wing_ch = ch + 16
                    
                    if not device.channels[ch].is_touched:
                        wing_fader = wing.getFader("ch", wing_ch)
                        if wing_fader is not None:
                            current = device.channels[ch].fader_value
                            if abs(wing_fader - current) > FADER_DEADBAND:
                                device.set_fader(ch, wing_fader)
                            device.channels[ch].fader_value = wing_fader
                            device.channels[ch].last_stable_value = wing_fader
                    
                    # Mute Status
                    muted = wing.getMute("ch", wing_ch)
                    device.set_led(ch, 'mute', muted)
            
            # Erfolg - Error Counter zurücksetzen
            consecutive_errors = 0
            
        except Exception as e:
            consecutive_errors += 1
            logger.error(f"Fehler in Sync Loop (#{consecutive_errors}): {e}")
            
            # Bei vielen Fehlern: Verbindung prüfen
            if consecutive_errors >= MAX_SYNC_RETRIES:
                logger.warning("Zu viele Sync-Fehler, prüfe Verbindung...")
                if not wing.check_connection():
                    logger.error("Wing nicht erreichbar!")
                consecutive_errors = 0  # Reset
        
        time.sleep(SYNC_INTERVAL)


def sync_gain_for_page(device: XTouchExtender, wing: WingControl, page: int):
    """Synchronisiert Gain-Werte und LED-Ringe für die aktuelle Seite"""
    try:
        if page == 0:
            # Nur Kanäle 1-4 sind Inputs
            for ch in range(1, 5):
                wing_ch = ch
                source = wing.getChannelInputSource(wing_ch)
                if source:
                    grp, num = source
                    gain = wing.getGain(grp, num)
                    if gain is not None and abs(gain - device.channels[ch].gain_db) > 0.5:
                        device.channels[ch].gain_db = gain
                        ring_value = int((gain / GAIN_MAX) * 11)
                        device.set_encoder_ring(ch, ring_value, mode=1)
        
        elif page == 1:
            # Kanäle 9-16
            for ch in range(1, 9):
                wing_ch = ch + 8
                source = wing.getChannelInputSource(wing_ch)
                if source:
                    grp, num = source
                    gain = wing.getGain(grp, num)
                    if gain is not None and abs(gain - device.channels[ch].gain_db) > 0.5:
                        device.channels[ch].gain_db = gain
                        ring_value = int((gain / GAIN_MAX) * 11)
                        device.set_encoder_ring(ch, ring_value, mode=1)
        
        elif page == 2:
            # Kanäle 17-24
            for ch in range(1, 9):
                wing_ch = ch + 16
                source = wing.getChannelInputSource(wing_ch)
                if source:
                    grp, num = source
                    gain = wing.getGain(grp, num)
                    if gain is not None and abs(gain - device.channels[ch].gain_db) > 0.5:
                        device.channels[ch].gain_db = gain
                        ring_value = int((gain / GAIN_MAX) * 11)
                        device.set_encoder_ring(ch, ring_value, mode=1)
    
    except Exception as e:
        logger.debug(f"Fehler beim Synchronisieren von Gain: {e}")


# --- HAUPTPROGRAMM ---
def main():
    """Hauptprogramm"""
    logger.info("=== Wing X-Touch Controller Start ===")
    
    # Initialisierung
    wing = WingControl(ip=WING_IP, port=WING_PORT, local_port=LOCAL_PORT)
    device = XTouchExtender()
    dmx = DMXController()
    dmx.connect()
    
    # Verbindung prüfen
    logger.info("Prüfe Wing Verbindung...")
    if not wing.check_connection():
        logger.error("Wing nicht erreichbar! Prüfe IP und Netzwerk.")
        logger.error(f"Versuche: {WING_IP}:{WING_PORT}")
        wing.close()
        return
    
    logger.info(f"Wing nutzt Query-Pfad: {wing._working_query_path}")
    
    # X-Touch öffnen
    logger.info("Öffne X-Touch...")
    if not device.open("xtouch"):
        logger.error("X-Touch nicht gefunden! Prüfe MIDI-Verbindung.")
        logger.info("Tipp: Wenn der Port nicht gefunden wird, kannst du in open() den exakten Namen angeben")
        wing.close()
        return
    
    logger.info("✓ Alle Geräte verbunden!")
    
    # Initialisiere Gain-Werte für Seite 0
    for ch in range(1, 5):
        # Finde Input-Quelle für diesen Channel
        source = wing.getChannelInputSource(ch)
        if source:
            grp, num = source
            gain = wing.getGain(grp, num)
            if gain is not None:
                device.channels[ch].gain_db = gain
                # Zeige Gain im LED-Ring (0-60dB → 0-11 LEDs)
                ring_value = int((gain / GAIN_MAX) * 11)
                device.set_encoder_ring(ch, ring_value, mode=1)

    # Button Handler
    def on_button(ch: int, btn: str, state: bool):
        """Verarbeitet Button-Presses"""
        if not state:  # Nur auf Press reagieren
            return
        
        page = device.selected_page
        
        # MUTE Button
        if btn == "mute":
            if page == 0:
                # Inputs 1-4
                if ch <= 4:
                    current_mute = wing.getMute("ch", ch)
                    new_state = not current_mute
                    wing.setMute("ch", ch, new_state)
                    device.set_led(ch, "mute", new_state)
                
                # Matrix 1-2
                elif ch in [5, 6, 7]:
                    mtx_num = ch - 4
                    current_mute = wing.getMute("mtx", mtx_num)
                    new_state = not current_mute
                    wing.setMute("mtx", mtx_num, new_state)
                    device.set_led(ch, "mute", new_state)
            
            elif page == 1:
                # Channels 9-16
                wing_ch = ch + 8
                current_mute = wing.getMute("ch", wing_ch)
                new_state = not current_mute
                wing.setMute("ch", wing_ch, new_state)
                device.set_led(ch, "mute", new_state)
            
            elif page == 2:
                # Channels 17-24
                wing_ch = ch + 16
                current_mute = wing.getMute("ch", wing_ch)
                new_state = not current_mute
                wing.setMute("ch", wing_ch, new_state)
                device.set_led(ch, "mute", new_state)
        
        # SELECT Button (Seitenwechsel)
        elif btn == "select":
            logger.info(f"Wechsel zu Seite {ch-1}")
            
            # Alle Select-LEDs aus & Encoder LEDS aus
            for i in range(1, 9):
                device.set_led(i, "select", False)
                device.set_encoder_ring(i, 0, 0)
                # State zurücksetzen
                device.channels[i].fader_value = 0
                device.channels[i].last_stable_value = 0
            
            
            # Neue Seite
            device.selected_page = ch - 1
            device.set_led(ch, "select", True)
            
            # Display, Fader und Mutes aktualisieren
            device.update_page_display(device.selected_page)
            time.sleep(0.05)
            device.update_page_faders(device.selected_page, wing, dmx)
            time.sleep(0.05)
            device.update_page_mutes(device.selected_page, wing)
            time.sleep(0.05)
            
            # Gain-Werte und LED-Ringe aktualisieren
            if device.selected_page in [0, 1, 2]:  # Nur bei Input-Kanälen
                for i in range(1, 9):
                    # Berechne Wing-Kanal
                    if device.selected_page == 0 and i <= 4:
                        wing_ch = i
                    elif device.selected_page == 1:
                        wing_ch = i + 8
                    elif device.selected_page == 2:
                        wing_ch = i + 16
                    else:
                        continue
                    
                    # Finde Input-Quelle und lade Gain
                    source = wing.getChannelInputSource(wing_ch)
                    if source:
                        grp, num = source
                        gain = wing.getGain(grp, num)
                        if gain is not None:
                            device.channels[i].gain_db = gain
                            ring_value = int((gain / GAIN_MAX) * 11)
                            device.set_encoder_ring(i, ring_value, mode=1)
    
    # Encoder Handler (für Gain-Steuerung)
    def on_encoder(ch: int, delta: int):
        """Verarbeitet Encoder-Drehungen (für Input-Gain)"""
        page = device.selected_page
        
        # Bestimme Wing-Channel basierend auf Seite
        if page == 0 and ch <= 4:
            wing_ch = ch
        elif page == 1:
            wing_ch = ch + 8
        elif page == 2:
            wing_ch = ch + 16
        else:
            return  # Keine Input-Kanäle auf dieser Seite
        
        # Finde Input-Quelle für diesen Channel
        source = wing.getChannelInputSource(wing_ch)
        if not source:
            logger.warning(f"Keine Input-Quelle für CH{wing_ch} gefunden")
            return
        
        grp, num = source
        
        # Aktueller Gain
        current_gain = device.channels[ch].gain_db
        
        # Neuer Gain berechnen
        if delta > 0:
            new_gain = current_gain + (GAIN_STEP)
        else:
            new_gain = current_gain - (GAIN_STEP)
        new_gain = max(GAIN_MIN, min(GAIN_MAX, new_gain))
        
        # An Wing senden
        if wing.setGain(grp, num, new_gain):
            device.channels[ch].gain_db = new_gain
            
            # LED-Ring aktualisieren (0-60dB → 0-11 LEDs)
            ring_value = int((new_gain / GAIN_MAX) * 11)
            device.set_encoder_ring(ch, ring_value, mode=1)
    
    # Fader Handler
    def on_fader(ch: int, val: int):
        """Verarbeitet Fader-Bewegungen"""
        # Deadband-Filter
        last_val = device.channels[ch].last_stable_value
        if abs(val - last_val) < FADER_DEADBAND:
            return
        
        device.channels[ch].last_stable_value = val
        page = device.selected_page
        
        # Page 0
        if page == 0:
            if ch <= 4:
                wing.setFader("ch", ch, val)
            elif ch == 5: 
                wing.setFader("mtx", 1, val)
            elif ch == 6:
                wing.setFader("mtx", 2, val)
            elif ch == 7:
                wing.setFader("mtx", 3, val)
            elif ch == 8:
                for i in range(1, 5, 1):
                    dmx.sendDMX(i, val)
        
        # Page 1: Channels 9-16
        elif page == 1:
            wing.setFader("ch", ch + 8, val)
        
        # Page 2: Channels 17-24
        elif page == 2:
            wing.setFader("ch", ch + 16, val)

        elif page == 3:
            dmx.sendDMX(ch, val)
            device.set_fader(ch, val)
    
    # Callbacks registrieren
    device.on_button_press = on_button
    device.on_fader_move = on_fader
    device.on_encoder_turn = on_encoder  # Neu!
    
    # Sync Loop starten
    sync_thread = threading.Thread(
        target=sync_loop,
        args=(device, wing, dmx),
        daemon=True,
        name="SyncLoop"
    )
    sync_thread.start()
    
    logger.info("=== System läuft ===")
    logger.info("Drücke Ctrl+C zum Beenden")
    
    # Hauptloop mit periodischer Verbindungsprüfung
    last_check = time.time()
    # ====== NEU: X-Touch Check Timer ======
    last_xtouch_check = time.time()
    xtouch_reconnect_attempts = 0
    try:
        while True:
            time.sleep(1)
            if time.time() - last_xtouch_check >= 2.0:
                if not device.is_connected():
                    logger.warning("X-Touch antwortet nicht mehr!")
                    
                    if device.reconnect("xtouch"):
                        logger.info("✓ X-Touch wieder verbunden")
                        xtouch_reconnect_attempts = 0
                        
                        # Aktualisiere Zustand nach Reconnect
                        device.update_page_faders(device.selected_page, wing, dmx)
                        time.sleep(0.05)
                        device.update_page_mutes(device.selected_page, wing)
                        
                        # Gain-Ringe wiederherstellen
                        if device.selected_page in [0, 1, 2]:
                            for i in range(1, 9):
                                if device.selected_page == 0 and i <= 4:
                                    wing_ch = i
                                elif device.selected_page == 1:
                                    wing_ch = i + 8
                                elif device.selected_page == 2:
                                    wing_ch = i + 16
                                else:
                                    continue
                                
                                source = wing.getChannelInputSource(wing_ch)
                                if source:
                                    grp, num = source
                                    gain = wing.getGain(grp, num)
                                    if gain is not None:
                                        device.channels[i].gain_db = gain
                                        ring_value = int((gain / GAIN_MAX) * 11)
                                        device.set_encoder_ring(i, ring_value, mode=1)
                    else:
                        xtouch_reconnect_attempts += 1
                        if xtouch_reconnect_attempts >= 5:
                            logger.error("X-Touch Reconnect fehlgeschlagen nach 5 Versuchen")
                            xtouch_reconnect_attempts = 0
                
                last_xtouch_check = time.time()
            
            # Periodische Verbindungsprüfung
            if time.time() - last_check >= CONNECTION_CHECK_INTERVAL:
                logger.debug("Periodische Verbindungsprüfung...")
                if not wing.check_connection():
                    logger.warning("Verbindung verloren! Versuche Reconnect...")
                    for i in range (1, 513,1 ):
                        dmx.sendDMX(i, 0)
                last_check = time.time()
                if not dmx.connect():
                    logger.warning("DMX Fehler, verbindung verloren!")
    
    except KeyboardInterrupt:
        logger.info("Beende Programm...")
        device.close()
        wing.close()
        dmx.close()
        logger.info("Programm beendet")


if __name__ == "__main__":
    main()
