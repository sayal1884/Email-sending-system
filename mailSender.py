import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from PIL import Image, ImageTk, ImageDraw, ImageFont
import re
import ttkthemes
from datetime import datetime, timedelta
import smtplib
import schedule
import time
import threading
import emoji
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from openpyxl import Workbook
import json

class EmailApp(tk.Tk):
    def __init__(self, width, height, title):
        super().__init__()
        self.title(title)
        self.schedule_file = "schedule.json"
        self.height = self.winfo_screenheight()
        self.geometry(f"{width}x{self.height}")
        
        # Application du thème breeze-dark
        self.style = ttkthemes.ThemedStyle(self)
        self.style.set_theme("plastik")

        # Listes pour les destinataires et les pièces jointes
        self.destinataires = []
        self.attachments = []

        try:
            banner_image = Image.open("none.png")  # Assurez-vous que l'image est dans le dossier 'assets'
            banner_width = width
            banner_height = int(height * 0.1)
            banner_image = banner_image.resize((banner_width, banner_height), Image.LANCZOS)

            # Adapter le titre sur la bannière avec une police personnalisée
            draw = ImageDraw.Draw(banner_image)
            # Utiliser une police TTF personnalisée
            font = ImageFont.truetype("assets/arial.ttf", size=40, encoding="unic")  # taille ajustée à 40pt
            bbox = draw.textbbox((0, 0), title, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            text_x = (banner_width - text_width) / 2
            text_y = (banner_height - text_height) / 2

            # Ajout d'une ombre pour le titre
            draw.text((text_x + 1, text_y + 1), title, font=font, fill="gray")
            draw.text((text_x, text_y), title, font=font, fill="black")

            banner_photo = ImageTk.PhotoImage(banner_image)
            self.banner_label = ttk.Label(self, image=banner_photo)
            self.banner_label.image = banner_photo
            self.banner_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)
        except Exception as e:
            print(f"Erreur lors du chargement de la bannière: {e}")

        # Configurer les styles pour ttk
        self.style.configure('TLabel', font=('Arial', 12))
        self.style.configure('TButton', font=('Arial', 12, 'bold'))
        self.style.configure('TEntry', font=('Arial', 12))
        self.style.configure('TCombobox', font=('Arial', 12))

        # Frame pour l'expediteur
        expediteur_frame = ttk.LabelFrame(self, text="Expediteur")
        expediteur_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        # Ajouter les destinataires
        label_expediteur = ttk.Label(expediteur_frame, text="expediteur : ")
        label_expediteur.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.entry_expediteur = ttk.Entry(expediteur_frame, width=30)
        self.entry_expediteur.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        label_password = ttk.Label(expediteur_frame, text="Mot de Passe d'application (google) : ")
        label_password.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.entry_password= ttk.Entry(expediteur_frame, width=30, show = '*')
        self.entry_password.grid(row= 0, column=3, padx=5, pady=5, sticky="ew")

        # Frame pour les destinataires
        destinataires_frame = ttk.LabelFrame(self, text="Destinataires")
        destinataires_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

        # Ajouter les destinataires
        label_destinataire = ttk.Label(destinataires_frame, text="Ajouter destinataire")
        label_destinataire.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.entry_destinataire = ttk.Entry(destinataires_frame, width=50)
        self.entry_destinataire.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        bouton_ajouter_destinataire = ttk.Button(destinataires_frame, text=emoji.emojize("Ajouter :inbox_tray:"), command=self.ajouter_destinataire)
        bouton_ajouter_destinataire.grid(row=0, column=2, padx=5, pady=5)

        # ComboBox pour afficher et supprimer les destinataires
        label_combobox = ttk.Label(destinataires_frame, text="Liste des destinataires")
        label_combobox.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.combo_destinataires = ttk.Combobox(destinataires_frame, values=self.destinataires, width=47)
        self.combo_destinataires.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        bouton_supprimer_destinataire = ttk.Button(destinataires_frame, text=emoji.emojize("Supprimer :outbox_tray:"), command=self.supprimer_destinataire)
        bouton_supprimer_destinataire.grid(row=1, column=2, padx=5, pady=5)

        # Frame pour le mail
        mail_frame = ttk.LabelFrame(self, text="Compose")
        mail_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

        # Sujet du mail
        label_sujet = ttk.Label(mail_frame, text="Sujet")
        label_sujet.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.entry_sujet = ttk.Entry(mail_frame, width=50)
        self.entry_sujet.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Message du mail
        label_message = ttk.Label(mail_frame, text="Message")
        label_message.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.text_message = tk.Text(mail_frame, height=10, width=50)
        self.text_message.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Bouton pour choisir les pièces jointes
        bouton_pieces_jointes = ttk.Button(mail_frame, text=emoji.emojize("Ajoutez des pièces jointes :paperclip:"), command=self.choisir_pieces_jointes)
        bouton_pieces_jointes.grid(row=2, column=1, padx=5, pady=5, sticky="e")

        # ComboBox pour afficher et supprimer les pièces jointes
        label_attachments = ttk.Label(mail_frame, text="Pièces jointes")
        label_attachments.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.combo_attachments = ttk.Combobox(mail_frame, values=self.attachments, width=47)
        self.combo_attachments.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        bouton_supprimer_piece = ttk.Button(mail_frame, text=emoji.emojize("Supprimer :wastebasket:"), command=self.supprimer_piece_jointe)
        bouton_supprimer_piece.grid(row=3, column=2, padx=5, pady=5)

        # Widgets pour programmer l'envoi du mail
        program_frame = ttk.LabelFrame(mail_frame, text="Programmer l'envoi")
        program_frame.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        self.combo_jour = ttk.Combobox(program_frame, values=["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"], width=10)
        self.combo_jour.grid(row=0, column=0, padx=5, pady=5)
        self.combo_heure = ttk.Combobox(program_frame, values=[f"{i:02d}" for i in range(24)], width=5)
        self.combo_heure.grid(row=0, column=1, padx=5, pady=5)
        self.combo_minute = ttk.Combobox(program_frame, values=[f"{i:02d}" for i in range(60)], width=5)
        self.combo_minute.grid(row=0, column=2, padx=5, pady=5)
        bouton_programmer = ttk.Button(program_frame, text=emoji.emojize("Programmer l'envoi :alarm_clock:"), command=self.programmer_envoi)
        bouton_programmer.grid(row=0, column=3, padx=5, pady=5)

        # Bouton pour envoyer le mail
        bouton_envoyer = ttk.Button(self, text=emoji.emojize("Envoyer :rocket:"), command=self.envoyer_mail)
        bouton_envoyer.grid(row=6, column=0, padx=10, pady=10, sticky="e")

        # Configurer grid pour redimensionner correctement
        self.columnconfigure(0, weight=1)
        destinataires_frame.columnconfigure(1, weight=1)
        mail_frame.columnconfigure(1, weight=1)

        self.load_schedule()
        self.time_format = self.detect_time_format()
        self.start_scheduler()

    def ajouter_destinataire(self):
        
        email = self.entry_destinataire.get()
        if self.verifier_email(email):
            self.destinataires.append(email)
            self.combo_destinataires['values'] = self.destinataires
            self.entry_destinataire.delete(0, 'end')
        else:
            messagebox.showerror("Erreur", "Adresse e-mail invalide. Veuillez entrer une adresse e-mail valide.")

    def supprimer_destinataire(self):
        email = self.combo_destinataires.get()
        if email in self.destinataires:
            self.destinataires.remove(email)
            self.combo_destinataires['values'] = self.destinataires
            self.combo_destinataires.set('')

    def verifier_email(self, email):
        # Vérifier la validité de l'adresse e-mail avec une expression régulière
        pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
        return re.match(pattern, email)

    def choisir_pieces_jointes(self):
        fichiers = filedialog.askopenfilenames(title="Choisir des pièces jointes")
        if fichiers:
            self.attachments.extend(fichiers)
            self.combo_attachments['values'] = self.attachments

    def supprimer_piece_jointe(self):
        piece = self.combo_attachments.get()
        if piece in self.attachments:
            self.attachments.remove(piece)
            self.combo_attachments['values'] = self.attachments
            self.combo_attachments.set('')

    def programmer_envoi(self):
        jour = self.combo_jour.get()
        heure = self.combo_heure.get()
        minute = self.combo_minute.get()
        
        # Vérification que toutes les informations nécessaires sont fournies
        if not (jour and heure and minute):
            messagebox.showerror("Erreur", "Veuillez sélectionner le jour, l'heure et la minute pour programmer l'envoi.")
            return

        # Calculer la date et l'heure de l'envoi programmé
        jours_semaine = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        jour_index = jours_semaine.index(jour)

        # Calculer le temps en secondes jusqu'à l'heure programmée
        now = datetime.now()
        diff_jours = (jour_index - now.weekday() + 7) % 7
        date_envoi = now + timedelta(days=diff_jours)
        date_envoi = date_envoi.replace(hour=int(heure), minute=int(minute), second=0, microsecond=0)

        # Calculer le délai en secondes
        delay = (date_envoi - now).total_seconds()

        # Programmer l'envoi
        self.after(int(delay * 1000), self.envoyer_mail_programme)
        self.save_schedule()

        messagebox.showinfo("Programmé", f"E-mail programmé pour être envoyé le {jour} à {heure}:{minute}.")

    def envoyer_mail_programme(self):
        sujet = self.entry_sujet.get()
        message = self.text_message.get("1.0", 'end')
        self.envoyer_mail()

    def envoyer_mail(self):
        for receiver in self.destinataires:
            message = MIMEMultipart()
            self.sender_email = self.entry_expediteur.get()
            message["From"] = self.entry_expediteur.get()
            message["To"] = receiver
            message["Subject"] = self.entry_sujet.get()
            message.attach(MIMEText(self.text_message.get("1.0", 'end'), "plain"))

            for file in self.attachments:
                with open(file, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(file)}")
                    message.attach(part)
            
            try:
                with smtplib.SMTP("smtp.gmail.com", 587) as server:
                    server.starttls()
                    server.login(self.sender_email,self.entry_password.get())
                    server.sendmail(self.sender_email, receiver, message.as_string())
                    print("Email envoyé avec succès!")
            except Exception as e:
                print(f"Erreur lors de l'envoi de l'email: {e}")
            if receiver == self.destinataires[-1]:
                messagebox.showinfo("Info", "E-mail envoyé avec succès !")

    def schedule_email(self):
        # Planifier l'envoi en fonction du jour sélectionné
        if self.selected_day.lower() == "monday":
            schedule.every().monday.at(self.selected_time).do(self.envoyer_mail)
        elif self.selected_day.lower() == "tuesday":
            schedule.every().tuesday.at(self.selected_time).do(self.envoyer_mail)
        elif self.selected_day.lower() == "wednesday":
            schedule.every().wednesday.at(self.selected_time).do(self.envoyer_mail)
        elif self.selected_day.lower() == "thursday":
            schedule.every().thursday.at(self.selected_time).do(self.envoyer_mail)
        elif self.selected_day.lower() == "friday":
            schedule.every().friday.at(self.selected_time).do(self.envoyer_mail)
        elif self.selected_day.lower() == "saturday":
            schedule.every().saturday.at(self.selected_time).do(self.envoyer_mail)
        elif self.selected_day.lower() == "sunday":
            schedule.every().sunday.at(self.selected_time).do(self.envoyer_mail)
        
        while True:
            schedule.run_pending()
            time.sleep(1)

    def start_scheduler(self):
        thread = threading.Thread(target=self.schedule_email)
        thread.start()

    def save_schedule(self):
        schedule_data = {
            "selected_day": self.combo_jour.get(),
            "selected_time": f"{self.combo_heure.get()}:{self.combo_minute.get()}",
            "receiver_email": str(self.destinataires),
            "file_path": str(self.attachments),
            "sender_email" : self.entry_expediteur.get(),
            "password" : self.entry_password.get(),
            "subject" : self.entry_sujet.get(),
            "message" : self.text_message.get("1.0", 'end')
        }
        self.selected_day = self.combo_jour.get()
        self.selected_time = f"{self.combo_heure.get()}:{self.combo_minute.get()}"
        with open(self.schedule_file, "w") as f:
            json.dump(schedule_data, f)

    def load_schedule(self):
        if os.path.exists(self.schedule_file):
            with open(self.schedule_file, "r") as f:
                schedule_data = json.load(f)
                self.selected_day = schedule_data.get("selected_day", "Monday")
                self.selected_time = schedule_data.get("selected_time", f"{self.combo_heure.get()}:{self.combo_minute.get()}")
                self.destinataires = eval(schedule_data.get("receiver_email", ""))
                self.attachments = eval(schedule_data.get("file_path", ""))
                self.entry_expediteur.insert("end",schedule_data.get("sender_email", ""))
                self.entry_password.insert("end",schedule_data.get("password", ""))
                self.entry_sujet.insert("end",schedule_data.get("subject", ""))
                self.text_message.insert("end", schedule_data.get("message", ""))

        self.combo_destinataires.config(values = self.destinataires)
        self.combo_attachments.config(values = self.attachments)

    def detect_time_format(self):
        time_string = time.strftime('%X')
        if 'AM' in time_string or 'PM' in time_string:
            return "12"
        else:
            return "24"

# Lancer l'application
app = EmailApp(width=800, height=700, title="Email Application")
app.mainloop()
