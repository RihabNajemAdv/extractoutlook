# Importations de modules nécessaires
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox
import requests
import pandas as pd
import threading
import re
import queue
from datetime import datetime
import openpyxl


# Création d'une file d'attente globale pour gérer les mises à jour de l'interface utilisateur
update_queue = queue.Queue()

# Variable pour contrôler l'arrêt du processus
is_terminating = False

# Liste globale pour stocker les événements de tous les utilisateurs
global_events_list = []

# Configuration pour multiples tenants (clients) pour l'API Microsoft Graph
tenants_info = [
    {
        "client_id": "155075e7-b710-4b4b-93cc-e66fcaee8cc3",
        "tenant_name": "a51ca36c-ae36-4f02-8c3a-8e501e5e8572",
        "client_secret": "fdk8Q~mLZL922enzlgJiwbfvqGKOA3ftznJeebYV"
    },
    {
        "client_id": "599f35e1-cc61-4e9e-a3ba-d8e61df03e4f",
        "tenant_name": "bb8adee9-c558-429c-ab3b-dbbaa59496fa",
        "client_secret": "DRG8Q~IE1Phmo3_q3ZBYbs4nqGvhZLSrPBKSsbwL"
    }
    # Ajoutez d'autres tenants ici si nécessaire
]


# Classe ScrollingFrame pour créer un cadre avec défilement dans l'interface Tkinter
class ScrollingFrame(tk.Frame):
    def __init__(self, master):
        super().__init__(master)

        # Création d'un canevas et d'une barre de défilement
        self.canvas = tk.Canvas(self)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        # Configuration de la fonctionnalité de défilement
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        # Ajout du cadre défilant au canevas
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Positionnement du canevas et de la barre de défilement dans le cadre
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")


# Classe pour créer une fenêtre avec une barre de progression et un affichage des données
class ProgressBarWindow:
    def __init__(self, master):
        # Initialisation de la fenêtre principale
        self.master = master
        self.master.title("Progression de la récupération des événements")
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Création d'un widget Treeview pour afficher les données
        self.tree = ttk.Treeview(self.master, columns=("Email", "Count", "Status"), show='headings')
        self.tree.heading("Email", text="Email")
        self.tree.heading("Count", text="Nombre d'événements")
        self.tree.heading("Status", text="Statut")
        self.tree.column("Email", width=200)
        self.tree.column("Count", width=100)
        self.tree.column("Status", width=100)

        # Barre de défilement pour le Treeview
        scrollbar = ttk.Scrollbar(self.master, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        # Dictionnaire pour garder la trace des lignes de l'utilisateur dans Treeview
        self.user_rows = {}

    def on_closing(self):
        # Gestion de la fermeture de la fenêtre
        global is_terminating
        is_terminating = True
        self.master.destroy()

    def add_user(self, upn):
        # Ajoute un nouvel utilisateur dans le Treeview
        row_id = self.tree.insert("", "end", values=(upn, "0", "Not Started"))
        self.user_rows[upn] = row_id  # Stocke la référence de la ligne pour cet utilisateur

    def process_queue(self):
        # Traite les éléments dans la file d'attente pour mettre à jour l'interface utilisateur
        while not update_queue.empty():
            upn, count, status = update_queue.get()  # Récupère les informations de la file d'attente
            self.update_user(upn, count, status)  # Met à jour les informations de l'utilisateur dans le Treeview
        self.master.after(100, self.process_queue)  # Ré-exécute cette méthode toutes les 100 millisecondes

    def update_user(self, upn, count, status):
        # Met à jour les informations d'un utilisateur spécifique dans le Treeview
        if upn not in self.user_rows:
            self.add_user(upn)  # Si l'utilisateur n'est pas encore dans le Treeview, l'ajoute
        self.tree.item(self.user_rows[upn], values=(upn, str(count), status))  # Met à jour la ligne de l'utilisateur

    def check_completion(self):
        # Vérifie si tous les utilisateurs ont terminé leur traitement
        all_done = True
        for row_id in self.user_rows.values():
            if self.tree.item(row_id)["values"][2] not in ["Completed", "Failed"]:
                all_done = False
                break
        if all_done:
            self.master.destroy()  # Ferme la fenêtre si tous les utilisateurs ont terminé


# Fonction pour récupérer les données de l'API Microsoft Graph
def fetch_data(start_date, end_date, base_file_path, window):
    threads = []
    for tenant in tenants_info:
        token = auth_o365(tenant)
        if token:
            headers = {'Authorization': f'Bearer {token["access_token"]}'}
            user_api = "https://graph.microsoft.com/v1.0/users"
            all_users = []

            while user_api:
                response = requests.get(user_api, headers=headers)
                if response.status_code != 200:
                    print("Erreur lors de la récupération des utilisateurs")
                    break

                data = response.json()
                all_users.extend(data.get('value', []))
                user_api = data.get('@odata.nextLink')

            filtered_users = [user for user in all_users if user_filter(user)]

            for user in filtered_users:
                thread = threading.Thread(target=process_user,
                                          args=(user, token, start_date, end_date, base_file_path, window))
                threads.append(thread)
                thread.start()

    for thread in threads:
        thread.join()

    # Exportation des données globales après la fin de tous les threads
    if global_events_list:
        columns_order = [
            'User', 'Subject', 'Code_Extracted', 'AttendeesCount', 'OrganizerName',
            'OrganizerEmail', 'Start', 'End', 'Duration', 'TimeZone', 'AllDayEvent',
            'Categories', 'WebLink', 'LastModifiedTime', 'OriginalStartTimeZone', 'OriginalEndTimeZone'
        ]
        # Créer un DataFrame en spécifiant l'ordre des colonnes
        global_df = pd.DataFrame(global_events_list, columns=columns_order)
        # Exporter le DataFrame dans un fichier Excel
        excel_path = f"{base_file_path}/global_events.xlsx"
        global_df.to_excel(excel_path, index=False, engine='openpyxl')
        print(f"Tous les événements exportés dans un fichier Excel à l'emplacement suivant : {excel_path}")


# Fonction pour obtenir un token d'accès pour l'API Microsoft Graph
def auth_o365(tenant):
    token_url = f"https://login.microsoftonline.com/{tenant['tenant_name']}/oauth2/v2.0/token"
    token_data = {
        'grant_type': 'client_credentials',
        'client_id': tenant['client_id'],
        'client_secret': tenant['client_secret'],
        'scope': 'https://graph.microsoft.com/.default'
    }
    token_r = requests.post(token_url, data=token_data)
    return token_r.json() if token_r.status_code == 200 else None


# Fonction pour filtrer les utilisateurs selon des critères spécifiques
def user_filter(user):
    upn = user.get('userPrincipalName', '')
    return (
                "@adv-sud.fr" in upn or "@foodsp.fr" in upn or "@adventae.com" in upn) and not "#EXT#" in upn and upn != "direction@adv-sud.fr"


# Fonction pour traiter chaque utilisateur et récupérer ses événements de calendrier
# Fonction pour traiter chaque utilisateur et récupérer ses événements de calendrier
def process_user(user, token, start_date, end_date, base_file_path, window):
    global is_terminating, update_queue, global_events_list

    upn = user.get('userPrincipalName')
    # Initialiser code_extracted et total_duration avant leur utilisation
    code_extracted = None
    total_duration = 0
    update_queue.put((upn, 0, "Starting"))

    calendar_api = f"https://graph.microsoft.com/v1.0/users/{upn}/calendarView?startDateTime={start_date}T00:00:00&endDateTime={end_date}T00:00:00&$top=99"
    event_list = []
    success = True

    while True:
        if is_terminating:
            update_queue.put((upn, len(event_list), "Interrupted"))
            break

        events_response = requests.get(calendar_api, headers={'Authorization': f'Bearer {token["access_token"]}'})
        if events_response.status_code != 200:
            print(f"Erreur lors de la récupération des événements pour {upn}")
            update_queue.put((upn, len(event_list), "Failed"))
            success = False
            break

        events_data = events_response.json()
        new_events = events_data.get('value', [])

        # Traitement de chaque événement récupéré
        for event in new_events:
            subject = event.get('subject', '')
            start_time = event.get('start', {}).get('dateTime')
            end_time = event.get('end', {}).get('dateTime')


            # Si les heures de début et de fin sont présentes, effectuer le calcul de la durée
            if start_time and end_time:
                start = pd.to_datetime(start_time)
                end = pd.to_datetime(end_time)
                total_duration = 0

                # Vérifie si l'événement commence et se termine le même jour
                if start.date() == end.date():
                    # Calcul de la durée pour un événement d'une seule journée
                    total_duration = min((end - start).total_seconds() / 3600, 8)
                else:
                    # Calcul pour le premier jour
                    end_of_first_day = start.replace(hour=23, minute=59, second=59, microsecond=0)
                    total_duration += min((end_of_first_day - start).total_seconds() / 3600, 8)

                    # Calcul pour les jours complets entre le premier et le dernier jour
                    current_day = start + pd.Timedelta(days=1)
                    while current_day.date() < end.date():
                        total_duration += 8  # Plafond de 8h par jour
                        current_day += pd.Timedelta(days=1)

                    # Calcul pour le dernier jour
                    if end.date() == current_day.date():
                        start_of_last_day = current_day.replace(hour=0, minute=0, second=0, microsecond=0)
                        total_duration += min((end - start_of_last_day).total_seconds() / 3600, 8)

                # Ajout de la durée calculée à l'événement
                event['Duration'] = total_duration

            # Recherche d'un code spécifique dans le sujet de l'événement à l'aide d'une expression régulière
            if subject:
                match = re.search(r"\[(\d{4,7})\]", subject)
                code_extracted = match.group(1) if match else None
            else:
                code_extracted = None

            # Construction d'un dictionnaire avec les données de l'événement pour l'ajouter à la liste
            event_dict = {
                'User': upn,
                'Subject': subject,
                'Code_Extracted': code_extracted,
                'AttendeesCount': len(event.get('attendees', [])),
                'OrganizerName': event.get('organizer', {}).get('emailAddress', {}).get('name'),
                'OrganizerEmail': event.get('organizer', {}).get('emailAddress', {}).get('address'),
                'Start': start_time,
                'End': end_time,
                'Duration': total_duration,
                'TimeZone': event.get('start', {}).get('timeZone'),
                'AllDayEvent': event.get('isAllDay'),
                'Categories': ', '.join(event.get('categories', [])),
                'WebLink': event.get('webLink'),
                'LastModifiedTime': event.get('lastModifiedDateTime'),
                'OriginalStartTimeZone': event.get('originalStartTimeZone', ''),
                'OriginalEndTimeZone': event.get('originalEndTimeZone', '')
            }
            global_events_list.append(event_dict)

        # Mise à jour de la file d'attente avec l'état actuel de traitement
        update_queue.put((upn, len(event_list), "Ongoing"))

        # Vérification s'il y a plus d'événements à récupérer
        if "@odata.nextLink" in events_data:
            calendar_api = events_data["@odata.nextLink"]  # Mise à jour de l'URL pour la prochaine requête
        else:
            update_queue.put((upn, len(event_list), "Completed"))  # Marquer le traitement comme terminé
            break

    # Si le traitement a réussi et qu'il y a des événements, les sauvegarder dans un fichier CSV
    if success and event_list:
        df = pd.DataFrame(event_list)  # Création d'un DataFrame avec les événements
        df.to_csv(f"{base_file_path}/{upn.replace('@', '_').replace('.', '_')}_events.csv", index=False)  # Sauvegarde en CSV
        print(f"Événements exportés pour {upn}")
    elif success:
        # Si le traitement a réussi mais qu'il n'y a pas d'événements, afficher un message
        print(f"Aucun événement à exporter pour {upn}")

# Fonction pour filtrer les utilisateurs selon des critères spécifiques
def user_filter(user):
    # Récupère le nom principal de l'utilisateur (UPN)
    upn = user.get('userPrincipalName', '')
    # Retourne True si l'UPN correspond aux critères spécifiés, False sinon
    # Ici, on filtre pour inclure uniquement certains domaines et exclure les comptes externes
    return ("@adv-sud.fr" in upn or "@foodsp.fr" in upn or "@adventae.com" in upn) and not "#EXT#" in upn and upn != "direction@adv-sud.fr"

# Fonction principale du script
def main():
    # Demande à l'utilisateur de saisir la date de début et de fin pour la récupération des données
    start_date = simpledialog.askstring("Input", 'Entrez la date de début (yyyy-mm-dd)')
    end_date = simpledialog.askstring("Input", 'Entrez la date de fin (yyyy-mm-dd)')
    # Demande à l'utilisateur de sélectionner un dossier pour enregistrer les fichiers CSV
    base_file_path = filedialog.askdirectory()

    # Vérifie que les entrées de l'utilisateur sont valides
    if not start_date or not end_date or not base_file_path:
        messagebox.showerror("Erreur", "Date de début, date de fin et emplacement du fichier sont requis.")
        return

    # Création de la fenêtre principale Tkinter
    root = tk.Tk()
    # Création de la fenêtre de barre de progression
    window = ProgressBarWindow(root)
    # Démarre la vérification périodique de la file d'attente
    root.after(100, window.process_queue)
    # Démarre un thread séparé pour récupérer les données sans bloquer l'interface utilisateur
    threading.Thread(target=fetch_data, args=(start_date, end_date, base_file_path, window)).start()
    # Lance la boucle principale de Tkinter
    root.mainloop()


# Vérifie si le script est exécuté directement (et non importé comme un module)
if __name__ == "__main__":
    main()  # Exécute la fonction principale