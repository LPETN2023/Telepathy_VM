import sys
import subprocess
import os
import json
from datetime import datetime, timezone, timedelta
import pandas as pd
import time
import re
from telethon.sync import TelegramClient
import asyncio
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

def get_user_info():
    username = input("Entrez votre nom d'utilisateur Telegram (sans @) : ")
    phone = input("Entrez votre numéro de téléphone Telegram (ex: +5511999999999) : ")
    api_id = input("Entrez votre API ID (depuis https://my.telegram.org/apps) : ")
    api_hash = input("Entrez votre API hash (depuis https://my.telegram.org/apps) : ")
    return {"username": username, "phone": phone, "api_id": api_id, "api_hash": api_hash}

def save_user_info(user_info):
    with open('telegram_user_info.json', 'w') as f:
        json.dump(user_info, f)

def load_user_info():
    if os.path.exists('telegram_user_info.json'):
        with open('telegram_user_info.json', 'r') as f:
            return json.load(f)
    return None

async def download_media(message, media_directory):
    if message.media:
        try:
            media_path = await message.download_media(file=media_directory)
            return os.path.basename(media_path) if media_path else "Download failed"
        except Exception as e:
            print(f"Error downloading media: {e}")
            return "Download error"
    else:
        return "No media"

def remove_unsupported_characters(text):
    valid_xml_chars = (
        "[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD"
        "\U00010000-\U0010FFFF]"
    )
    cleaned_text = re.sub(valid_xml_chars, '', text)
    return cleaned_text

def format_time(seconds):
    days = seconds // 86400
    hours = (seconds % 86400) // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f'{int(days):02}:{int(hours):02}:{int(minutes):02}:{int(seconds):02}'

def print_progress(t_index, message_id, start_time, max_t_index):
    elapsed_time = time.time() - start_time
    current_progress = t_index / (t_index + message_id) if (t_index + message_id) <= max_t_index else t_index / max_t_index
    percentage = current_progress * 100
    estimated_total_time = elapsed_time / current_progress
    remaining_time = estimated_total_time - elapsed_time
    elapsed_time_str = format_time(elapsed_time)
    remaining_time_str = format_time(remaining_time)
    print(f'Progression : {percentage:.2f}% | Temps écoulé : {elapsed_time_str} | Temps restant : {remaining_time_str}')

async def main():
    user_info = load_user_info()
    if user_info:
        use_existing_input = input(f"Voulez-vous utiliser le profil existant pour {user_info['username']}? (oui/non) [par défaut : oui] : ").lower()
        use_existing = use_existing_input if use_existing_input else "oui"
        if use_existing != 'oui':
            os.remove('telegram_user_info.json')
            user_info = get_user_info()
            save_user_info(user_info)
    else:
        user_info = get_user_info()
        save_user_info(user_info)

    username = user_info['username']
    phone = user_info['phone']
    api_id = user_info['api_id']
    api_hash = user_info['api_hash']

    print("\nConfiguration des paramètres de scraping")

    while True:
        channel = input("Entrez le nom du canal ou du compte à scraper (au format https://t.me/xxxx ou @xxxx) : ").strip()
        if channel:
            break
        print("Le nom du canal ne peut pas être vide. Veuillez entrer un nom valide.")

    while True:
        file_name = input("Nom du fichier de sortie : ").strip()
        file_name = file_name.replace(" ", "_")
        if file_name:
            break
        print("Le nom du fichier ne peut pas être vide. Veuillez entrer un nom valide.")

    current_date = datetime.now(timezone.utc)
    one_year_ago = current_date - timedelta(days=365)

    date_min_input = input(f"Date de début (YYYY-MM-DD) (par défaut : {one_year_ago.strftime('%Y-%m-%d')}) : ")
    date_min = datetime.fromisoformat(date_min_input).replace(tzinfo=timezone.utc) if date_min_input else one_year_ago

    date_max_input = input(f"Date de fin (YYYY-MM-DD) (par défaut : {current_date.strftime('%Y-%m-%d')}) : ")
    date_max = datetime.fromisoformat(date_max_input).replace(tzinfo=timezone.utc) if date_max_input else current_date

    key_search = input("Mot-clé de recherche (laisser vide pour tout extraire) : ")

    max_t_index_input = input("Nombre maximum de messages à scraper (par défaut : 1000000) : ")
    max_t_index = int(max_t_index_input) if max_t_index_input else 1000000

    time_limit_input = input("Limite de temps en secondes (par défaut : 21600) : ")
    time_limit = int(time_limit_input) if time_limit_input else 21600

    project_directory = f'Telegram_{file_name}'
    os.makedirs(project_directory, exist_ok=True)

    media_directory = os.path.join(project_directory, 'media')
    os.makedirs(media_directory, exist_ok=True)

    data = []
    t_index = 0
    start_time = time.time()

    if t_index >= max_t_index:
        print("Le nombre maximum de messages a été atteint.")
    elif time.time() - start_time > time_limit:
        print("La limite de temps a été atteinte.")
    else:
        loop_start_time = time.time()
        
        try:
            c_index = 0
            async with TelegramClient(username, api_id, api_hash) as client:
                async for message in client.iter_messages(channel, search=key_search):
                    try:
                        if date_min <= message.date <= date_max:
                            media_file = await download_media(message, media_directory)
                            
                            emoji_string = ''
                            if message.reactions:
                                for reaction_count in message.reactions.results:
                                    emoji = reaction_count.reaction.emoticon
                                    count = str(reaction_count.count)
                                    emoji_string += emoji + " " + count + " "
                            
                            date_time = message.date.strftime('%Y-%m-%d %H:%M:%S')
                            cleaned_content = remove_unsupported_characters(message.text)
                            
                            data.append({
                                'Type': 'text',
                                'Group': channel,
                                'Author ID': message.sender_id,
                                'Content': cleaned_content,
                                'Date': date_time,
                                'Message ID': message.id,
                                'Author': message.post_author,
                                'Views': message.views,
                                'Reactions': emoji_string,
                                'Shares': message.forwards,
                                'Media': 'True' if message.media else 'False',
                                'Media File': media_file,
                                'Url': f'https://t.me/{channel}/{message.id}'.replace('@', ''),
                            })
                            
                            c_index += 1
                            t_index += 1
                            
                            print(f'{"-" * 80}')
                            print_progress(t_index, message.id, start_time, max_t_index)
                            current_max_id = min(c_index + message.id, max_t_index)
                            print(f'De {channel}: {c_index:05} contenus sur {current_max_id:05}')
                            print(f'Id: {message.id:05} / Date: {date_time}')
                            print(f"Total: {t_index:05} contenus jusqu'à présent")
                            print(f'{"-" * 80}\n\n')
                            
                            if t_index % 1000 == 0:
                                backup_filename = os.path.join(project_directory, f'backup_{file_name}_until_{t_index:05}_{channel}_ID{message.id:07}.xlsx')
                                pd.DataFrame(data).to_excel(backup_filename, index=False, engine='openpyxl')
                            
                            if t_index >= max_t_index:
                                break
                            
                            if time.time() - start_time > time_limit:
                                break
                        
                        elif message.date < date_min:
                            break
                    
                    except Exception as e:
                        print(f'Erreur lors du traitement du message : {e}')
                
                print(f'\n\n##### {channel} terminé avec {c_index:05} posts #####\n\n')
        
        except Exception as e:
            print(f'Erreur pour {channel}: {e}')
        
        loop_end_time = time.time()
        loop_duration = loop_end_time - loop_start_time
        if loop_duration < 60:
            time.sleep(60 - loop_duration)

    print(f'\n{"-" * 50}\n#Terminé ! #{t_index:05} posts ont été scrapés !\n{"-" * 50}\n\n\n\n')

    df = pd.DataFrame(data)
    final_filename = os.path.join(project_directory, f'FINAL_{file_name}_with_{t_index:05}.xlsx')
    
    wb = Workbook()
    ws = wb.active
    for r, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c, value in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=value)
            if c == df.columns.get_loc('Media File') + 1 and r > 1 and value != "No media" and value != "Download failed" and value != "Download error":
                cell.hyperlink = os.path.join('media', value)
                cell.font = Font(color="0000FF", underline="single")

    wb.save(final_filename)

    print(f"Le fichier final a été sauvegardé sous le nom : {final_filename}")
    
    try:
        command = f"cp -r ./Telegram_{file_name} \"$(wslpath \"$(powershell.exe '[System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::DesktopDirectory)')\" | tr -d '\\r')\" && rm -R ./Telegram_{file_name}"
        subprocess.run(command, shell=True, check=True)
        print(f"Le dossier Telegram_{file_name} a été copié sur le Bureau et supprimé du répertoire courant.")
    except subprocess.CalledProcessError as e:
        print(f"Erreur lors de la copie/suppression du dossier : {e}")
        
asyncio.run(main())