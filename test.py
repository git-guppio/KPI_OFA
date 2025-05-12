import pandas as pd
import os
from typing import List, Dict, Optional
from collections import Counter
import pyperclip
import time

# Contenuto di test che simula i dati problematici della clipboard
test_content = """-------------------------------------------------------------------------------------------------------------------------------------------------
|  Avviso    |Mod. il   |Data      |Descrizione                             |Tp.|Sede tecnica        |St.sist.      |Ordine      |Pse|Sis.Legacy|
-------------------------------------------------------------------------------------------------------------------------------------------------
|  1100103572|28.04.2025|31.03.2025|819 abb error de encoder                |Z1 |MXW-MXPA-X4-18-GE-ES|MELA          |            |MX |OFA       |
|  1100103640|28.04.2025|01.04.2025|767; Mal funcionamiento subsistema  pitc|Z1 |MXW-MXPA-X5-05-PS-HG|MELA          |            |MX |OFA       |
|  1100103637|28.04.2025|01.04.2025|1830; Diferencia excesiva pala 2        |Z1 |MXW-MXPA-X6-21-PS-HG|MELA          |            |MX |OFA       |
|  1100103876|28.04.2025|11.04.2025|
3038 Disparo disyuntores refrigeración |Z1 |MXW-MXPA-X7-28-CF-RF|MELA          |            |MX |OFA       |
|  1200790761|28.04.2025|01.04.2025|4082 error cabecera gh                  |Z2 |MXW-MXPA-X1-31-HG-ES|MELA          |            |MX |OFA       |
|  1200797742|05.05.2025|20.04.2025|WTG 21 783  Cables untwisting           |Z2 |ZAW-ZAW4-X3-21-SI-ME|FCAN MECO     |            |ZA |POM       |
|  1200799069|          |25.04.2025|WTG 04 2931 GearOilLevelTooLow M:
      |Z2 |ZAW-ZAW5-X7-04-ML-RF|MAPE          |            |ZA |POM       |
|  1500179021|01.04.2025|06.02.2025|Inspección de palas WTG 19 PET          |Z5 |CLW-CLWL-X3-19-RT-PA|MECO ORAT     |210001163962|CL |          |
|  1200796096|14.04.2025|11.04.2025|falla under voltage modulefailure INV-G3|Z2 |CLS-CLSB-07-03-IT   |MECO ORAT     |240000489332|CL |OFA       |
|  1200796097|11.04.2025|11.04.2025|
3038 Disparo disyuntores refrigeración/|Z2 |MXW-MXPA-X7-28-CF-RF|MELA ORAT     |240000489333|MX |OFA       |
|  1200796100|11.04.2025|11.04.2025|
3038 Disparo disyuntores refrigeración |Z2 |MXW-MXPA-X7-28-CF-RF|MELA ORAT     |240000489335|MX |OFA       |
|  1200796101|29.04.2025|29.04.2025|Falla Interna SI 30-SC15                |Z2 |COS-COFD-15-P1-30-01|MECO ORAT     |240000489336|CO |OFA       |"""




def fix_clipboard_table_content(result: str) -> tuple[bool, str]:
    """
    Elabora il contenuto della tabella proveniente dalla clipboard.
    A partire dalla quarta riga, verifica che ogni riga inizi con il carattere '|'.
    Se una riga non inizia con '|', unisce il suo contenuto alla riga precedente.
    
    Args:
        result: Il contenuto della clipboard da elaborare
        
    Returns:
        tuple: (successo, contenuto_elaborato)
            - successo (bool): True se l'elaborazione è andata a buon fine, False altrimenti
            - contenuto_elaborato (str): Il contenuto elaborato se successo, altrimenti il contenuto originale
    """
    try:
        # Dividi il testo in righe
        lines = result.split('\n')
        
        # Verifica se ci sono almeno 4 righe
        if len(lines) < 4:
            print("Il contenuto ha meno di 4 righe. Impossibile procedere.")
            return False, result
        
        # Preserva le prime 3 righe
        processed_lines = lines[:3]
        
        # Contatore delle sostituzioni
        replacements_count = 0
        
        # Elabora a partire dalla quarta riga
        i = 3
        while i < len(lines):
            current_line = lines[i]
            
            # Verifica se la riga inizia con '|'
            if not current_line.startswith('|'):
                # Verifica se c'è una riga precedente da modificare
                if i > 0 and len(processed_lines) > 0:
                    prev_line = processed_lines[-1]
                    
                    # Trova l'ultima occorrenza di '|' nella riga precedente
                    last_pipe_index = prev_line.rfind('|')
                    
                    if last_pipe_index != -1:
                        # Estrai la parte iniziale fino all'ultimo '|'
                        prefix = prev_line[:last_pipe_index + 1]
                        
                        # Unisci con la riga corrente
                        merged_line = prefix + " " + current_line.strip()
                        
                        # Sostituisci la riga precedente
                        processed_lines[-1] = merged_line
                        
                        replacements_count += 1
                    else:
                        # Se non troviamo '|' nella riga precedente, semplicemente aggiungi la riga corrente
                        processed_lines.append(current_line)
                else:
                    # Se non c'è una riga precedente, aggiungi semplicemente la riga
                    processed_lines.append(current_line)
            else:
                # La riga inizia con '|', aggiungila normalmente
                processed_lines.append(current_line)
            
            # Passa alla riga successiva
            i += 1
        
        # Ricostruisci il testo elaborato
        processed_result = '\n'.join(processed_lines)
        
        # Stampa il numero di sostituzioni
        print(f"Elaborazione completata. Sono state eseguite {replacements_count} sostituzioni.")
        
        # Restituisci il risultato dell'elaborazione
        return True, processed_result
        
    except Exception as e:
        print(f"Errore durante l'elaborazione del contenuto: {str(e)}")
        return False, result

def handle_duplicate_headers(headers: List[str]) -> List[str]:
    """
    Gestisce le intestazioni duplicate aggiungendo un postfisso numerico
    e rinomina le intestazioni vuote con il prefisso "Unnamed_"
    
    Args:
        headers: Lista delle intestazioni originali
        
    Returns:
        Lista delle intestazioni con postfissi per i duplicati e rinominate per le vuote
    """
    # Inizializziamo il contatore per le colonne senza nome
    unnamed_counter = 0
    
    # Per prima cosa, rinominiamo le intestazioni vuote
    processed_headers = []
    for header in headers:
        # Verifica se l'header è vuoto (dopo strip)
        if not header or header.strip() == "":
            # Assegna un nome "Unnamed_N"
            processed_headers.append(f"Unnamed_{unnamed_counter}")
            unnamed_counter += 1
        else:
            # Applica strip e aggiungi l'header non vuoto
            processed_headers.append(header.strip())
    
    # Ora gestiamo i duplicati
    header_counts = Counter()
    unique_headers = []
    
    for header in processed_headers:
        # Se l'header è già stato visto
        if header in header_counts:
            # Incrementa il contatore e aggiungi il postfisso
            header_counts[header] += 1
            unique_headers.append(f"{header}_{header_counts[header]}")
        else:
            # Prima occorrenza dell'header
            header_counts[header] = 0
            unique_headers.append(header)
    
    return unique_headers  

def clean_and_load_data(file_path):
    """
    Legge un file di testo con campi separati da pipe, corregge le occorrenze extra di pipe
    nella colonna Descrizione, e carica i dati in un DataFrame.
    
    Args:
        file_path: Percorso del file di testo da analizzare
    
    Returns:
        DataFrame Pandas con i dati strutturati
    """
    # Leggi il file
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    # Filtra le righe vuote e quelle che contengono solo trattini
    filtered_lines = []
    for line in lines:
        line = line.strip()
        if line and not all(c in '-|' for c in line):
            filtered_lines.append(line)
    
    if not filtered_lines:
        print("Nessuna riga valida trovata nel file")
        return None
    
    # La prima riga è l'intestazione
    original_headers = filtered_lines[0]
    list_original_headers = original_headers.strip().split('|')
    print(f"Intestazioni originali: {list_original_headers} - lunghezza: {len(list_original_headers)}")

    # Gestisci gli header duplicati
    list_unique_headers = handle_duplicate_headers(list_original_headers)
    print(f"Intestazioni uniche: {list_unique_headers} - lunghezza: {len(list_unique_headers)}")

    expected_pipes = original_headers.count('|')
    
    # Creo un array per memorizzare le righe corrette
    corrected_lines = [original_headers]
    # Correggi le righe con occorrenze extra di pipe
    for line in filtered_lines[1:]:
        actual_pipes = line.count('|')
        
        if actual_pipes > expected_pipes:
            # Calcola quanti pipe extra ci sono
            extra_pipes = actual_pipes - expected_pipes
            
            # Trova l'indice del quarto pipe (che precede la colonna Descrizione)
            pipe_indices = [i for i, char in enumerate(line) if char == '|']
            if len(pipe_indices) >= 4:
                description_start = pipe_indices[3] + 1
                description_end = pipe_indices[4 + extra_pipes]
                
                # Estrai la descrizione
                description = line[description_start:description_end]
                
                # Sostituisci i pipe nella descrizione con trattini
                corrected_description = description.replace('|', '-')
                
                # Ricostruisci la riga
                corrected_line = line[:description_start] + corrected_description + line[description_end:]
                
                # Verifica che la correzione abbia funzionato
                if corrected_line.count('|') == expected_pipes:
                    corrected_lines.append(corrected_line)
                else:
                    # Fallback: se la correzione non ha funzionato, rimuovi tutti i pipe extra
                    parts = line.split('|')
                    if len(parts) > expected_pipes + 1:
                        # Unisci i campi extra nella Descrizione (indice 4 considerando i separatori)
                        merged_description = '-'.join(parts[4:4+extra_pipes+1])
                        new_parts = parts[:4] + [merged_description] + parts[4+extra_pipes+1:]
                        corrected_line = '|'.join(new_parts)
                        corrected_lines.append(corrected_line)
                    else:
                        print(f"Impossibile correggere la riga: {line}")
            else:
                print(f"La riga non ha abbastanza separatori pipe: {line}")
        else:
            # La riga è già corretta o ha meno pipe del previsto
            corrected_lines.append(line)
    
    # Ora che abbiamo corretto le righe, le convertiamo in DataFrame
    data_rows = []
    for line in corrected_lines:
        fields = [field.strip() for field in line.split('|')]
        # Rimuovi i campi vuoti all'inizio e alla fine (dovuti al pipe iniziale/finale)
        # fields = [field for field in fields if field]
        data_rows.append(fields)
    
    # Estrai gli header e i dati
    headers = list_unique_headers
    data = data_rows[1:]
    
    # Crea il DataFrame
    df = pd.DataFrame(data, columns=headers)     
          
    # Rimuoviamo le colonne senza intestazione
    unnamed_cols = [col for col in df.columns if col == '' or pd.isna(col) or str(col).startswith('Unnamed_')]
    df_cleaned = df.drop(columns=unnamed_cols)
    
    # Reset dell'indice
    df_cleaned = df_cleaned.reset_index(drop=True)

    
    return df_cleaned

def main():

    # Simulazione del contenuto della clipboard
    pyperclip.copy(test_content)

    # Ottieni il contenuto della clipboard
    clipboard_content = pyperclip.paste()
    time.sleep(0.1)
    # Elabora il contenuto
    success, fixed_content = fix_clipboard_table_content(clipboard_content)

    if success:
        # Se l'elaborazione è riuscita, aggiorna il contenuto nella clipboard
        pyperclip.copy(fixed_content)
        print("Il contenuto della clipboard è stato corretto e aggiornato.")
    else:
        print("Non è stato possibile correggere il contenuto della clipboard.")


if __name__ == "__main__":
    main()