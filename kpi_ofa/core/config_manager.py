# kpi_ofa/core/config_manager.py

import os
import json
import logging

logger = logging.getLogger(__name__)

class ConfigManager:
    """
    Gestore della configurazione dell'applicazione KPI OFA.
    
    Questa classe è responsabile del caricamento, salvataggio e gestione
    delle impostazioni di configurazione dell'applicazione.
    
    Implementa il pattern Singleton per garantire un'unica istanza in tutta l'applicazione.
    """
    
    _instance = None
    
    def __new__(cls, config_file_path=None):
        """
        Implementazione del pattern Singleton.
        
        Garantisce che esista una sola istanza del ConfigManager nell'applicazione.
        
        Args:
            config_file_path (str, optional): Percorso del file di configurazione.
                                              Se non specificato, usa il valore predefinito.
        
        Returns:
            ConfigManager: L'istanza unica del ConfigManager.
        """
        if cls._instance is None:
            cls._instance = super(ConfigManager, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self, config_file_path=None):
        """
        Inizializza il gestore di configurazione.
        
        L'inizializzazione avviene solo la prima volta che viene creata un'istanza.
        
        Args:
            config_file_path (str, optional): Percorso del file di configurazione.
                                              Se non specificato, usa il valore predefinito.
        """
        # Se l'istanza è già stata inizializzata, esci
        if self._initialized:
            return
        
        # Importa qui per evitare dipendenze circolari
        from kpi_ofa.constants import configuration_json, default_config
        
        # Memorizza il percorso del file di configurazione
        self.config_file = config_file_path if config_file_path else configuration_json
        
        # Memorizza la configurazione predefinita
        self.default_config = default_config
        
        # Carica la configurazione
        self.config = self._load_config()
        
        # Segna l'istanza come inizializzata
        self._initialized = True
        
        logger.info(f"ConfigManager inizializzato. File di configurazione: {self.config_file}")
    
    def _load_config(self):
        """
        Carica la configurazione dal file JSON o utilizza valori predefiniti.
        
        Returns:
            dict: La configurazione caricata o quella predefinita in caso di errore.
        """
        # Se il file non esiste, utilizza la configurazione predefinita
        if not os.path.exists(self.config_file):
            logger.warning(f"File di configurazione non trovato: {self.config_file}. "
                          f"Utilizzo valori predefiniti.")
            return self.default_config.copy()
        
        try:
            # Leggi il file JSON
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            logger.info("Configurazione caricata con successo.")
            return config
        except json.JSONDecodeError as e:
            logger.error(f"Errore nella decodifica del file JSON: {str(e)}. "
                        f"Utilizzo valori predefiniti.")
            return self.default_config.copy()
        except Exception as e:
            logger.error(f"Errore nel caricamento della configurazione: {str(e)}. "
                        f"Utilizzo valori predefiniti.")
            return self.default_config.copy()
    
    def save_config(self, new_config):
        """
        Salva la configurazione nel file JSON.
        
        Args:
            new_config (dict): Nuova configurazione da salvare.
            
        Returns:
            bool: True se il salvataggio è riuscito, False altrimenti.
        """
        try:
            # Assicurati che la directory esista
            config_dir = os.path.dirname(self.config_file)
            if config_dir and not os.path.exists(config_dir):
                os.makedirs(config_dir, exist_ok=True)
                logger.info(f"Creata directory per la configurazione: {config_dir}")
            
            # Salva la configurazione in formato JSON
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(new_config, f, indent=4)
            
            # Aggiorna la configurazione in memoria
            self.config = new_config
            
            logger.info("Configurazione salvata con successo.")
            return True
        except Exception as e:
            logger.error(f"Errore nel salvataggio della configurazione: {str(e)}")
            return False
    
    def get_config(self):
        """
        Restituisce una copia della configurazione corrente.
        
        Returns:
            dict: Copia della configurazione corrente.
        """
        return self.config.copy()
    
    def get_save_directory(self):
        """
        Restituisce la directory di salvataggio.
        
        Returns:
            str: Percorso della directory di salvataggio.
        """
        return self.config.get("save_directory", "")
    
    def get_technologies(self):
        """
        Restituisce la configurazione delle tecnologie.
        
        Returns:
            dict: Configurazione delle tecnologie.
        """
        return self.config.get("technologies", {})
    
    def get_operations(self):
        """
        Restituisce la configurazione delle operazioni.
        
        Returns:
            dict: Configurazione delle operazioni.
        """
        return self.config.get("operations", {})
    
    def validate_config(self, config=None):
        """
        Valida la configurazione.
        
        Args:
            config (dict, optional): Configurazione da validare.
                                    Se non specificata, usa la configurazione corrente.
        
        Returns:
            tuple: (bool, str) - True se la configurazione è valida, 
                  False e messaggio di errore altrimenti.
        """
        # Usa la configurazione fornita o quella corrente
        config_to_validate = config if config is not None else self.config
        
        # Verifica che la directory di salvataggio sia specificata
        save_dir = config_to_validate.get("save_directory", "")
        if not save_dir:
            return False, "La directory di salvataggio non è specificata."
        
        # Verifica che la directory esista
        if not os.path.exists(save_dir):
            return False, f"La directory di salvataggio '{save_dir}' non esiste."
        
        # Verifica che ci sia almeno un'operazione attiva
        operations = config_to_validate.get("operations", {})
        if not any(operations.values()):
            return False, "È necessario selezionare almeno un'operazione da eseguire."
        
        # Verifica che ci sia almeno una tecnologia con prefissi
        technologies = config_to_validate.get("technologies", {})
        if not technologies:
            return False, "Non sono state configurate tecnologie."
        
        has_prefixes = False
        for tech, prefixes in technologies.items():
            if prefixes:
                has_prefixes = True
                break
        
        if not has_prefixes:
            return False, "È necessario specificare almeno un prefisso per una tecnologia."
        
        # Tutte le verifiche sono state superate
        return True, ""
    
    def reset_config(self):
        """
        Ripristina la configurazione ai valori predefiniti.
        
        Returns:
            bool: True se il ripristino è riuscito, False altrimenti.
        """
        try:
            # Ripristina la configurazione in memoria
            self.config = self.default_config.copy()
            
            # Salva la configurazione nel file
            return self.save_config(self.config)
        except Exception as e:
            logger.error(f"Errore nel ripristino della configurazione predefinita: {str(e)}")
            return False
    
    def update_config_value(self, key, value):
        """
        Aggiorna un singolo valore nella configurazione.
        
        Args:
            key (str): Chiave da aggiornare (può essere un percorso con punti, es. "operations.estrai_AdM").
            value: Valore da impostare.
        
        Returns:
            bool: True se l'aggiornamento è riuscito, False altrimenti.
        """
        try:
            # Crea una copia della configurazione corrente
            new_config = self.config.copy()
            
            # Gestisci percorsi con punti (es. "operations.estrai_AdM")
            if "." in key:
                parts = key.split(".")
                current = new_config
                
                # Naviga la struttura fino all'ultimo livello
                for part in parts[:-1]:
                    if part not in current:
                        current[part] = {}
                    current = current[part]
                
                # Imposta il valore
                current[parts[-1]] = value
            else:
                # Imposta direttamente il valore
                new_config[key] = value
            
            # Salva la configurazione aggiornata
            return self.save_config(new_config)
        except Exception as e:
            logger.error(f"Errore nell'aggiornamento del valore di configurazione '{key}': {str(e)}")
            return False
    
    def get_config_value(self, key, default=None):
        """
        Ottiene un singolo valore dalla configurazione.
        
        Args:
            key (str): Chiave da leggere (può essere un percorso con punti, es. "operations.estrai_AdM").
            default: Valore predefinito da restituire se la chiave non esiste.
        
        Returns:
            Il valore associato alla chiave o il valore predefinito se la chiave non esiste.
        """
        try:
            # Gestisci percorsi con punti (es. "operations.estrai_AdM")
            if "." in key:
                parts = key.split(".")
                current = self.config
                
                # Naviga la struttura
                for part in parts:
                    if part not in current:
                        return default
                    current = current[part]
                
                return current
            else:
                # Leggi direttamente il valore
                return self.config.get(key, default)
        except Exception as e:
            logger.error(f"Errore nella lettura del valore di configurazione '{key}': {str(e)}")
            return default