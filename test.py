import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QProgressBar, 
                            QPushButton, QVBoxLayout, QHBoxLayout, 
                            QWidget, QLabel, QStatusBar)
from PyQt5.QtCore import Qt, QTimer

class ExemploProgressBar(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Imposta la finestra principale
        self.setWindowTitle("Esempio QProgressBar")
        self.setGeometry(100, 100, 600, 400)
        
        # Widget centrale
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principale
        main_layout = QVBoxLayout(central_widget)
        
        # SEZIONE 1: Progress bar standard
        # Etichetta
        label_standard = QLabel("Progress Bar Standard:")
        main_layout.addWidget(label_standard)
        
        # Progress bar standard
        self.progress_standard = QProgressBar()
        self.progress_standard.setMinimum(0)
        self.progress_standard.setMaximum(100)
        self.progress_standard.setValue(0)
        self.progress_standard.setTextVisible(True)  # Mostra la percentuale
        main_layout.addWidget(self.progress_standard)
        
        # SEZIONE 2: Progress bar con stile personalizzato
        label_custom = QLabel("Progress Bar Personalizzata:")
        main_layout.addWidget(label_custom)
        
        # Progress bar con stile personalizzato
        self.progress_custom = QProgressBar()
        self.progress_custom.setMinimum(0)
        self.progress_custom.setMaximum(100)
        self.progress_custom.setValue(0)
        
        # Applica stile CSS personalizzato
        self.progress_custom.setStyleSheet("""
            QProgressBar {
                border: 2px solid grey;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                width: 10px;
                margin: 0.5px;
            }
        """)
        main_layout.addWidget(self.progress_custom)
        
        # SEZIONE 3: Progress bar indeterminata
        label_indeterminate = QLabel("Progress Bar Indeterminata (per operazioni con durata sconosciuta):")
        main_layout.addWidget(label_indeterminate)
        
        # Progress bar indeterminata
        self.progress_indeterminate = QProgressBar()
        self.progress_indeterminate.setMinimum(0)
        self.progress_indeterminate.setMaximum(0)  # 0 = modalità indeterminata
        main_layout.addWidget(self.progress_indeterminate)
        
        # Pulsanti di controllo
        button_layout = QHBoxLayout()
        
        # Pulsante per avviare 
        self.start_button = QPushButton("Avvia")
        self.start_button.clicked.connect(self.start_progress)
        button_layout.addWidget(self.start_button)
        
        # Pulsante per interrompere
        self.stop_button = QPushButton("Interrompi")
        self.stop_button.clicked.connect(self.stop_progress)
        self.stop_button.setEnabled(False)
        button_layout.addWidget(self.stop_button)
        
        # Pulsante per resettare
        self.reset_button = QPushButton("Reset")
        self.reset_button.clicked.connect(self.reset_progress)
        button_layout.addWidget(self.reset_button)
        
        main_layout.addLayout(button_layout)
        
        # Aggiunge spazio vuoto
        main_layout.addStretch()
        
        # Status Bar con progress bar integrata
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Pronto")
        
        # Aggiunge una progress bar alla status bar
        self.status_progress = QProgressBar()
        self.status_progress.setMaximumWidth(150)
        self.status_progress.setMaximumHeight(16)
        self.status_progress.setValue(0)
        self.statusBar.addPermanentWidget(self.status_progress)
        
        # Stile status bar
        self.statusBar.setStyleSheet("""
            QStatusBar {
                border-top: 1px solid #aaaaaa;
                background-color: #f0f0f0;
            }
            QStatusBar QProgressBar {
                border: 1px solid #aaaaaa;
                border-radius: 2px;
                background-color: #ffffff;
                max-height: 14px;
            }
            QStatusBar QProgressBar::chunk {
                background-color: #0078d7;
            }
        """)
        
        # Timer per simulare il progresso
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_progress)
        self.progress_value = 0
        
    def start_progress(self):
        self.timer.start(100)  # Aggiorna ogni 100ms
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.statusBar.showMessage("Operazione in corso...")
        
    def stop_progress(self):
        self.timer.stop()
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.statusBar.showMessage("Operazione interrotta")
        
    def reset_progress(self):
        self.timer.stop()
        self.progress_value = 0
        self.progress_standard.setValue(0)
        self.progress_custom.setValue(0)
        self.status_progress.setValue(0)
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.statusBar.showMessage("Reset completato")
        
    def update_progress(self):
        self.progress_value += 1
        if self.progress_value > 100:
            self.progress_value = 0
            
        # Aggiorna tutte le progress bar
        self.progress_standard.setValue(self.progress_value)
        self.progress_custom.setValue(self.progress_value)
        self.status_progress.setValue(self.progress_value)
        
        # Aggiorna messaggio nella status bar
        if self.progress_value == 100:
            self.statusBar.showMessage("Operazione completata")
            self.timer.stop()
            self.start_button.setEnabled(True)
            self.stop_button.setEnabled(False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExemploProgressBar()
    window.show()
    sys.exit(app.exec_())