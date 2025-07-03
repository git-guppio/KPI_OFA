#!/usr/bin/env python3
"""
PROGRAMMA DI TEST - METODI TEMPO DI ESECUZIONE
==============================================

Test per verificare i metodi di formattazione del tempo di esecuzione
e risolvere il problema dell'errore "takes 1 positional argument but 2 were given"

Autore: Assistant
Data: 2025-06-12
"""

import time
import random

# ============================================================================
# VERSIONE 1: METODO DI ISTANZA (come nell'originale)
# ============================================================================

class MainWindow:
    """Simula la classe MainWindow originale"""
    
    def __init__(self):
        self.name = "Test MainWindow"
    
    def format_execution_time(self, seconds: float) -> str:
        """
        Formatta il tempo di esecuzione in modo leggibile.
        
        Args:
            seconds: Tempo in secondi
            
        Returns:
            Stringa formattata (es. "1.23s", "123ms", "12.3μs")
        """
        if seconds >= 1.0:
            return f"{seconds:.2f}s"
        elif seconds >= 0.001:
            return f"{seconds * 1000:.1f}ms"
        elif seconds >= 0.000001:
            return f"{seconds * 1000000:.1f}μs"
        else:
            return f"{seconds * 1000000000:.0f}ns"
    
    def test_with_instance_method(self):
        """Test usando il metodo di istanza"""
        print("🔧 TEST METODO DI ISTANZA")
        print("-" * 30)
        
        test_times = [2.5, 0.5, 0.05, 0.005, 0.0005, 0.00005, 0.000005, 0.0000005]
        
        for test_time in test_times:
            try:
                formatted = self.format_execution_time(test_time)
                print(f"  {test_time:>12} s → {formatted}")
            except Exception as e:
                print(f"  {test_time:>12} s → ❌ ERRORE: {e}")
        
        # Test con tempo reale
        print(f"\n  Test con operazione reale:")
        start_time = time.perf_counter()
        
        # Simula operazione
        time.sleep(150)  # 10ms
        
        end_time = time.perf_counter()
        execution_time = end_time - start_time
        
        try:
            time_str = self.format_execution_time(execution_time)
            print(f"  Tempo operazione: {time_str}")
        except Exception as e:
            print(f"  ❌ ERRORE nel test reale: {e}")


# ============================================================================
# VERSIONE 2: METODO STATICO
# ============================================================================

class MainWindowStatic:
    """Versione con metodo statico"""
    
    def __init__(self):
        self.name = "Test MainWindow Static"
    
    @staticmethod
    def format_execution_time(seconds: float) -> str:
        """
        Metodo statico per formattare il tempo di esecuzione.
        
        Args:
            seconds: Tempo in secondi
            
        Returns:
            Stringa formattata (es. "1.23s", "123ms", "12.3μs")
        """
        if seconds >= 1.0:
            return f"{seconds:.2f}s"
        elif seconds >= 0.001:
            return f"{seconds * 1000:.1f}ms"
        elif seconds >= 0.000001:
            return f"{seconds * 1000000:.1f}μs"
        else:
            return f"{seconds * 1000000000:.0f}ns"
    
    def test_with_static_method(self):
        """Test usando il metodo statico"""
        print("\n⚡ TEST METODO STATICO")
        print("-" * 30)
        
        test_times = [2.5, 0.5, 0.05, 0.005, 0.0005, 0.00005, 0.000005, 0.0000005]
        
        for test_time in test_times:
            try:
                # Chiamata con self (funziona con metodo statico)
                formatted = self.format_execution_time(test_time)
                print(f"  {test_time:>12} s → {formatted}")
            except Exception as e:
                print(f"  {test_time:>12} s → ❌ ERRORE: {e}")
        
        # Test con tempo reale
        print(f"\n  Test con operazione reale:")
        start_time = time.perf_counter()
        
        # Simula operazione più lunga
        sum([i**2 for i in range(10000)])
        
        end_time = time.perf_counter()
        execution_time = end_time - start_time
        
        try:
            time_str = self.format_execution_time(execution_time)
            print(f"  Tempo operazione: {time_str}")
        except Exception as e:
            print(f"  ❌ ERRORE nel test reale: {e}")


# ============================================================================
# VERSIONE 3: FUNZIONE STANDALONE
# ============================================================================

def format_execution_time_standalone(seconds: float) -> str:
    """
    Funzione standalone per formattare il tempo di esecuzione.
    
    Args:
        seconds: Tempo in secondi
        
    Returns:
        Stringa formattata (es. "1.23s", "123ms", "12.3μs")
    """
    if seconds >= 1.0:
        return f"{seconds:.2f}s"
    elif seconds >= 0.001:
        return f"{seconds * 1000:.1f}ms"
    elif seconds >= 0.000001:
        return f"{seconds * 1000000:.1f}μs"
    else:
        return f"{seconds * 1000000000:.0f}ns"


def test_standalone_function():
    """Test della funzione standalone"""
    print("\n🔄 TEST FUNZIONE STANDALONE")
    print("-" * 30)
    
    test_times = [5.0, 1.5, 0.1, 0.01, 0.001, 0.0001, 0.00001, 0.000001]
    
    for test_time in test_times:
        try:
            formatted = format_execution_time_standalone(test_time)
            print(f"  {test_time:>12} s → {formatted}")
        except Exception as e:
            print(f"  {test_time:>12} s → ❌ ERRORE: {e}")


# ============================================================================
# VERSIONE 4: CLASSE UTILITY
# ============================================================================

class TimeUtils:
    """Classe utility per operazioni sui tempi"""
    
    @staticmethod
    def format_execution_time(seconds: float) -> str:
        """
        Metodo statico in classe utility per formattare il tempo.
        
        Args:
            seconds: Tempo in secondi
            
        Returns:
            Stringa formattata (es. "1.23s", "123ms", "12.3μs")
        """
        if seconds >= 1.0:
            return f"{seconds:.2f}s"
        elif seconds >= 0.001:
            return f"{seconds * 1000:.1f}ms"
        elif seconds >= 0.000001:
            return f"{seconds * 1000000:.1f}μs"
        else:
            return f"{seconds * 1000000000:.0f}ns"
    
    @staticmethod
    def format_with_details(seconds: float) -> dict:
        """
        Formatta il tempo con dettagli aggiuntivi.
        
        Args:
            seconds: Tempo in secondi
            
        Returns:
            dict: Dizionario con diverse rappresentazioni
        """
        return {
            'seconds': seconds,
            'formatted': TimeUtils.format_execution_time(seconds),
            'milliseconds': seconds * 1000,
            'microseconds': seconds * 1000000,
            'nanoseconds': seconds * 1000000000,
            'category': TimeUtils._get_time_category(seconds)
        }
    
    @staticmethod
    def _get_time_category(seconds: float) -> str:
        """Determina la categoria temporale"""
        if seconds >= 1.0:
            return "slow"
        elif seconds >= 0.1:
            return "moderate"
        elif seconds >= 0.01:
            return "fast"
        else:
            return "very_fast"


def test_utility_class():
    """Test della classe utility"""
    print("\n🛠️ TEST CLASSE UTILITY")
    print("-" * 30)
    
    test_times = [3.0, 0.5, 0.05, 0.005, 0.0005]
    
    for test_time in test_times:
        try:
            # Test metodo base
            formatted = TimeUtils.format_execution_time(test_time)
            print(f"  {test_time:>10} s → {formatted}")
            
            # Test metodo con dettagli
            details = TimeUtils.format_with_details(test_time)
            print(f"    Categoria: {details['category']}")
            
        except Exception as e:
            print(f"  {test_time:>10} s → ❌ ERRORE: {e}")


# ============================================================================
# TEST DI PERFORMANCE
# ============================================================================

def performance_test():
    """Test di performance per confrontare i metodi"""
    print("\n🏃 TEST DI PERFORMANCE")
    print("-" * 30)
    
    # Crea istanze
    main_window = MainWindow()
    main_window_static = MainWindowStatic()
    
    # Numero di iterazioni
    iterations = 100000
    test_time = 0.123456789
    
    # Test metodo di istanza
    start = time.perf_counter()
    for _ in range(iterations):
        main_window.format_execution_time(test_time)
    time_instance = time.perf_counter() - start
    
    # Test metodo statico
    start = time.perf_counter()
    for _ in range(iterations):
        main_window_static.format_execution_time(test_time)
    time_static = time.perf_counter() - start
    
    # Test funzione standalone
    start = time.perf_counter()
    for _ in range(iterations):
        format_execution_time_standalone(test_time)
    time_standalone = time.perf_counter() - start
    
    # Test classe utility
    start = time.perf_counter()
    for _ in range(iterations):
        TimeUtils.format_execution_time(test_time)
    time_utility = time.perf_counter() - start
    
    print(f"  Iterazioni: {iterations:,}")
    print(f"  Metodo istanza:  {time_instance:.4f}s")
    print(f"  Metodo statico:  {time_static:.4f}s") 
    print(f"  Funzione:        {time_standalone:.4f}s")
    print(f"  Classe utility:  {time_utility:.4f}s")


# ============================================================================
# TEST SCENARIO REALI
# ============================================================================

def test_real_scenarios():
    """Test con scenari reali di utilizzo"""
    print("\n🎯 TEST SCENARI REALI")
    print("-" * 30)
    
    main_window = MainWindowStatic()  # Usa versione statica
    
    # Scenario 1: Caricamento dati
    print("  Scenario 1: Caricamento dati")
    start_time = time.perf_counter()
    
    # Simula caricamento
    data = [random.random() for _ in range(10000)]
    processed = [x**2 for x in data]
    
    end_time = time.perf_counter()
    execution_time = end_time - start_time
    time_str = main_window.format_execution_time(execution_time)
    print(f"    Tempo caricamento: {time_str}")
    
    # Scenario 2: Elaborazione dati
    print("  Scenario 2: Elaborazione dati")
    start_time = time.perf_counter()
    
    # Simula elaborazione
    result = sum(processed) / len(processed)
    
    end_time = time.perf_counter()
    execution_time = end_time - start_time
    time_str = main_window.format_execution_time(execution_time)
    print(f"    Tempo elaborazione: {time_str}")
    
    # Scenario 3: Operazione veloce
    print("  Scenario 3: Operazione veloce")
    start_time = time.perf_counter()
    
    # Simula operazione veloce
    quick_result = len(data)
    
    end_time = time.perf_counter()
    execution_time = end_time - start_time
    time_str = main_window.format_execution_time(execution_time)
    print(f"    Tempo operazione veloce: {time_str}")


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Funzione principale di test"""
    print("🧪 PROGRAMMA DI TEST - METODI TEMPO DI ESECUZIONE")
    print("="*60)

    
    # Test metodo di istanza (quello che dava errore)
    try:
        main_window = MainWindow()
        main_window.test_with_instance_method()
    except Exception as e:
        print(f"❌ ERRORE METODO ISTANZA: {e}")
    

    
    print(f"\n✅ TEST COMPLETATI")
    print("="*60)
    print("💡 RACCOMANDAZIONE: Usa @staticmethod per risolvere l'errore!")


if __name__ == "__main__":
    main()